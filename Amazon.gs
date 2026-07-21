function fetchAmazonListingData_(asin, formatHint) {
  const normalized = normalizeAsin_(asin);
  const productUrl = 'https://www.amazon.com/dp/' + (normalized || clean_(asin));
  const urls = [
    productUrl,
    'https://www.amazon.com/gp/product/' + (normalized || clean_(asin)),
    'https://www.amazon.com/gp/aw/d/' + (normalized || clean_(asin)),
    // Fallback reader when Amazon blocks Apps Script direct fetches.
    AD.AMAZON_READER_FALLBACK_PREFIX + (normalized || clean_(asin))
  ];
  if (!normalized || !/^[A-Z0-9]{10}$/.test(normalized)) {
    return { success: false, asin: clean_(asin), url: productUrl, status: 'INVALID_ASIN', message: 'ASIN missing or invalid.' };
  }

  let lastFailure = null;
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    const isReader = url.indexOf('r.jina.ai') !== -1;
    try {
      if (i > 0) Utilities.sleep(isReader ? 1200 : 800);
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: isReader
          ? { 'Accept': 'text/plain', 'User-Agent': AD.AMAZON_USER_AGENT }
          : amazonFetchHeaders_()
      });
      const code = response.getResponseCode();
      const body = response.getContentText() || '';
      if (code >= 400) {
        lastFailure = {
          success: false,
          asin: normalized,
          url: productUrl,
          status: 'HTTP_' + code,
          message: 'Amazon HTTP ' + code
        };
        continue;
      }
      if (!isReader && isAmazonRobotCheck_(body)) {
        lastFailure = {
          success: false,
          asin: normalized,
          url: productUrl,
          status: 'ROBOT_CHECK',
          message: 'Amazon returned a robot-check page.'
        };
        continue;
      }
      if (/sorry.?we couldn.?t find|page you requested does not exist|dogs of amazon/i.test(body)) {
        lastFailure = {
          success: false,
          asin: normalized,
          url: productUrl,
          status: 'UNAVAILABLE',
          message: 'Amazon listing unavailable.'
        };
        continue;
      }
      const parsed = isReader
        ? parseAmazonReaderText_(body, normalized, productUrl, formatHint)
        : parseAmazonHtml_(body, normalized, productUrl, formatHint);
      if (parsed.success) {
        if (isReader) parsed.source = 'reader-fallback';
        return parsed;
      }
      lastFailure = parsed;
    } catch (err) {
      console.error('fetchAmazonListingData_ error for ' + normalized + ' @ ' + url + ': ' + err);
      lastFailure = {
        success: false,
        asin: normalized,
        url: productUrl,
        status: 'PARSER_ERROR',
        message: String(err && err.message ? err.message : err)
      };
    }
  }
  return lastFailure || {
    success: false,
    asin: normalized,
    url: productUrl,
    status: 'MISSING_RANK',
    message: 'Amazon rank not found.'
  };
}

function amazonFetchHeaders_() {
  return {
    'User-Agent': AD.AMAZON_USER_AGENT,
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Cache-Control': 'no-cache',
    'Pragma': 'no-cache',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1'
  };
}

function isAmazonRobotCheck_(html) {
  const h = html || '';
  if (h.length < 8000 && /captcha|robot|opfcaptcha|validateCaptcha|sorry/i.test(h)) return true;
  return /api-services-support@amazon\.com|enter the characters you see|type the characters you see|robot check|opfcaptcha|validateCaptcha|errors\/validateCaptcha/i.test(h);
}

function parseAmazonHtml_(html, asin, url, formatHint) {
  try {
    const decoded = decodeHtmlEntities_(html);
    const text = stripHtmlToText_(decoded);
    const markupRanks = extractRanksFromHtmlMarkup_(decoded);
    const textRanks = extractRanksFromText_(text);
    const detailRanks = extractRanksFromDetailBlocks_(decoded);
    const merged = mergeRankCandidates_(markupRanks.concat(textRanks).concat(detailRanks));
    const reviews = extractRatingAndReviews_(decoded, text, asin);
    return buildAmazonParseResult_(merged, {
      asin: asin,
      url: url,
      formatHint: formatHint || detectFormatHint_(decoded + ' ' + text),
      rating: reviews.rating,
      reviewCount: reviews.reviewCount,
      reviewsConfirmedZero: reviews.confirmedZero,
      publicationDate: extractPublicationDateFromHtml_(decoded, text)
    });
  } catch (err) {
    console.error('parseAmazonHtml_ error for ' + asin + ': ' + err);
    return {
      success: false,
      asin: asin,
      url: url,
      status: 'PARSER_ERROR',
      message: 'Amazon parser error: ' + (err && err.message ? err.message : err)
    };
  }
}

/** Parse plain/markdown text from reader fallback (e.g. jina.ai). */
function parseAmazonReaderText_(text, asin, url, formatHint) {
  try {
    const normalized = String(text || '')
      .replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '$1')
      .replace(/\*\s+/g, '\n');
    const ranks = mergeRankCandidates_(
      extractRanksFromText_(normalized).concat(extractRanksFromMarkdownLinks_(text))
    );
    const reviews = extractRatingAndReviews_(text, normalized, asin);
    return buildAmazonParseResult_(ranks, {
      asin: asin,
      url: url,
      formatHint: formatHint || detectFormatHint_(normalized),
      rating: reviews.rating,
      reviewCount: reviews.reviewCount,
      reviewsConfirmedZero: reviews.confirmedZero,
      publicationDate: extractPublicationDateFromHtml_('', normalized)
    });
  } catch (err) {
    console.error('parseAmazonReaderText_ error for ' + asin + ': ' + err);
    return {
      success: false,
      asin: asin,
      url: url,
      status: 'PARSER_ERROR',
      message: 'Amazon parser error: ' + (err && err.message ? err.message : err)
    };
  }
}

function extractRanksFromMarkdownLinks_(text) {
  const out = [];
  if (!text) return out;
  const re = /#\s*([\d,]+)\s*in\s*\[([^\]]+)\]/gi;
  let m;
  while ((m = re.exec(text)) !== null) {
    const rank = parseRankNumber_(m[1]);
    const category = cleanCategoryName_(m[2]);
    if (rank > 0 && category) out.push({ rank: rank, category: category });
  }
  // Overall often appears as: #50,966 in Kindle Store ([See Top 100...](...))
  const overallRe = /#\s*([\d,]+)\s*in\s*([A-Za-z][^(\[\n]{1,80}?)(?:\s*\(|\s*\[|$)/gi;
  while ((m = overallRe.exec(text)) !== null) {
    const rank = parseRankNumber_(m[1]);
    const category = cleanCategoryName_(m[2]);
    if (rank > 0 && category) out.push({ rank: rank, category: category });
  }
  return out;
}

function buildAmazonParseResult_(merged, meta) {
  const overall = pickOverallRank_(merged, meta.formatHint);
  const categoryRanks = (merged || [])
    .filter(r => !overall || normalizeKey_(r.category) !== normalizeKey_(overall.category) || r.rank !== overall.rank)
    .filter((r, i, arr) => arr.findIndex(x => normalizeKey_(x.category) === normalizeKey_(r.category) && x.rank === r.rank) === i);

  const hasRank = !!(overall && overall.rank > 0) || categoryRanks.some(r => r.rank > 0);
  // Never keep a star rating without at least one review/rating count.
  const reviewCount = meta.reviewCount != null ? Math.round(number_(meta.reviewCount)) : null;
  const rating = (reviewCount && reviewCount > 0 && meta.rating != null && number_(meta.rating) > 0)
    ? number_(meta.rating)
    : null;

  if (!hasRank) {
    return {
      success: false,
      asin: meta.asin,
      url: meta.url,
      overallRank: null,
      overallCategory: '',
      categoryRanks: [],
      rating: rating,
      reviewCount: reviewCount && reviewCount > 0 ? reviewCount : null,
      reviewsConfirmedZero: !!meta.reviewsConfirmedZero,
      publicationDate: meta.publicationDate || null,
      status: 'MISSING_RANK',
      message: 'Amazon rank not found.'
    };
  }

  return {
    success: true,
    asin: meta.asin,
    url: meta.url,
    overallRank: overall ? overall.rank : null,
    overallCategory: overall ? overall.category : '',
    categoryRanks: categoryRanks.filter(r => r.rank > 0),
    rating: rating,
    reviewCount: reviewCount && reviewCount > 0 ? reviewCount : null,
    reviewsConfirmedZero: !!meta.reviewsConfirmedZero,
    publicationDate: meta.publicationDate || null,
    status: 'OK'
  };
}

/**
 * Conservative rating/review parse.
 * Amazon pages include ratings for related products; only accept a rating when
 * review/rating count is >= 1 for this ASIN (or explicit zero-review signals).
 */
function extractRatingAndReviews_(html, text, asin) {
  const blob = String(html || '') + '\n' + String(text || '');
  const confirmedZero = /be the first to (?:review|rate)|no customer reviews|0\s+ratings?|has not received enough ratings|not yet rated/i.test(blob);

  // Prefer the product ACR block when present.
  let rating = null;
  let reviewCount = null;
  const acr = html && html.match(/id=["']averageCustomerReviews["'][\s\S]{0,2500}/i);
  if (acr) {
    const acrText = stripHtmlToText_(acr[0]);
    const rm = acrText.match(/([0-9]\.[0-9])\s*out of\s*5/i) || acr[0].match(/([0-9]\.[0-9])\s*out of\s*5/i);
    const cm = acrText.match(/([\d,]+)\s+ratings?/i) || acr[0].match(/([\d,]+)\s+ratings?/i);
    if (rm) rating = number_(rm[1]);
    if (cm) reviewCount = Math.round(number_(cm[1]));
  }

  // ASIN-scoped JSON-LD only (avoid related-product schemas).
  if ((rating == null || reviewCount == null) && html && asin) {
    const json = extractJsonLdForAsin_(html, asin);
    if (json.rating != null && rating == null) rating = json.rating;
    if (json.reviewCount != null && reviewCount == null) reviewCount = json.reviewCount;
  }

  if (confirmedZero || reviewCount === 0) {
    return { rating: null, reviewCount: null, confirmedZero: true };
  }
  if (!(reviewCount > 0) || !(rating > 0 && rating <= 5)) {
    return { rating: null, reviewCount: null, confirmedZero: false };
  }
  return { rating: rating, reviewCount: reviewCount, confirmedZero: false };
}

function extractJsonLdForAsin_(html, asin) {
  const out = { rating: null, reviewCount: null, publicationDate: null };
  const target = normalizeAsin_(asin);
  const blocks = html.match(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi) || [];
  blocks.forEach(block => {
    const body = block.replace(/^[\s\S]*?>/, '').replace(/<\/script>$/i, '').trim();
    try {
      const data = JSON.parse(body);
      const items = Array.isArray(data) ? data : [data];
      items.forEach(item => walkJsonLdForAsin_(item, out, target));
    } catch (e) {}
  });
  return out;
}

function walkJsonLdForAsin_(node, out, targetAsin) {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) {
    node.forEach(n => walkJsonLdForAsin_(n, out, targetAsin));
    return;
  }
  const type = node['@type'];
  const types = Array.isArray(type) ? type.map(String) : [String(type || '')];
  if (types.some(t => /Product|Book/i.test(t))) {
    const nodeAsin = normalizeAsin_(node.sku || node.productID || node.asin || node['@id'] || '');
    const urlAsin = normalizeAsin_(node.url || '');
    if (targetAsin && nodeAsin && nodeAsin !== targetAsin && urlAsin !== targetAsin) {
      // Skip unrelated product nodes.
    } else if (!targetAsin || !nodeAsin || nodeAsin === targetAsin || urlAsin === targetAsin || (!nodeAsin && !urlAsin)) {
      const ar = node.aggregateRating || {};
      const rc = ar.reviewCount != null ? ar.reviewCount : ar.ratingCount;
      const count = rc != null ? Math.round(number_(rc)) : null;
      if (count != null && count > 0 && ar.ratingValue != null) {
        if (!out.rating) out.rating = number_(ar.ratingValue);
        if (!out.reviewCount) out.reviewCount = count;
      }
      if (!out.publicationDate) {
        const parsed = parseFlexibleDate_(node.datePublished || node.releaseDate || node.publicationDate);
        if (parsed) out.publicationDate = parsed;
      }
    }
  }
  Object.keys(node).forEach(k => {
    if (k === '@context') return;
    walkJsonLdForAsin_(node[k], out, targetAsin);
  });
}

function extractPublicationDateFromHtml_(html, text) {
  const patterns = [
    /Publication\s*date\s*[:\u200f\u200e\s]*([A-Za-z]+ \d{1,2}, \d{4})/i,
    /Publication\s*Date\s*[:\u200f\u200e\s]*([A-Za-z]+ \d{1,2}, \d{4})/i,
    /Publisher[\s\S]{0,80}?(\d{1,2} [A-Za-z]+ \d{4})/i,
    /"datePublished"\s*:\s*"([^"]+)"/i
  ];
  const blob = html + '\n' + text;
  for (let i = 0; i < patterns.length; i++) {
    const m = blob.match(patterns[i]);
    if (!m) continue;
    const parsed = parseFlexibleDate_(m[1]);
    if (parsed) return parsed;
  }
  return null;
}

function parseFlexibleDate_(raw) {
  if (!raw) return null;
  if (Object.prototype.toString.call(raw) === '[object Date]' && isValidDate_(raw)) {
    return startOfDay_(raw);
  }
  const s = clean_(raw);
  if (!s) return null;
  const d = new Date(s);
  if (isValidDate_(d)) return startOfDay_(d);
  return null;
}

/** Parse Amazon rank markup where category names are often inside <a> tags. */
function extractRanksFromHtmlMarkup_(html) {
  const out = [];
  if (!html) return out;

  const bsr = html.match(/Best\s*Sellers?\s*Rank[\s\S]{0,400}?#\s*([\d,]+)\s*in\s*(?:<[^>]+>\s*)*([^<(\n]+)/i);
  if (bsr) {
    const rank = parseRankNumber_(bsr[1]);
    const category = cleanCategoryName_(bsr[2]);
    if (rank > 0 && category) out.push({ rank: rank, category: category });
  }

  const catRe = /#\s*([\d,]+)\s*in\s*<a[^>]*>\s*([^<]+?)\s*<\/a>/gi;
  let m;
  while ((m = catRe.exec(html)) !== null) {
    const rank = parseRankNumber_(m[1]);
    const category = cleanCategoryName_(m[2]);
    if (rank > 0 && category) out.push({ rank: rank, category: category });
  }

  // zg_hrsr list items sometimes wrap differently
  const liRe = /zg_hrsr[\s\S]{0,4000}/i;
  const liChunk = html.match(liRe);
  if (liChunk) {
    const itemRe = /#\s*([\d,]+)\s*in\s*(?:<[^>]+>\s*)*([^<#]+)/gi;
    let im;
    while ((im = itemRe.exec(liChunk[0])) !== null) {
      const rank = parseRankNumber_(im[1]);
      const category = cleanCategoryName_(stripHtmlToText_(im[2]));
      if (rank > 0 && category) out.push({ rank: rank, category: category });
    }
  }

  return out;
}

function extractRanksFromDetailBlocks_(html) {
  const out = [];
  const detailMatch = html.match(/id=["'](?:productDetails_detailBullets_sections1|detailBullets_feature_div|detailBulletsWrapper_feature_div|productDetails_db_sections)["'][\s\S]{0,20000}/i);
  const chunk = detailMatch ? detailMatch[0] : html;
  return out
    .concat(extractRanksFromHtmlMarkup_(chunk))
    .concat(extractRanksFromText_(stripHtmlToText_(chunk)));
}

function extractRanksFromText_(text) {
  const out = [];
  if (!text) return out;
  const patterns = [
    /Best\s*Sellers?\s*Rank[:\s]*#?\s*([\d,]+)\s*in\s*([^\n\(#]+)/gi,
    /#\s*([\d,]+)\s*in\s*([^\n\(#]+)/gi,
    /Amazon\s*Best\s*Sellers?\s*Rank[:\s]*#?\s*([\d,]+)\s*in\s*([^\n\(#]+)/gi
  ];
  patterns.forEach(re => {
    let m;
    while ((m = re.exec(text)) !== null) {
      const rank = parseRankNumber_(m[1]);
      const category = cleanCategoryName_(m[2]);
      if (rank > 0 && category) out.push({ rank: rank, category: category });
    }
  });
  return out;
}

function mergeRankCandidates_(list) {
  const map = new Map();
  (list || []).forEach(item => {
    if (!item || !item.rank || item.rank <= 0 || !item.category) return;
    const key = normalizeKey_(item.category);
    if (!map.has(key) || item.rank < map.get(key).rank) map.set(key, { rank: item.rank, category: item.category });
  });
  return [...map.values()];
}

function pickOverallRank_(ranks, formatHint) {
  if (!ranks || !ranks.length) return null;
  const hint = normalizeKey_(formatHint || '');
  const prefer = [];
  if (/kindle|ebook|e-book/.test(hint)) {
    prefer.push(/kindle\s*store/i);
  }
  if (/hard|paper|print/.test(hint)) {
    prefer.push(/^books$/i, /\bbooks\b/i);
  }
  if (/audio|audible/.test(hint)) {
    prefer.push(/audible|audiobooks?/i);
  }
  prefer.push(/kindle\s*store/i, /^books$/i, /\bbooks\b/i, /audible|audiobooks?/i);
  for (let i = 0; i < prefer.length; i++) {
    const hit = ranks.find(r => prefer[i].test(r.category));
    if (hit) return hit;
  }
  return ranks[0];
}

function detectFormatHint_(text) {
  if (/kindle/i.test(text)) return 'kindle';
  if (/audible|audiobook/i.test(text)) return 'audiobook';
  if (/paperback|hardcover/i.test(text)) return 'print';
  return '';
}

function parseRankNumber_(v) {
  const n = Number(String(v || '').replace(/,/g, '').replace(/[^\d]/g, ''));
  return isFinite(n) && n > 0 ? n : 0;
}

function cleanCategoryName_(v) {
  return clean_(decodeHtmlEntities_(v))
    .replace(/\s*\(.*?\)\s*$/, '')
    .replace(/\s*See Top \d+.*$/i, '')
    .replace(/\s*#\s*[\d,]+.*$/i, '')
    .replace(/\s+/g, ' ')
    .replace(/[>]+/g, '')
    .trim();
}

function stripHtmlToText_(html) {
  return decodeHtmlEntities_(String(html || ''))
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(li|tr|p|div|h\d|span)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n+/g, '\n')
    .trim();
}

function decodeHtmlEntities_(s) {
  return String(s || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16)));
}
