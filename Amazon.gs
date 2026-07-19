function fetchAmazonListingData_(asin) {
  const normalized = normalizeAsin_(asin);
  const url = 'https://www.amazon.com/dp/' + (normalized || clean_(asin));
  if (!normalized || !/^[A-Z0-9]{10}$/.test(normalized)) {
    return { success: false, asin: clean_(asin), url: url, status: 'INVALID_ASIN', message: 'ASIN missing or invalid.' };
  }

  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': AD.AMAZON_USER_AGENT,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'no-cache'
      }
    });
    const code = response.getResponseCode();
    const html = response.getContentText() || '';
    if (code >= 400) {
      return {
        success: false,
        asin: normalized,
        url: url,
        status: 'HTTP_' + code,
        message: 'Amazon HTTP ' + code
      };
    }
    if (isAmazonRobotCheck_(html)) {
      return {
        success: false,
        asin: normalized,
        url: url,
        status: 'ROBOT_CHECK',
        message: 'Amazon returned a robot-check page.'
      };
    }
    if (/sorry.?we couldn.?t find|page you requested does not exist|dogs of amazon/i.test(html)) {
      return {
        success: false,
        asin: normalized,
        url: url,
        status: 'UNAVAILABLE',
        message: 'Amazon listing unavailable.'
      };
    }
    return parseAmazonHtml_(html, normalized, url);
  } catch (err) {
    console.error('fetchAmazonListingData_ error for ' + normalized + ': ' + err);
    return {
      success: false,
      asin: normalized,
      url: url,
      status: 'PARSER_ERROR',
      message: String(err && err.message ? err.message : err)
    };
  }
}

function isAmazonRobotCheck_(html) {
  return /api-services-support@amazon\.com|enter the characters you see|type the characters you see|robot check|opfcaptcha|validateCaptcha|errors\/validateCaptcha/i.test(html || '');
}

function parseAmazonHtml_(html, asin, url) {
  try {
    const decoded = decodeHtmlEntities_(html);
    const jsonLd = extractJsonLdProduct_(decoded);
    const text = stripHtmlToText_(decoded);
    const ranks = extractRanksFromText_(text);
    const detailRanks = extractRanksFromDetailBlocks_(decoded);
    const merged = mergeRankCandidates_(ranks.concat(detailRanks));

    const formatHint = detectFormatHint_(decoded + ' ' + text);
    const overall = pickOverallRank_(merged, formatHint);
    const categoryRanks = merged
      .filter(r => !overall || normalizeKey_(r.category) !== normalizeKey_(overall.category) || r.rank !== overall.rank)
      .filter((r, i, arr) => arr.findIndex(x => normalizeKey_(x.category) === normalizeKey_(r.category) && x.rank === r.rank) === i);

    let rating = jsonLd.rating;
    let reviewCount = jsonLd.reviewCount;
    if (!rating) rating = extractRatingFromHtml_(decoded, text);
    if (!reviewCount) reviewCount = extractReviewCountFromHtml_(decoded, text);
    let publicationDate = jsonLd.publicationDate || extractPublicationDateFromHtml_(decoded, text);

    const hasRank = !!(overall && overall.rank > 0) || categoryRanks.some(r => r.rank > 0);
    if (!hasRank) {
      return {
        success: false,
        asin: asin,
        url: url,
        overallRank: null,
        overallCategory: '',
        categoryRanks: [],
        rating: rating || null,
        reviewCount: reviewCount || null,
        publicationDate: publicationDate || null,
        status: 'MISSING_RANK',
        message: 'Amazon rank not found.'
      };
    }

    return {
      success: true,
      asin: asin,
      url: url,
      overallRank: overall ? overall.rank : null,
      overallCategory: overall ? overall.category : '',
      categoryRanks: categoryRanks.filter(r => r.rank > 0),
      rating: rating || null,
      reviewCount: reviewCount || null,
      publicationDate: publicationDate || null,
      status: 'OK'
    };
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

function extractJsonLdProduct_(html) {
  const out = { rating: null, reviewCount: null, publicationDate: null };
  const blocks = html.match(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi) || [];
  blocks.forEach(block => {
    const body = block.replace(/^[\s\S]*?>/, '').replace(/<\/script>$/i, '').trim();
    try {
      const data = JSON.parse(body);
      const items = Array.isArray(data) ? data : [data];
      items.forEach(item => walkJsonLd_(item, out));
    } catch (e) {}
  });
  return out;
}

function walkJsonLd_(node, out) {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) {
    node.forEach(n => walkJsonLd_(n, out));
    return;
  }
  const type = node['@type'];
  const types = Array.isArray(type) ? type.map(String) : [String(type || '')];
  if (types.some(t => /Product|Book/i.test(t))) {
    const ar = node.aggregateRating || {};
    if (!out.rating && ar.ratingValue != null) out.rating = number_(ar.ratingValue);
    if (!out.reviewCount) {
      const rc = ar.reviewCount != null ? ar.reviewCount : ar.ratingCount;
      if (rc != null) out.reviewCount = Math.round(number_(rc));
    }
    if (!out.publicationDate) {
      const raw = node.datePublished || node.releaseDate || node.publicationDate;
      const parsed = parseFlexibleDate_(raw);
      if (parsed) out.publicationDate = parsed;
    }
  }
  Object.keys(node).forEach(k => {
    if (k === '@context') return;
    walkJsonLd_(node[k], out);
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

function extractRanksFromDetailBlocks_(html) {
  const out = [];
  const detailMatch = html.match(/id=["'](?:productDetails_detailBullets_sections1|detailBullets_feature_div|detailBulletsWrapper_feature_div|productDetails_db_sections)["'][\s\S]{0,20000}/i);
  const chunk = detailMatch ? detailMatch[0] : html;
  const text = stripHtmlToText_(chunk);
  return out.concat(extractRanksFromText_(text));
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
  const prefer = [];
  if (/audio|audible/i.test(formatHint || '')) {
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

function extractRatingFromHtml_(html, text) {
  const patterns = [
    /"ratingValue"\s*:\s*"?([0-9.]+)"?/i,
    /([0-9]\.[0-9])\s*out of\s*5\s*stars/i,
    /averageCustomerReviews[\s\S]{0,200}?([0-9]\.[0-9])\s*out of\s*5/i
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = (html + '\n' + text).match(patterns[i]);
    if (m) {
      const n = number_(m[1]);
      if (n > 0 && n <= 5) return n;
    }
  }
  return null;
}

function extractReviewCountFromHtml_(html, text) {
  const patterns = [
    /"reviewCount"\s*:\s*"?([\d,]+)"?/i,
    /"ratingCount"\s*:\s*"?([\d,]+)"?/i,
    /([\d,]+)\s+ratings?/i,
    /([\d,]+)\s+customer\s+reviews?/i
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = (html + '\n' + text).match(patterns[i]);
    if (m) {
      const n = Math.round(number_(m[1]));
      if (n > 0) return n;
    }
  }
  return null;
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
