/**
 * GP海外販売ツール - Web App API
 * Version: 1.0.0
 */

// ========== 設定 ==========
const CONFIG_APP = {
  adminKey: 'gptools_admin_2024',
  cacheExpiry: 300
};

// ========== Web App エンドポイント ==========

function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';

  try {
    let result;

    switch (action) {
      case 'product':
        result = getProductInfo(params.url);
        break;
      case 'rate':
        result = getExchangeRate(params.currency || 'USD');
        break;
      case 'categories':
        result = getCategories();
        break;
      case 'user':
        result = getUserByKey(params.key);
        break;
      case 'cart':
        result = getCart(params.key);
        break;
      case 'orders':
        result = getOrders(params.key);
        break;
      case 'offsets':
        result = getOffsets(params.key);
        break;
      case 'admin_users':
        if (params.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = getAllUsers();
        break;
      case 'admin_orders':
        if (params.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = getAllOrders();
        break;
      default:
        result = { error: 'Unknown action', availableActions: ['product', 'rate', 'categories', 'user', 'cart', 'orders', 'offsets'] };
    }

    return jsonResponse(result);

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || '';

    let result;

    switch (action) {
      case 'addToCart':
        result = addToCart(data.userKey, data.item);
        break;
      case 'removeFromCart':
        result = removeFromCart(data.userKey, data.itemIndex);
        break;
      case 'clearCart':
        result = clearCart(data.userKey);
        break;
      case 'submitOrder':
        result = submitOrder(data.userKey);
        break;
      case 'admin_addUser':
        if (data.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = addUser(data.user);
        break;
      case 'admin_updateUser':
        if (data.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = updateUser(data.key, data.updates);
        break;
      case 'admin_updateCategory':
        if (data.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = updateCategory(data.name, data.updates);
        break;
      case 'admin_addCategory':
        if (data.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = addCategory(data.category);
        break;
      case 'admin_addOffset':
        if (data.adminKey !== CONFIG_APP.adminKey) throw new Error('Unauthorized');
        result = addOffset(data.offset);
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return jsonResponse(result);

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== 商品情報取得 ==========

function getProductInfo(url) {
  if (!url) throw new Error('URL is required');

  const html = fetchHtml_(url);
  const stock = stockState_(html);
  const nameJP = parseNameJP2_(html);
  const nameEN = nameJP ? translateCached_(nameJP) : '';
  const price = parsePrice_(html);
  const productId = parseProductId_(url, html);

  // カテゴリ自動判定
  const detectedCategory = detectCategory(url, html);
  const setCount = parseSetCount(html);

  return {
    url: url,
    productId: productId,
    nameJP: nameJP,
    nameEN: nameEN,
    price: parseInt(price) || 0,
    stock: stock,
    category: detectedCategory,
    setCount: setCount,
    fetchedAt: new Date().toISOString()
  };
}

// ========== カテゴリ自動判定 ==========

function detectCategory(url, html) {
  // 1) パンくずリストのURLパターンから判定
  const categoryPatterns = [
    { pattern: /\/usedgoods\/h010001/i, category: 'Driver' },
    { pattern: /\/usedgoods\/h010002/i, category: 'Fairway' },
    { pattern: /\/usedgoods\/h010003/i, category: 'Utility' },
    { pattern: /\/usedgoods\/h010004/i, category: 'Iron Set' },
    { pattern: /\/usedgoods\/h010005/i, category: 'single Iron' },
    { pattern: /\/usedgoods\/h010006/i, category: 'wedges' },
    { pattern: /\/usedgoods\/h010007/i, category: 'Putter' },
    { pattern: /\/usedgoods\/h010008/i, category: 'Club Set' },
  ];

  for (const p of categoryPatterns) {
    if (p.pattern.test(html) || p.pattern.test(url)) {
      // アイアンセットの場合、本数で分岐
      if (p.category === 'Iron Set') {
        const setCount = parseSetCount(html);
        if (setCount >= 8) {
          return '8 or More Irons';
        } else {
          return 'Irons 4-7';
        }
      }
      return p.category;
    }
  }

  // 2) クラブ種類テキストから判定（フォールバック）
  const clubTypePatterns = [
    { pattern: /ドライバー|driver/i, category: 'Driver' },
    { pattern: /フェアウェイ|fairway|ウッド|wood/i, category: 'Fairway' },
    { pattern: /ユーティリティ|utility|ハイブリッド|hybrid/i, category: 'Utility' },
    { pattern: /アイアンセット|iron.*set/i, category: 'Irons 4-7' },
    { pattern: /単品アイアン|single.*iron/i, category: 'single Iron' },
    { pattern: /ウェッジ|wedge/i, category: 'wedges' },
    { pattern: /パター|putter/i, category: 'Putter' },
    { pattern: /ヘッドカバー|head.*cover/i, category: 'Driver cover' },
    { pattern: /ヘッドのみ|head.*only/i, category: 'Head Only' },
  ];

  for (const p of clubTypePatterns) {
    if (p.pattern.test(html)) {
      return p.category;
    }
  }

  return 'nothing';
}

function parseSetCount(html) {
  const patterns = [
    /(\d+)\s*本セット/i,
    /クラブセット本数[^\d]*(\d+)/i,
    /本数[^\d]*(\d+)/i,
    /(\d+)\s*本/i
  ];

  for (const re of patterns) {
    const m = html.match(re);
    if (m) {
      return parseInt(m[1], 10);
    }
  }

  return 0;
}

// ========== 為替取得 ==========

function getExchangeRate(currency) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'rate_JPY_' + currency;
  const cached = cache.get(cacheKey);

  if (cached) {
    return { currency: currency, rate: parseFloat(cached), cached: true };
  }

  const apiUrl = 'https://api.exchangerate-api.com/v4/latest/JPY';
  const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  const data = JSON.parse(response.getContentText());

  const rate = data.rates[currency];
  if (!rate) throw new Error('Currency not found: ' + currency);

  const jpyPerUnit = 1 / rate;

  cache.put(cacheKey, jpyPerUnit.toString(), CONFIG_APP.cacheExpiry);

  return { currency: currency, rate: jpyPerUnit, cached: false };
}

// ========== カテゴリ取得 ==========

function getCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Categories');
  const data = sheet.getDataRange().getValues();

  const categories = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    categories.push({
      name: row[0],
      boxPrice: row[1],
      perBox: row[2],
      multiplier: row[3],
      groupWith: row[4],
      note: row[5]
    });
  }

  return { categories: categories };
}

// ========== ユーザー管理 ==========

function getUserByKey(key) {
  if (!key) throw new Error('User key is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return {
        key: data[i][0],
        name: data[i][1],
        margin: data[i][2],
        currency: data[i][3],
        status: data[i][4]
      };
    }
  }

  throw new Error('User not found');
}

function getAllUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      key: data[i][0],
      name: data[i][1],
      margin: data[i][2],
      currency: data[i][3],
      status: data[i][4]
    });
  }

  return { users: users };
}

function addUser(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');

  const key = user.key || Utilities.getUuid().substring(0, 8);
  sheet.appendRow([key, user.name, user.margin || 1.15, user.currency || 'USD', 'active']);

  let cartSheet = ss.getSheetByName('Cart_' + key);
  if (!cartSheet) {
    cartSheet = ss.insertSheet('Cart_' + key);
    cartSheet.getRange('A1:F1').setValues([['url', 'productId', 'nameEN', 'price', 'category', 'addedAt']]);
    cartSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#9900ff').setFontColor('white');
  }

  return { success: true, key: key };
}

function updateUser(key, updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      if (updates.name !== undefined) sheet.getRange(i + 1, 2).setValue(updates.name);
      if (updates.margin !== undefined) sheet.getRange(i + 1, 3).setValue(updates.margin);
      if (updates.currency !== undefined) sheet.getRange(i + 1, 4).setValue(updates.currency);
      if (updates.status !== undefined) sheet.getRange(i + 1, 5).setValue(updates.status);
      return { success: true };
    }
  }

  throw new Error('User not found');
}

// ========== カート管理 ==========

function getCart(userKey) {
  const user = getUserByKey(userKey);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cartSheet = ss.getSheetByName('Cart_' + userKey);

  if (!cartSheet) {
    return { items: [], user: user };
  }

  const data = cartSheet.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    items.push({
      index: i - 1,
      url: data[i][0],
      productId: data[i][1],
      nameEN: data[i][2],
      price: data[i][3],
      category: data[i][4],
      addedAt: data[i][5]
    });
  }

  return { items: items, user: user };
}

function addToCart(userKey, item) {
  const user = getUserByKey(userKey);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cartSheet = ss.getSheetByName('Cart_' + userKey);

  if (!cartSheet) {
    cartSheet = ss.insertSheet('Cart_' + userKey);
    cartSheet.getRange('A1:F1').setValues([['url', 'productId', 'nameEN', 'price', 'category', 'addedAt']]);
    cartSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#9900ff').setFontColor('white');
  }

  cartSheet.appendRow([
    item.url,
    item.productId,
    item.nameEN,
    item.price,
    item.category,
    new Date().toISOString()
  ]);

  return { success: true };
}

function removeFromCart(userKey, itemIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cartSheet = ss.getSheetByName('Cart_' + userKey);

  if (!cartSheet) throw new Error('Cart not found');

  const rowToDelete = itemIndex + 2;
  cartSheet.deleteRow(rowToDelete);

  return { success: true };
}

function clearCart(userKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cartSheet = ss.getSheetByName('Cart_' + userKey);

  if (!cartSheet) return { success: true };

  const lastRow = cartSheet.getLastRow();
  if (lastRow > 1) {
    cartSheet.deleteRows(2, lastRow - 1);
  }

  return { success: true };
}

// ========== 注文管理 ==========

function submitOrder(userKey) {
  const cart = getCart(userKey);
  if (cart.items.length === 0) throw new Error('Cart is empty');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');

  const shipping = calculateShipping(cart.items);

  const user = cart.user;
  const subtotal = cart.items.reduce((sum, item) => sum + (item.price * user.margin), 0);
  const totalJPY = subtotal + shipping.totalShipping;

  const rateInfo = getExchangeRate(user.currency);
  const totalFX = totalJPY / rateInfo.rate;

  const orderId = 'ORD-' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd-HHmmss');

  ordersSheet.appendRow([
    orderId,
    userKey,
    new Date().toISOString(),
    JSON.stringify(cart.items),
    totalJPY,
    totalFX.toFixed(2) + ' ' + user.currency,
    'pending'
  ]);

  clearCart(userKey);

  return {
    success: true,
    orderId: orderId,
    totalJPY: totalJPY,
    totalFX: totalFX.toFixed(2),
    currency: user.currency,
    shipping: shipping
  };
}

function getOrders(userKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Orders');
  const data = sheet.getDataRange().getValues();

  const orders = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userKey) {
      orders.push({
        orderId: data[i][0],
        orderDate: data[i][2],
        items: JSON.parse(data[i][3] || '[]'),
        totalJPY: data[i][4],
        totalFX: data[i][5],
        status: data[i][6]
      });
    }
  }

  return { orders: orders };
}

function getAllOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Orders');
  const data = sheet.getDataRange().getValues();

  const orders = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    orders.push({
      orderId: data[i][0],
      userKey: data[i][1],
      orderDate: data[i][2],
      items: JSON.parse(data[i][3] || '[]'),
      totalJPY: data[i][4],
      totalFX: data[i][5],
      status: data[i][6]
    });
  }

  return { orders: orders };
}

// ========== 送料計算 ==========

function calculateShipping(items) {
  const categories = getCategories().categories;

  const counts = {};
  items.forEach(item => {
    const cat = categories.find(c => c.name === item.category);
    if (!cat) return;

    const groupKey = cat.groupWith || cat.name;
    if (!counts[groupKey]) counts[groupKey] = 0;
    counts[groupKey] += (cat.multiplier || 1);
  });

  const breakdown = {};
  let totalShipping = 0;

  Object.keys(counts).forEach(groupKey => {
    const cat = categories.find(c => c.name === groupKey);
    if (!cat) return;

    const count = counts[groupKey];
    const boxes = Math.ceil(count / cat.perBox);
    const shipping = boxes * cat.boxPrice;

    breakdown[groupKey] = {
      count: count,
      boxes: boxes,
      pricePerBox: cat.boxPrice,
      shipping: shipping
    };

    totalShipping += shipping;
  });

  return {
    breakdown: breakdown,
    totalShipping: totalShipping
  };
}

// ========== 相殺管理 ==========

function getOffsets(userKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Offsets');
  const data = sheet.getDataRange().getValues();

  const offsets = [];
  let balance = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userKey) {
      const amount = data[i][1] || 0;
      balance += amount;
      offsets.push({
        amountJPY: amount,
        reason: data[i][2],
        createdAt: data[i][3],
        appliedTo: data[i][4]
      });
    }
  }

  return { offsets: offsets, balance: balance };
}

function addOffset(offset) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Offsets');

  sheet.appendRow([
    offset.userKey,
    offset.amountJPY,
    offset.reason,
    new Date().toISOString(),
    offset.appliedTo || ''
  ]);

  return { success: true };
}

// ========== カテゴリ管理 ==========

function updateCategory(name, updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Categories');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      if (updates.boxPrice !== undefined) sheet.getRange(i + 1, 2).setValue(updates.boxPrice);
      if (updates.perBox !== undefined) sheet.getRange(i + 1, 3).setValue(updates.perBox);
      if (updates.multiplier !== undefined) sheet.getRange(i + 1, 4).setValue(updates.multiplier);
      if (updates.groupWith !== undefined) sheet.getRange(i + 1, 5).setValue(updates.groupWith);
      if (updates.note !== undefined) sheet.getRange(i + 1, 6).setValue(updates.note);
      return { success: true };
    }
  }

  throw new Error('Category not found');
}

function addCategory(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Categories');

  sheet.appendRow([
    category.name,
    category.boxPrice || 0,
    category.perBox || 10,
    category.multiplier || 1,
    category.groupWith || '',
    category.note || ''
  ]);

  return { success: true };
}

// ========== HTML取得・解析（既存ロジック流用） ==========

function fetchHtml_(url) {
  const tsUrl = url + (url.includes('?') ? '&' : '?') + '_ts=' + Date.now();
  const resp = UrlFetchApp.fetch(tsUrl, {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (AppsScript)',
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache'
    }
  });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('HTTP ' + code);
  return bestDecode_(resp);
}

function bestDecode_(resp) {
  const utf8 = resp.getContentText();
  const sjis = resp.getContentText('Shift_JIS');
  const scoreJP = s => (s.match(/[一-龥ぁ-ゖァ-ヺ々ー]/g) || []).length;
  const uScore = scoreJP(utf8), sScore = scoreJP(sjis);
  if (/charset\s*=\s*shift[_-]?jis/i.test(sjis) || sScore > uScore) return sjis;
  if (/charset\s*=\s*utf-?8/i.test(utf8) || uScore >= sScore) return utf8;
  return utf8;
}

function stockState_(html) {
  const h = String(html || '');
  const idNoStock  = /\bid\s*=\s*(?:"|')?nostock\b/i.test(h);
  const srcNoStock = /(?:\bsrc|\bdata-src|\bsrcset)\s*=\s*["'][^"']*nostock2\.gif(?:\?[^"']*)?["']/i.test(h);
  if (idNoStock || srcNoStock) return 'OUT';
  if (/(sold\s*out|売り切れ|在庫(?:なし|切れ)|販売終了|販売停止|完売|お取り扱いできません)/i.test(h)) return 'OUT';
  if (/(在庫あり|在庫有り|残りわずか|カートに入れる|買い物かごに入れる|購入手続き|今すぐ購入|add[\s_-]*to[\s_-]*cart|buy[\s_-]*now|name=["']variation_cart["']|id=["']variation_cart1["'])/i.test(h)) return 'IN';
  return 'UNKNOWN';
}

function parsePrice_(html) {
  const ps = [
    /販売価格：<\/span>\s*([\d,]+)\s*円/i,
    /販売価格[：:\s]*([\d,]+)\s*円/i,
    /価格[：:\s]*<[^>]*>\s*([\d,]+)\s*円/i,
    /税込[価格：:\s]*([\d,]+)\s*円/i,
    /"price"\s*:\s*"([\d,]+)"/i,
    /"price":\s*([\d.]+)\s*,\s*"priceCurrency":"JPY"/i
  ];
  for (const re of ps) {
    const m = html.match(re);
    if (m) return String(m[1]).replace(/[^\d.]/g, '').replace(/\.\d+$/, '');
  }
  return '';
}

function parseProductId_(url, html) {
  const m1 = url.match(/\/g\/g(\d+)(?:\/|$)/i); if (m1) return m1[1];
  const m2 = html.match(/\/g\/g(\d+)(?:\/|["'])/i); if (m2) return m2[1];
  const m3 = html.match(/product(?:_id|Id)["']?\s*[:=]\s*["']?(\d{4,})/i); if (m3) return m3[1];
  return '';
}

function parseNameJP2_(htmlOriginal) {
  let html = String(htmlOriginal || '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '');

  html = cutAtAuxBlocks_(html);

  const buyIdx = indexOfBuyArea_(html);
  if (buyIdx >= 0) {
    const slice = html.slice(Math.max(0, buyIdx - 20000), buyIdx + 1000);
    const x = extractH2GoodsNameNear_(slice);
    if (x) return x;
  }

  const pidx = indexOfPriceBlock_(html);
  if (pidx >= 0) {
    const slice = html.slice(Math.max(0, pidx - 12000), pidx + 2000);
    const x = extractH2GoodsNameNear_(slice);
    if (x) return x;
  }

  {
    const x = extractH2GoodsNameNear_(html);
    if (x) return x;
  }

  return '';
}

function cutAtAuxBlocks_(html) {
  const markers = [
    /おすすめ商品/i, /関連商品/i, /ランキング/i, /最近見た商品/i, /この商品を見た人は/i,
    /こちらもおすすめ/i, /ピックアップ/i
  ];
  let end = html.length;
  for (const re of markers) { const i = html.search(re); if (i >= 0 && i < end) end = i; }
  return html.slice(0, end);
}

function indexOfBuyArea_(html) {
  const RE = [
    /id=["']variation_cart1["']/i,
    /name=["']variation_cart["']/i,
    /カートに入れる|購入手続き|今すぐ購入/i
  ];
  for (const re of RE) { const idx = html.search(re); if (idx >= 0) return idx; }
  return -1;
}

function indexOfPriceBlock_(html) {
  const RE = [
    /販売価格：/i,
    /税込[価格：:\s]/i,
    /"price"\s*:\s*"\d[\d,]*"/i,
    /"price":\s*\d+\s*,\s*"priceCurrency":"JPY"/i
  ];
  for (const re of RE) { const idx = html.search(re); if (idx >= 0) return idx; }
  return -1;
}

function extractH2GoodsNameNear_(htmlSlice) {
  const re = /<h2[^>]*\bclass\s*=\s*(["'])[^"']*\bgoods_name_\b[^"']*\1[^>]*>([\s\S]*?)<\/h2>/ig;
  let m, last = null;
  while ((m = re.exec(htmlSlice)) !== null) last = m;
  if (!last) return '';
  const inner = last[2].replace(/<[^>]*>/g, '');
  return inner.replace(/\s+/g, ' ').trim();
}

function decodeHtmlEntities_(s) {
  if (!s) return s;
  const map = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  return s
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCharCode(parseInt(n, 16)))
    .replace(/&([a-z]+);/gi, (m, n) => map[n.toLowerCase()] || m)
    .trim();
}

function translateCached_(ja) {
  if (!ja) return '';
  const cache = CacheService.getDocumentCache();
  const key = 'tr:' + Utilities.base64EncodeWebSafe(ja).slice(0, 100);
  const hit = cache.get(key); if (hit) return hit;
  const clean = ja.replace(/\s+/g, ' ').trim();
  const en = LanguageApp.translate(clean, 'ja', 'en');
  cache.put(key, en, 6 * 60 * 60);
  return en;
}

// ========== 初期セットアップ ==========

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
  }
  usersSheet.getRange('A1:E1').setValues([['key', 'name', 'margin', 'currency', 'status']]);
  usersSheet.getRange('A2:E2').setValues([['sample123', 'Sample User', 1.15, 'USD', 'active']]);
  usersSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  let catSheet = ss.getSheetByName('Categories');
  if (!catSheet) {
    catSheet = ss.insertSheet('Categories');
  }
  catSheet.getRange('A1:F1').setValues([['name', 'boxPrice', 'perBox', 'multiplier', 'groupWith', 'note']]);
  const categories = [
    ['Driver', 5940, 23, 1, '', ''],
    ['Driver cover', 5940, 23, 2, 'Driver', 'Driverと合算'],
    ['Fairway', 4330, 15, 1, '', ''],
    ['Utility', 4330, 15, 1, '', ''],
    ['Irons 4-7', 4330, 10, 1, '', ''],
    ['8 or More Irons', 5940, 10, 1, '', ''],
    ['single Iron', 650, 50, 1, '', ''],
    ['wedges', 800, 30, 1, '', ''],
    ['Putter', 4330, 15, 1, '', ''],
    ['Head Only', 650, 50, 1, '', ''],
    ['nothing', 0, 999, 1, '', ''],
    ['Hosel Fairway', 0, 999, 1, '', ''],
    ['over size', 9677, 5, 1, '', '']
  ];
  catSheet.getRange(2, 1, categories.length, 6).setValues(categories);
  catSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#6aa84f').setFontColor('white');

  let ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet('Orders');
  }
  ordersSheet.getRange('A1:G1').setValues([['orderId', 'userKey', 'orderDate', 'items', 'totalJPY', 'totalFX', 'status']]);
  ordersSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#e69138').setFontColor('white');

  let offsetsSheet = ss.getSheetByName('Offsets');
  if (!offsetsSheet) {
    offsetsSheet = ss.insertSheet('Offsets');
  }
  offsetsSheet.getRange('A1:E1').setValues([['userKey', 'amountJPY', 'reason', 'createdAt', 'appliedTo']]);
  offsetsSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#cc0000').setFontColor('white');

  const sheet1 = ss.getSheetByName('シート1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
}

// 権限承認用テスト関数
function testPermissions() {
  const resp = UrlFetchApp.fetch('https://api.exchangerate-api.com/v4/latest/JPY');
  console.log(resp.getContentText().substring(0, 200));

  const translated = LanguageApp.translate('テスト', 'ja', 'en');
  console.log('Translation:', translated);

  const cache = CacheService.getScriptCache();
  cache.put('test', 'ok', 60);
  console.log('Cache:', cache.get('test'));

  console.log('✅ 全権限OK');
}
