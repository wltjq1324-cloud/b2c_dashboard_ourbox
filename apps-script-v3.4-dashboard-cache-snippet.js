// ============================================================
// 아워박스 MVP 대시보드 — Apps Script v3.4 cache patch
// ============================================================
// 적용 방법
// 1. 기존 doGet(e)를 아래 doGet(e)로 교체
// 2. 기존 refreshProcessedData() 안에서 가공_데이터 setValues 직후,
//    alert 메시지를 만들기 전에 아래 한 줄 추가
//
//    var cacheStats = refreshDashboardCache(ss, rows);
//
// 3. alert 메시지에 캐시 결과를 보고 싶으면 아래 한 줄 추가
//
//    msg += '\n요약 캐시: base ' + cacheStats.baseRows +
//      ' / product ' + cacheStats.productRows +
//      ' / quality ' + cacheStats.qualityRows + '행';
//
// 4. 이 파일의 나머지 helper 함수들을 기존 Apps Script 맨 아래에 붙여넣기
//
// 기대 효과
// - 기존: doGet()이 가공_데이터 5만+ 행 전체를 JSON으로 생성/전송
// - 변경: doGet()이 dashboard_cache 시트의 요약 JSON만 읽어서 전송
// - HTML은 dashboardCache.baseRows/productRows/qualityRows를 우선 사용
// ============================================================

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var data;

    // 디버그나 원본 상세 확인이 필요할 때만 전체 주문 JSON을 받습니다.
    if (params.mode === 'orders') {
      data = getProcessedData();
    } else {
      data = getDashboardCache();
    }

    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      error: true,
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getDashboardCache() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('dashboard_cache');

  if (sheet && sheet.getLastRow() > 1) {
    return readDashboardCacheSheet_(sheet);
  }

  // 캐시가 아직 없을 때만 느린 폴백을 탑니다.
  // 이후 refreshDashboardCache()를 한 번 실행하면 doGet은 이 경로를 타지 않습니다.
  var processed = getProcessedData();
  var cache = buildDashboardCacheFromOrders_(processed.orders || []);
  cache.meta.source = 'built_on_demand';
  cache.meta.warning = 'dashboard_cache sheet was missing; run refreshProcessedData once';
  return { dashboardCache: cache, meta: cache.meta };
}

function refreshDashboardCache(ss, processedRows) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var orders = processedRows
    ? processedRows.map(function(row, i) { return processedRowToOrder_(row, i); })
    : readProcessedOrdersForCache_(ss);

  var cache = buildDashboardCacheFromOrders_(orders);
  writeDashboardCache_(ss, cache);

  return {
    rawRows: orders.length,
    baseRows: cache.baseRows.length,
    productRows: cache.productRows.length,
    qualityRows: cache.qualityRows.length
  };
}

function readProcessedOrdersForCache_(ss) {
  var sheet = ss.getSheetByName('가공_데이터');
  if (!sheet || sheet.getLastRow() < 2) {
    var fallback = getProcessedData();
    return fallback.orders || [];
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  return data.map(function(row, i) {
    return processedRowToOrder_(row, i);
  }).filter(function(order) {
    return order.id || order.item;
  });
}

function processedRowToOrder_(row, i) {
  return {
    sheetRow: i + 2,
    id: String(row[0] || '').trim(),
    item: String(row[1] || '').trim(),
    qty: Number(row[2]) || 0,
    revenue: Number(row[3]) || 0,
    ship: Number(row[4]) || 0,
    shop: String(row[5] || '').trim(),
    dt: String(row[6] || '').trim(),
    date: normalizeDateOnly_(row[7] || row[6]),
    channel: String(row[8] || '미매핑').trim(),
    product: String(row[9] || '미매핑').trim(),
    category: String(row[10] || '미매핑').trim(),
    feeRate: Number(row[11]) || 0,
    settlement: Number(row[12]) || 0,
    cost: Number(row[13]) || 0,
    shipCost: Number(row[14]) || 0,
    margin: Number(row[15]) || 0,
    manager: String(row[16] || '미지정').trim()
  };
}

function buildDashboardCacheFromOrders_(orders) {
  var baseMap = {};
  var productMap = {};
  var qualityRows = [];

  for (var i = 0; i < orders.length; i++) {
    var order = orders[i];
    if (!order || !order.date) continue;

    addAgg_(baseMap, [
      order.date,
      order.channel,
      order.category,
      order.manager
    ], {
      date: order.date,
      channel: order.channel,
      category: order.category,
      manager: order.manager
    }, order);

    addAgg_(productMap, [
      order.date,
      order.channel,
      order.category,
      order.manager,
      order.product
    ], {
      date: order.date,
      channel: order.channel,
      category: order.category,
      manager: order.manager,
      product: order.product
    }, order);

    if (qualityIssueKeys_(order).length > 0) {
      qualityRows.push(compactQualityRow_(order));
    }
  }

  var baseRows = finalizeAgg_(baseMap);
  var productRows = finalizeAgg_(productMap);
  var meta = {
    source: 'dashboard_cache',
    version: 'v3.4',
    generatedAt: new Date().toISOString(),
    rawRows: orders.length,
    baseRows: baseRows.length,
    productRows: productRows.length,
    qualityRows: qualityRows.length
  };

  return {
    meta: meta,
    baseRows: baseRows,
    productRows: productRows,
    qualityRows: qualityRows
  };
}

function addAgg_(map, keyParts, seed, order) {
  var key = keyParts.join('\u001f');
  if (!map[key]) {
    map[key] = seed;
    map[key].qty = 0;
    map[key].revenue = 0;
    map[key].settlement = 0;
    map[key].cost = 0;
    map[key].shipCost = 0;
    map[key].margin = 0;
    map[key]._orders = {};
  }

  var row = map[key];
  row.qty += order.qty || 0;
  row.revenue += order.revenue || 0;
  row.settlement += order.settlement || 0;
  row.cost += order.cost || 0;
  row.shipCost += order.shipCost || 0;
  row.margin += order.margin || 0;
  if (order.id) row._orders[order.id] = true;
}

function finalizeAgg_(map) {
  var rows = [];
  for (var key in map) {
    var row = map[key];
    row.orders = Object.keys(row._orders || {}).length;
    delete row._orders;
    rows.push(row);
  }
  return rows.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    if ((a.channel || '') !== (b.channel || '')) return (a.channel || '').localeCompare(b.channel || '');
    if ((a.category || '') !== (b.category || '')) return (a.category || '').localeCompare(b.category || '');
    return (a.manager || '').localeCompare(b.manager || '');
  });
}

function qualityIssueKeys_(order) {
  var keys = [];
  if (order.channel === '미매핑') keys.push('channelUnmapped');
  if (order.product === '미매핑' || order.category === '미매핑') keys.push('productUnmapped');
  if (!order.manager || order.manager === '미지정') keys.push('managerMissing');
  if (order.cost <= 1 && order.revenue > 0) keys.push('costMissing');
  if (order.revenue === 0) keys.push('zeroRevenue');
  if (order.margin < 0) keys.push('negativeMargin');
  return keys;
}

function compactQualityRow_(order) {
  return {
    sheetRow: order.sheetRow,
    id: order.id,
    item: order.item,
    qty: order.qty,
    revenue: order.revenue,
    ship: order.ship,
    shop: order.shop,
    date: order.date,
    channel: order.channel,
    product: order.product,
    category: order.category,
    settlement: order.settlement,
    cost: order.cost,
    shipCost: order.shipCost,
    margin: order.margin,
    manager: order.manager,
    issueKeys: qualityIssueKeys_(order).join(',')
  };
}

function writeDashboardCache_(ss, cache) {
  var sheet = ss.getSheetByName('dashboard_cache');
  if (!sheet) sheet = ss.insertSheet('dashboard_cache');
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();

  var rows = [['section', 'part', 'json']];
  appendJsonChunks_(rows, 'meta', cache.meta);
  appendJsonChunks_(rows, 'baseRows', cache.baseRows);
  appendJsonChunks_(rows, 'productRows', cache.productRows);
  appendJsonChunks_(rows, 'qualityRows', cache.qualityRows);

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF');
  sheet.autoResizeColumns(1, 3);

  // 사용자가 직접 볼 필요 없는 캐시 시트입니다.
  try { sheet.hideSheet(); } catch (err) {}
}

function appendJsonChunks_(rows, section, value) {
  var text = JSON.stringify(value || null);
  var chunkSize = 45000;
  var part = 0;
  for (var i = 0; i < text.length; i += chunkSize) {
    rows.push([section, part, text.substring(i, i + chunkSize)]);
    part++;
  }
  if (text.length === 0) rows.push([section, 0, '']);
}

function readDashboardCacheSheet_(sheet) {
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var buckets = {};

  for (var i = 0; i < data.length; i++) {
    var section = String(data[i][0] || '').trim();
    if (!section) continue;
    if (!buckets[section]) buckets[section] = [];
    buckets[section].push({
      part: Number(data[i][1]) || 0,
      text: String(data[i][2] || '')
    });
  }

  var cache = {};
  for (var key in buckets) {
    buckets[key].sort(function(a, b) { return a.part - b.part; });
    var jsonText = buckets[key].map(function(part) { return part.text; }).join('');
    cache[key] = jsonText ? JSON.parse(jsonText) : null;
  }

  cache.meta = cache.meta || {};
  cache.baseRows = cache.baseRows || [];
  cache.productRows = cache.productRows || [];
  cache.qualityRows = cache.qualityRows || [];

  return {
    dashboardCache: cache,
    meta: cache.meta
  };
}

function normalizeDateOnly_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getFullYear() + '-' + p2_(value.getMonth() + 1) + '-' + p2_(value.getDate());
  }

  var text = String(value).trim();
  var m = text.match(/^(\d{4})[-.\/]\s*(\d{1,2})[-.\/]\s*(\d{1,2})/);
  if (m) return m[1] + '-' + p2_(m[2]) + '-' + p2_(m[3]);

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed.getFullYear() + '-' + p2_(parsed.getMonth() + 1) + '-' + p2_(parsed.getDate());
  }
  return '';
}

function p2_(value) {
  var text = String(value);
  return text.length < 2 ? '0' + text : text;
}
