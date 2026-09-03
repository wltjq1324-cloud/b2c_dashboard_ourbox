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
//    msg += '\n품질 신호: ' + cacheStats.issueText;
//    msg += '\n소요 시간: ' + cacheStats.timingText;
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
  attachSheetMeta_(ss, cache);
  cache.meta.source = 'built_on_demand';
  cache.meta.warning = 'dashboard_cache sheet was missing; run refreshProcessedData once';
  return { dashboardCache: cache, meta: cache.meta };
}

function refreshDashboardCache(ss, processedRows) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var t0 = Date.now();
  var orders = processedRows
    ? processedRows.map(function(row, i) { return processedRowToOrder_(row, i); })
    : readProcessedOrdersForCache_(ss);
  var t1 = Date.now();

  var cache = buildDashboardCacheFromOrders_(orders);
  attachSheetMeta_(ss, cache);
  var t2 = Date.now();

  writeDashboardCache_(ss, cache);
  var t3 = Date.now();

  // 어느 단계가 느린지 alert / 실행 로그에서 바로 볼 수 있게 초 단위로 남깁니다.
  var timings = {
    readSec: Math.round((t1 - t0) / 100) / 10,
    buildSec: Math.round((t2 - t1) / 100) / 10,
    writeSec: Math.round((t3 - t2) / 100) / 10,
    totalSec: Math.round((t3 - t0) / 100) / 10
  };
  Logger.log('dashboard_cache 갱신: 읽기 %ss / 집계 %ss / 쓰기 %ss / 합계 %ss',
    timings.readSec, timings.buildSec, timings.writeSec, timings.totalSec);

  return {
    rawRows: orders.length,
    baseRows: cache.baseRows.length,
    productRows: cache.productRows.length,
    qualityRows: cache.qualityRows.length,
    issues: cache.meta.issueCounts,
    issueText: qualityCountText_(cache.meta.issueCounts),
    timings: timings,
    timingText: '캐시 ' + timings.totalSec + 's (읽기 ' + timings.readSec +
      ' / 집계 ' + timings.buildSec + ' / 쓰기 ' + timings.writeSec + ')'
  };
}

// 대시보드의 '가공_데이터 행' 칩이 정확한 탭으로 점프하도록 시트 id와 탭 gid를 실어 보냅니다.
function attachSheetMeta_(ss, cache) {
  var processed = ss.getSheetByName('가공_데이터');
  cache.meta.spreadsheetId = ss.getId();
  cache.meta.processedSheetGid = processed ? processed.getSheetId() : null;
}

function qualityCountText_(counts) {
  var labels = {
    channelUnmapped: '채널 미매핑',
    productUnmapped: '상품 미매핑',
    managerMissing: '담당자 미지정',
    costMissing: '원가 0/1',
    zeroRevenue: '매출 0원',
    negativeMargin: '마진 음수'
  };
  var out = [];
  for (var key in labels) {
    out.push(labels[key] + ' ' + ((counts && counts[key]) || 0));
  }
  return out.join(' / ');
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
    id: text_(row[0]),
    item: text_(row[1]),
    qty: num_(row[2]),
    revenue: num_(row[3]),
    ship: num_(row[4]),
    shop: text_(row[5]),
    dt: text_(row[6]),
    date: normalizeDateOnly_(row[7] || row[6]),
    channel: text_(row[8], '미매핑'),
    product: text_(row[9], '미매핑'),
    category: text_(row[10], '미매핑'),
    feeRate: num_(row[11]),
    settlement: num_(row[12]),
    cost: num_(row[13]),
    shipCost: num_(row[14]),
    margin: num_(row[15]),
    manager: text_(row[16], '미지정')
  };
}

// 가공_데이터 셀이 숫자가 아니라 "33,000" 같은 텍스트로 들어오는 경우가 있습니다.
// Number("33,000")은 NaN이라 예전 코드에서는 그대로 0이 되어
// 대시보드에 '매출 0원 품목' / '원가 0 또는 1' 품질 신호로 잡혔습니다.
function num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  if (value instanceof Date) return 0;
  var text = String(value == null ? '' : value).replace(/[,\s₩원]/g, '');
  if (!text || text === '-') return 0;
  var parsed = Number(text);
  return isFinite(parsed) ? parsed : 0;
}

// String(x || 'fallback').trim() 은 x가 ' ' 처럼 공백만 있을 때 ''를 돌려줍니다.
// 그러면 미매핑으로 분류되지 않고 빈 채널/빈 상품군이 조용히 통과합니다.
function text_(value, fallback) {
  var text = String(value == null ? '' : value).trim();
  return text || (fallback || '');
}

function buildDashboardCacheFromOrders_(orders) {
  var baseMap = {};
  var productMap = {};
  var qualityRows = [];
  var issueCounts = {
    channelUnmapped: 0,
    productUnmapped: 0,
    managerMissing: 0,
    costMissing: 0,
    zeroRevenue: 0,
    negativeMargin: 0
  };

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

    var issues = qualityIssueKeys_(order);
    if (issues.length > 0) {
      for (var k = 0; k < issues.length; k++) {
        issueCounts[issues[k]] = (issueCounts[issues[k]] || 0) + 1;
      }
      qualityRows.push(compactQualityRow_(order, issues));
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
    qualityRows: qualityRows.length,
    issueCounts: issueCounts
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

function compactQualityRow_(order, issues) {
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
    issueKeys: (issues || qualityIssueKeys_(order)).join(',')
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

  // autoResizeColumns는 셀 하나가 45,000자인 이 시트에서 특히 느리고,
  // 숨김 시트라 사람이 볼 일도 없어서 뺐습니다. 헤더 서식도 같은 이유로 생략합니다.

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
