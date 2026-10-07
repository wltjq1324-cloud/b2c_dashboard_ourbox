// ============================================================
// 아워박스 MVP 대시보드 — Apps Script v3.8
// ============================================================
// 이 파일 전체를 Apps Script 편집기의 기존 v3.7 코드 위에 덮어쓰면 됩니다.
// 배포: 저장만으로는 /exec에 반영되지 않습니다. 배포 관리 → 기존 배포 → 새 버전.
//
// v3.3 유지:
//   - map_channel E열 "담당자" 읽기
//   - 가공_데이터 17열 "담당자" 추가
//   - JSON orders에 manager 필드 포함
//
// v3.4 유지:
//   - 기본 doGet()은 기존 orders 응답 유지
//   - 새 대시보드만 ?mode=cache로 dashboard_cache 요약 응답 사용
//   - refreshProcessedData() 실행 시 dashboard_cache 숨김 시트 함께 갱신(단일 버튼)
//
// v3.5 유지:
//   - writeDashboardCache_의 autoResizeColumns 제거, 청크 flush 제거
//   - updateMappingStatus 숫자 포맷 1회 일괄 적용
//   - LockService 동시 실행 방지, notifyUser_ alert, 행 수 임계치 경고
//
// v3.6 변경:
//   - 성능: 증분 갱신. raw_orders에 새로 붙은 행만 매핑해서 가공_데이터 끝에 추가합니다.
//           매핑 시트(map_channel/map_product/map_cost/출고배송비)가 바뀌었거나
//           raw_orders가 줄었거나 순서가 바뀐 흔적이 있으면 자동으로 전체 재생성합니다.
//   - 성능: 단계별 소요 시간을 alert와 실행 로그에 표시합니다.
//   - 성능: qualityRows에서 대시보드가 쓰지 않는 필드(ship/settlement/shipCost)를 뺐습니다.
//   - 정확성: num_() — "33,000" 같은 텍스트 숫자를 0이 아니라 33000으로 읽습니다.
//             기존 Number("33,000")은 NaN → 0 이라 '매출 0원 품목'으로 잡혔습니다.
//   - 정확성: text_() — 공백만 있는 셀을 미매핑/미지정으로 확실히 분류합니다.
//   - 정확성: map_product에 자동 추가만 되고 표준품목명이 빈 품목은 가공_데이터에도 '미매핑'으로 씁니다.
//   - 대시보드: cache.meta.issueCounts 품질 항목별 건수.
//   - 메뉴: '🧹 가공_데이터 전체 재생성' 추가. raw_orders 기존 행을 고쳤을 때 누르세요.
//
// v3.6.1 변경:
//   - cache.meta에서 spreadsheetId / processedSheetGid 제거. 대시보드가 시트 링크를 더 이상
//     만들지 않으므로 공개 JSON에 시트 ID를 실을 이유가 없습니다.
//
// v3.6.2 변경 (품질 신호를 주문 단위로):
//   - 도착보장 합배송은 한 주문이 여러 행으로 나뉘고 금액이 한 행에 몰립니다.
//     나머지 0원 행은 데이터 오류가 아니라 구조라서 'zeroRevenueSplit'(합배송 분할, 참고)으로 따로 분류합니다.
//     주문 전체가 0원인 행만 'zeroRevenue'(확인)로 남습니다.
//   - 'negativeMargin'은 행 마진이 아니라 **주문 합계 마진**이 음수일 때만 잡습니다.
//     (합배송 0원 행은 행 마진이 늘 음수라 12,000건 넘게 오탐이었습니다)
//   - 전체 재생성 사유가 '매핑 시트 변경'일 때 무엇이 바뀌었는지(행 수 변화) alert에 표시합니다.
//
// v3.7 변경 (대시보드 첫 로딩 속도):
//   - doGet(mode=cache)가 dashboard_cache 시트의 JSON 텍스트를 **파싱하지 않고 그대로 이어 붙여** 응답합니다.
//     v3.6까지는 수 MB JSON을 JSON.parse → JSON.stringify 두 번 거쳤고, 이 구간이 응답 시간의 대부분이었습니다.
//   - ?mode=cache&part=summary : meta + baseRows + productRows 만 (첫 화면용, 작고 빠름)
//     ?mode=cache&part=quality : meta + qualityRows 만 (품질 패널이 뒤에서 따로 받음)
//     ?mode=cache             : 예전과 같은 전체 응답 (구 대시보드 호환)
//   - CacheService(스크립트 캐시)에 섹션별 JSON 텍스트를 6시간 보관합니다. 캐시가 있으면 시트를 읽지 않습니다.
//     refreshDashboardCache()가 시트를 쓸 때 같이 갱신하므로 사람이 신경 쓸 것은 없습니다.
//   - 시트를 읽어야 할 때도 요청한 섹션의 행만 읽습니다(part=summary면 qualityRows 청크를 읽지 않음).
//   - meta에 servedFrom / servedMs 가 실려 옵니다. 대시보드 사이드바에서 어디가 느린지 바로 보입니다.
//
// v3.8 변경 (매핑 시트 자동 채움):
//   - 🔄 갱신 때 새 품목명·새 쇼핑몰명을 기존 매핑에서 **추론해 값까지 채웁니다.**
//     품목: 접두어([L]·[만월상회]) 제거 후 일치 / (NEA) 배수 → '표준명 * N' / A (1EA) + B (1EA) 세트 → 구성단품 나열
//     채널: 기존 쇼핑몰명을 포함하면 그 행의 채널그룹·수수료·담당자 복사
//   - 빈 칸만 채웁니다. 사람이 쓴 값은 건드리지 않습니다.
//   - 색: 초록 = 규칙으로 확정, 파랑 = 추론(검토), 노랑 = 못 채움. map_product G열 / map_channel F열 '매핑출처'.
//   - 세트·배수 원가는 기존 로직이 계산하므로 map_cost에는 **완전히 새 단품**만 빈 행으로 추가합니다.
//   - '점검_리포트' 탭: 검토 필요(파랑)·사람이 채울 것(노랑)을 최근 30일 매출 영향 순으로 정렬.
//   - 채운 키가 과거 가공_데이터에 미매핑으로 남아 있으면 이번 실행을 전체 재생성으로 전환합니다.
//   - 메뉴 '🧭 매핑 백로그 자동 채움': 쌓인 미매핑 전부에 같은 추론을 돌리고 전체 재생성.
// ============================================================

// 가공_데이터 행 수가 이 값을 넘으면 갱신 완료 메시지에 경고를 덧붙입니다.
var ROW_WARN_THRESHOLD = 150000;

// 증분 갱신 상태를 저장하는 문서 속성 키
var PROC_STATE_KEY = 'OURBOX_PROC_STATE';
var PROC_STATE_VERSION = 'v3.6';

// 집계 키 구분자 (제어문자 U+001F — 셀 값에 나올 일이 없는 문자)
var KEY_SEP = String.fromCharCode(31);

// 가공_데이터 헤더(17열). 열을 추가하면 여기와 buildProcessedRows_, processedRowToOrder_를 같이 고칩니다.
var PROC_HEADERS = [
  '주문번호', '품목명', '수량', '실결제금액', '배송비(고객)', '쇼핑몰명', '주문일시',
  '주문일', '채널그룹', '표준품목명', '상품군', '수수료율', '정산액',
  '상품원가', '출고배송비', '마진', '담당자'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 아워박스')
    .addItem('🔄 가공_데이터 갱신 (증분)', 'refreshProcessedData')
    .addItem('🧹 가공_데이터 전체 재생성', 'refreshProcessedDataFull')
    .addItem('🧭 매핑 백로그 자동 채움', 'fillMappingBacklogMenu')
    .addItem('⚡ 대시보드 캐시만 갱신', 'refreshDashboardCacheMenu')
    .addItem('📋 테스트: 데이터 확인', 'testGetData')
    .addToUi();
}

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};

    // 기존 대시보드 호환성을 위해 기본값은 기존 전체 orders 응답으로 유지합니다.
    // 새 HTML은 ?mode=cache&part=summary → ?mode=cache&part=quality 두 번에 나눠 받습니다.
    if (params.mode === 'cache') {
      return jsonOutput_(getDashboardCacheText_(params.part));
    }
    return jsonOutput_(JSON.stringify(getProcessedData()));
  } catch (err) {
    return jsonOutput_(JSON.stringify({
      error: true,
      message: err.message
    }));
  }
}

function jsonOutput_(text) {
  return ContentService.createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

// 메뉴 실행 결과를 사용자에게 표시합니다.
// UI가 없는 컨텍스트(시간 기반 트리거 등)에서는 alert가 실패하므로 무시합니다.
function notifyUser_(msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (err) {
    // UI 컨텍스트 없음 — 로그만 남깁니다.
  }
}

// ============================================================
// 가공_데이터 자동 생성/갱신
// ============================================================
function refreshProcessedData() {
  runWithLock_('가공_데이터 갱신', function() {
    return refreshProcessedDataCore_(false);
  });
}

function refreshProcessedDataFull() {
  runWithLock_('가공_데이터 전체 재생성', function() {
    return refreshProcessedDataCore_(true);
  });
}

// 쌓인 미매핑(표준품목명 빈 행, 미매핑 쇼핑몰) 전부에 추론을 돌리고 전체 재생성합니다.
// 전체 재생성 경로는 raw_orders 전체를 읽으므로 모든 쇼핑몰명이 추론 대상에 들어갑니다.
function fillMappingBacklogMenu() {
  runWithLock_('매핑 백로그 자동 채움', function() {
    return refreshProcessedDataCore_(true);
  });
}

// 동시 실행 방지: 갱신 중 재실행하면 깨진 시트가 노출될 수 있습니다.
function runWithLock_(label, fn) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) {
    notifyUser_('⚠️ 이미 갱신이 진행 중입니다. 잠시 후 다시 시도해주세요.');
    return;
  }

  try {
    notifyUser_(fn());
  } catch (err) {
    notifyUser_('❌ ' + label + ' 실패\n\n' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function refreshProcessedDataCore_(forceFull) {
  var timer = timer_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName('raw_orders');
  if (!rawSheet || rawSheet.getLastRow() < 2) {
    throw new Error('raw_orders 시트가 비어있습니다.');
  }

  var rawCount = rawSheet.getLastRow() - 1;
  var procSheet = ss.getSheetByName('가공_데이터');
  var procCount = procSheet ? Math.max(0, procSheet.getLastRow() - 1) : 0;

  // 1. 매핑 테이블 로드 (작아서 항상 전체 로드)
  var chMap = loadChannelMap(ss);
  var prMap = loadProductMap(ss);
  var costMap = loadCostMap(ss);
  var shipCost = loadShipCost(ss);
  timer.mark('매핑 로드');

  // 2. 증분 가능 여부 판단
  var state = loadProcState_();
  var mapSig = mapSignature_(chMap, prMap, costMap, shipCost);
  var mapSummary = mapSummary_(chMap, prMap, costMap, shipCost);
  var decision = decideRefreshMode_(state, {
    forceFull: !!forceFull,
    rawCount: rawCount,
    procCount: procCount,
    mapSig: mapSig,
    firstRawKey: rawRowKey_(rawSheet, 2),
    watermarkKey: (state && state.rawCount) ? rawRowKey_(rawSheet, state.rawCount + 1) : ''
  });
  var incremental = decision.mode === 'incremental';

  // 3. raw_orders 읽기 — 증분이면 새 행만
  var rawStart = incremental ? state.rawCount + 2 : 2;
  var rawRowsToRead = incremental ? rawCount - state.rawCount : rawCount;
  var rawData = rawRowsToRead > 0
    ? rawSheet.getRange(rawStart, 1, rawRowsToRead, 7).getValues()
    : [];
  timer.mark('raw 읽기');

  // 4. 새 품목 자동 추가 + 매핑상태/원가 열 업데이트
  var newItems = autoAddNewItems(ss, rawData);
  updateMappingStatus(ss, costMap);
  timer.mark('map_product 정리');

  // 5. 증분이면 기존 가공_데이터를 읽어 둡니다 (출고배송비 첫 행 판정 + 캐시 재생성 + 백로그 판정용).
  //    읽기는 쓰기보다 훨씬 싸서 이 비용은 감수합니다.
  var cacheSheet = ss.getSheetByName('dashboard_cache');
  var cacheMissing = !cacheSheet || cacheSheet.getLastRow() < 2;
  var seenOrders = {};
  var existingRows = [];
  var nothingToDo = incremental && rawData.length === 0 && !cacheMissing;
  if (incremental && procCount > 0) {
    existingRows = procSheet.getRange(2, 1, procCount, PROC_HEADERS.length).getValues();
    for (var e = 0; e < existingRows.length; e++) {
      var oid = text_(existingRows[e][0]);
      if (oid) seenOrders[oid] = true;
    }
    timer.mark('가공_데이터 읽기');
  }

  // 5b. 매핑 자동 채움 — 새 쇼핑몰명(raw 신규 행 + 과거 미매핑 행) / 표준품목명이 빈 map_product 행
  var fill = null;
  var fillError = '';
  try {
    var unknownShops = collectUnknownShops_(rawData, existingRows, chMap);
    fill = fillMappings_(ss, unknownShops, chMap, prMap, costMap);
    timer.mark('매핑 자동 채움');
  } catch (err) {
    fillError = err.message;
    Logger.log('fillMappings_ 실패: ' + err.message);
  }

  if (fill && fill.changed) {
    chMap = loadChannelMap(ss);
    prMap = loadProductMap(ss);
    costMap = loadCostMap(ss);
    updateMappingStatus(ss, costMap);
    mapSig = mapSignature_(chMap, prMap, costMap, shipCost);
    mapSummary = mapSummary_(chMap, prMap, costMap, shipCost);

    // 채운 키가 과거 가공_데이터에 미매핑으로 남아 있으면 과거 행도 다시 계산해야 합니다 → 전체 재생성
    if (incremental && hasBacklogRows_(existingRows, fill)) {
      incremental = false;
      decision = { mode: 'full', reason: 'auto_fill_backlog' };
      rawData = rawSheet.getRange(2, 1, rawCount, 7).getValues();
      existingRows = [];
      seenOrders = {};
      nothingToDo = false;
      timer.mark('raw 전체 재읽기');
    }
  }

  // 6. 매핑 → 가공 행 생성
  var newRows = buildProcessedRows_(rawData, chMap, prMap, costMap, shipCost, seenOrders);
  timer.mark('매핑 계산');

  // 7. 시트 쓰기
  if (!incremental) {
    procSheet = writeProcessedFull_(ss, procSheet, newRows);
    SpreadsheetApp.flush();
    timer.mark('가공_데이터 전체 쓰기');
  } else if (newRows.length > 0) {
    appendProcessedRows_(procSheet, procCount, newRows);
    SpreadsheetApp.flush();
    timer.mark('가공_데이터 추가 쓰기');
  }

  // 8. 새 대시보드용 요약 캐시 생성
  var allRows = incremental ? existingRows.concat(newRows) : newRows;
  var processedTotal = nothingToDo ? procCount : allRows.length;
  var cacheStats = null;
  if (!nothingToDo) {
    cacheStats = refreshDashboardCache(ss, allRows);
    timer.mark('캐시 생성');
  }

  // 9. 상태 저장
  saveProcState_({
    version: PROC_STATE_VERSION,
    rawCount: rawCount,
    processedRows: processedTotal,
    mapSig: mapSig,
    mapSummary: mapSummary,
    firstRawKey: rawRowKey_(rawSheet, 2),
    lastRawKey: rawRowKey_(rawSheet, rawCount + 1),
    updatedAt: new Date().toISOString()
  });

  // 9b. 점검 리포트 (검토 필요 / 사람이 채울 것)
  try {
    writeCheckReport_(ss, {
      mode: incremental ? '증분' : '전체 재생성',
      fill: fill,
      fillError: fillError,
      rows: nothingToDo ? existingRows : allRows,
      costMap: costMap
    });
    timer.mark('점검 리포트');
  } catch (err) {
    Logger.log('writeCheckReport_ 실패: ' + err.message);
  }

  // 10. 결과 메시지
  var msg = '✅ 가공_데이터 갱신 완료!\n\n' +
    '모드: ' + (incremental
      ? '증분 (+' + newRows.length + '행)'
      : '전체 재생성 (' + decisionReasonText_(decision.reason) +
        (decision.reason === 'map_changed' ? ' — ' + mapChangeText_(state && state.mapSummary, mapSummary) : '') + ')') + '\n' +
    '처리 행 수: ' + processedTotal + '건\n' +
    (nothingToDo ? '' : '고유 주문 수: ' + Object.keys(seenOrders).length + '건\n');

  if (cacheStats) {
    msg += '요약 캐시: base ' + cacheStats.baseRows +
      ' / product ' + cacheStats.productRows +
      ' / quality ' + cacheStats.qualityRows + '행\n' +
      '품질 신호: ' + cacheStats.issueText + '\n';
  } else {
    msg += '요약 캐시: 변경 없음(새 행 0건)\n';
  }

  msg += '소요 시간: ' + timer.text();
  msg += '\n' + fillSummaryText_(fill, fillError);

  var openItems = fill ? fill.productsOpen : newItems;
  if (openItems.length > 0) {
    var previewItems = openItems.slice(0, 15);
    msg += '\n\n⚠️ 추론하지 못한 품목 ' + openItems.length + '개 (map_product 노란 행):\n' + previewItems.join('\n');
    if (openItems.length > previewItems.length) {
      msg += '\n...외 ' + (openItems.length - previewItems.length) + '개';
    }
    msg += '\n→ 표준품목명/상품군을 입력하면 다음 갱신에서 자동으로 전체 재생성됩니다.';
  }
  if (fill && fill.channelsOpen.length > 0) {
    msg += '\n\n⚠️ 추론하지 못한 쇼핑몰 ' + fill.channelsOpen.length + '개 (map_channel 노란 행):\n' + fill.channelsOpen.slice(0, 10).join('\n');
  }
  if (fill && fill.costsAdded.length > 0) {
    msg += '\n\n💰 원가 입력 필요한 새 단품 ' + fill.costsAdded.length + '개 (map_cost 노란 행):\n' + fill.costsAdded.slice(0, 10).join('\n');
  }

  if (processedTotal > ROW_WARN_THRESHOLD) {
    msg += '\n\n⚠️ 행 수(' + processedTotal + ')가 많아 전체 재생성이 Apps Script 6분 제한에' +
      ' 근접할 수 있습니다.\n오래된 주문은 별도 시트로 분리(아카이브)를 검토하세요.';
  }

  msg += '\n\n갱신 시각: ' + new Date().toLocaleString('ko-KR');
  return msg;
}

// 증분 갱신을 해도 되는지 판단합니다. 조금이라도 의심스러우면 전체 재생성입니다.
function decideRefreshMode_(state, ctx) {
  var full = function(reason) { return { mode: 'full', reason: reason }; };
  if (ctx.forceFull) return full('forced');
  if (!state || state.version !== PROC_STATE_VERSION) return full('no_state');
  if (state.mapSig !== ctx.mapSig) return full('map_changed');
  if (ctx.rawCount < state.rawCount) return full('raw_shrunk');
  if (ctx.procCount !== state.processedRows) return full('processed_mismatch');
  if (state.firstRawKey !== ctx.firstRawKey) return full('raw_mismatch');
  if (state.lastRawKey !== ctx.watermarkKey) return full('raw_mismatch');
  return { mode: 'incremental', reason: 'ok', newRows: ctx.rawCount - state.rawCount };
}

// 매핑 시트 규모 요약. 전체 재생성 사유가 '매핑 시트 변경'일 때 무엇이 바뀌었는지 보여주는 용도입니다.
function mapSummary_(chMap, prMap, costMap, shipCost) {
  var mapped = 0;
  for (var k in prMap) {
    var p = prMap[k];
    if (p.std || p.cat || p.parts) mapped++;
  }
  return {
    channels: Object.keys(chMap).length,
    productsMapped: mapped,
    costs: Object.keys(costMap).length,
    shipCost: shipCost
  };
}

function mapChangeText_(before, after) {
  if (!before) return '이전 요약 없음';
  var parts = [];
  if (before.channels !== after.channels) parts.push('map_channel ' + before.channels + '→' + after.channels + '행');
  if (before.productsMapped !== after.productsMapped) parts.push('map_product 매핑완료 ' + before.productsMapped + '→' + after.productsMapped + '행');
  if (before.costs !== after.costs) parts.push('map_cost ' + before.costs + '→' + after.costs + '행');
  if (before.shipCost !== after.shipCost) parts.push('출고배송비 ' + before.shipCost + '→' + after.shipCost);
  return parts.length ? parts.join(', ') : '행 수는 같고 값이 바뀜';
}

function decisionReasonText_(reason) {
  var text = {
    forced: '수동 실행',
    no_state: '첫 실행 또는 버전 변경',
    map_changed: '매핑 시트 변경 감지',
    raw_shrunk: 'raw_orders 행 수 감소',
    processed_mismatch: '가공_데이터 행 수 불일치',
    raw_mismatch: 'raw_orders 순서/내용 변경 감지',
    auto_fill_backlog: '매핑 자동 채움 → 과거 미매핑 행 재계산'
  };
  return text[reason] || reason;
}

// raw_orders 특정 행의 주문번호+품목명. 증분 기준점이 그대로인지 확인하는 용도입니다.
function rawRowKey_(sheet, row) {
  if (!sheet || row < 2 || row > sheet.getLastRow()) return '';
  var v = sheet.getRange(row, 1, 1, 2).getValues()[0];
  return text_(v[0]) + KEY_SEP + text_(v[1]);
}

// 매핑 테이블 내용 서명. 표준품목명/상품군/구성단품이 모두 빈 map_product 행은
// 매핑 결과에 영향이 없으므로(어차피 미매핑) 제외합니다.
// 그래서 autoAddNewItems가 빈 행을 추가해도 서명이 바뀌지 않고, 사람이 값을 채우면 바뀝니다.
function mapSignature_(chMap, prMap, costMap, shipCost) {
  var parts = [];
  var k;
  var keys = Object.keys(chMap).sort();
  for (k = 0; k < keys.length; k++) {
    var c = chMap[keys[k]];
    parts.push('C', keys[k], c.group, c.fee, c.manager);
  }
  keys = Object.keys(prMap).sort();
  for (k = 0; k < keys.length; k++) {
    var p = prMap[keys[k]];
    if (!p.std && !p.cat && !p.parts) continue;
    parts.push('P', keys[k], p.std, p.cat, p.parts);
  }
  keys = Object.keys(costMap).sort();
  for (k = 0; k < keys.length; k++) {
    parts.push('$', keys[k], costMap[keys[k]]);
  }
  parts.push('S', shipCost);
  return fnv1a_(parts.join(KEY_SEP));
}

function fnv1a_(text) {
  var hash = 0x811c9dc5;
  for (var i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = (hash + ((hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24))) >>> 0;
  }
  return ('00000000' + hash.toString(16)).slice(-8) + ':' + text.length;
}

function loadProcState_() {
  try {
    var text = PropertiesService.getDocumentProperties().getProperty(PROC_STATE_KEY);
    return text ? JSON.parse(text) : null;
  } catch (err) {
    return null;
  }
}

function saveProcState_(state) {
  PropertiesService.getDocumentProperties().setProperty(PROC_STATE_KEY, JSON.stringify(state));
}

// raw 행 → 가공_데이터 행. seenOrders는 호출자가 넘긴 객체를 갱신합니다
// (출고배송비는 주문번호당 첫 행에만 붙습니다).
function buildProcessedRows_(rawData, chMap, prMap, costMap, shipCost, seenOrders) {
  var rows = [];
  for (var i = 0; i < rawData.length; i++) {
    var raw = rawData[i];
    var orderId = text_(raw[0]);
    var itemName = text_(raw[1]);
    if (!orderId && !itemName) continue;

    var qty = num_(raw[2]);
    var revenue = num_(raw[3]);
    var shipCustomer = num_(raw[4]);
    var shopName = text_(raw[5]);
    var dtStr = parseAnyDate(raw[6]);
    var dateStr = dtStr ? dtStr.substring(0, 10) : '';

    var chInfo = chMap[shopName] || { group: '미매핑', fee: 0, manager: '미지정' };
    var prInfo = prMap[itemName] || { std: '미매핑', cat: '미매핑', parts: '' };
    var settlement = Math.round(revenue * (1 - chInfo.fee));
    var unitCost = calcProductCost(prInfo.std, prInfo.parts, costMap, qty);

    var ship = 0;
    if (orderId && !seenOrders[orderId]) {
      seenOrders[orderId] = true;
      ship = shipCost;
    }

    rows.push([
      orderId, itemName, qty, revenue, shipCustomer, shopName, dtStr, dateStr,
      chInfo.group || '미매핑', prInfo.std || '미매핑', prInfo.cat || '미매핑',
      chInfo.fee, settlement, unitCost, ship,
      settlement - (unitCost + ship),
      chInfo.manager || '미지정'
    ]);
  }
  return rows;
}

function writeProcessedFull_(ss, procSheet, rows) {
  if (!procSheet) {
    procSheet = ss.insertSheet('가공_데이터');
  } else {
    if (procSheet.getFilter()) procSheet.getFilter().remove();
    procSheet.clear();
  }

  procSheet.getRange(1, 1, 1, PROC_HEADERS.length)
    .setValues([PROC_HEADERS])
    .setFontWeight('bold')
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF');

  if (rows.length > 0) {
    setValuesChunked_(procSheet, 2, 1, rows, 4000);
    applyProcessedFormats_(procSheet, 2, rows.length);
    procSheet.getRange(1, 1, rows.length + 1, PROC_HEADERS.length).createFilter();
  }
  return procSheet;
}

function appendProcessedRows_(procSheet, procCount, rows) {
  if (rows.length === 0) return;
  var startRow = procCount + 2;
  setValuesChunked_(procSheet, startRow, 1, rows, 4000);
  applyProcessedFormats_(procSheet, startRow, rows.length);

  // 필터 범위는 자동으로 늘어나지 않아서 새 행까지 포함해 다시 만듭니다.
  if (procSheet.getFilter()) procSheet.getFilter().remove();
  procSheet.getRange(1, 1, procCount + rows.length + 1, PROC_HEADERS.length).createFilter();
}

function applyProcessedFormats_(sheet, startRow, count) {
  sheet.getRange(startRow, 4, count, 1).setNumberFormat('#,##0');
  sheet.getRange(startRow, 12, count, 1).setNumberFormat('0.0%');
  sheet.getRange(startRow, 13, count, 4).setNumberFormat('#,##0');
}

function timer_() {
  var start = Date.now();
  var last = start;
  var marks = [];
  return {
    mark: function(name) {
      var now = Date.now();
      marks.push(name + ' ' + ((now - last) / 1000).toFixed(1) + 's');
      last = now;
    },
    total: function() { return (Date.now() - start) / 1000; },
    text: function() {
      return marks.join(' / ') + ' / 합계 ' + this.total().toFixed(1) + 's';
    }
  };
}

// 셀 값이 숫자가 아니라 "33,000" 같은 텍스트로 들어오는 경우가 있습니다.
// Number("33,000")은 NaN이라 예전 코드에서는 0이 되어 '매출 0원 품목'으로 잡혔습니다.
function num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  if (value instanceof Date) return 0;
  var text = String(value == null ? '' : value).replace(/[,\s₩원]/g, '');
  if (!text || text === '-') return 0;
  // "15%" 같은 텍스트 비율 → 0.15 (셀이 텍스트로 들어온 경우 대비)
  if (/%$/.test(text)) {
    var pct = Number(text.slice(0, -1));
    return isFinite(pct) ? pct / 100 : 0;
  }
  var parsed = Number(text);
  return isFinite(parsed) ? parsed : 0;
}

// String(x || 'fallback').trim() 은 x가 ' ' 처럼 공백만 있을 때 ''를 돌려줍니다.
// 그러면 미매핑으로 분류되지 않고 빈 채널/빈 상품군이 조용히 통과합니다.
function text_(value, fallback) {
  var text = String(value == null ? '' : value).trim();
  return text || (fallback || '');
}

// ============================================================
// 새 품목 자동 추가
// ============================================================
function autoAddNewItems(ss, rawData) {
  var prSheet = ss.getSheetByName('map_product');
  if (!prSheet) return [];

  var lastRow = prSheet.getLastRow();
  var existingItems = {};
  if (lastRow > 1) {
    var existing = prSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      var name = String(existing[i][0]).trim();
      if (name) existingItems[name] = true;
    }
  }

  var newItems = [];
  var seen = {};
  for (var j = 0; j < rawData.length; j++) {
    var itemName = String(rawData[j][1]).trim();
    if (!itemName || itemName === '-' || existingItems[itemName] || seen[itemName]) continue;
    seen[itemName] = true;
    newItems.push(itemName);
  }

  if (newItems.length > 0) {
    var newRows = newItems.map(function(name) {
      return [name, '', '', 'Y', ''];
    });
    prSheet.getRange(lastRow + 1, 1, newRows.length, 5).setValues(newRows);
    prSheet.getRange(lastRow + 1, 1, newRows.length, 5).setBackground('#FEF3C7');
  }

  return newItems;
}

// ============================================================
// 매핑상태/원가 열 업데이트
// ============================================================
function updateMappingStatus(ss, costMap) {
  var prSheet = ss.getSheetByName('map_product');
  if (!prSheet || prSheet.getLastRow() < 2) return;

  var lastRow = prSheet.getLastRow();
  prSheet.getRange('F1')
    .setValue('매핑상태/원가')
    .setFontWeight('bold')
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF');

  var data = prSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var statusValues = [];

  for (var i = 0; i < data.length; i++) {
    var origName = String(data[i][0]).trim();
    var stdName = String(data[i][1]).trim();
    var parts = data[i][4] ? String(data[i][4]).trim() : '';

    if (!origName || origName === '-') {
      statusValues.push(['']);
      continue;
    }
    if (!stdName) {
      statusValues.push(['⚠️ 미매핑']);
      continue;
    }

    var cost = 0;
    if (parts && parts.length > 0) {
      var partList = parts.split(',');
      var allFound = true;
      for (var p = 0; p < partList.length; p++) {
        var partName = partList[p].trim();
        if (costMap[partName]) {
          cost += costMap[partName];
        } else {
          allFound = false;
        }
      }
      if (!allFound) {
        statusValues.push(['❌ 원가미등록 (구성단품)']);
        continue;
      }
    } else {
      cost = costMap[stdName] || 0;
    }

    if (cost > 0) {
      statusValues.push([cost]);
    } else {
      if (stdName.indexOf('*') >= 0) {
        var baseName = stdName.split('*')[0].trim();
        var baseCost = costMap[baseName] || 0;
        if (baseCost > 0) {
          var multiplier = parseInt(stdName.split('*')[1], 10) || 1;
          statusValues.push([baseCost * multiplier]);
          continue;
        }
      }
      statusValues.push(['❌ 원가미등록']);
    }
  }

  // 값 일괄 쓰기
  prSheet.getRange(2, 6, statusValues.length, 1).setValues(statusValues);

  // 숫자 포맷도 1회 일괄 적용 (행별 setNumberFormat 루프 제거 — 성능)
  var numberFormats = statusValues.map(function(row) {
    return [typeof row[0] === 'number' ? '#,##0' : '@'];
  });
  prSheet.getRange(2, 6, numberFormats.length, 1).setNumberFormats(numberFormats);
}

// ============================================================
// 데이터 읽기: 기존 웹앱용 orders 응답
// ============================================================
function getProcessedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('가공_데이터');
  if (sheet && sheet.getLastRow() > 1) {
    var result = readProcessedSheet(sheet);
    result.unmapped = getUnmappedInfo(ss);
    return result;
  }

  var fallback = buildFromRawFallback(ss);
  fallback.unmapped = getUnmappedInfo(ss);
  return fallback;
}

function getUnmappedInfo(ss) {
  var prSheet = ss.getSheetByName('map_product');
  if (!prSheet || prSheet.getLastRow() < 2) return [];

  var data = prSheet.getRange(2, 1, prSheet.getLastRow() - 1, 6).getValues();
  var unmapped = [];

  for (var i = 0; i < data.length; i++) {
    var orig = String(data[i][0]).trim();
    var std = String(data[i][1]).trim();
    var status = String(data[i][5]).trim();
    if (!orig || orig === '-') continue;
    if (!std || status.indexOf('⚠️') >= 0 || status.indexOf('❌') >= 0) {
      unmapped.push({ item: orig, stdName: std, status: status || '미확인' });
    }
  }

  return unmapped;
}

function setValuesChunked_(sheet, startRow, startCol, values, chunkSize, flushEachChunk) {
  if (!values || values.length === 0) return;

  chunkSize = chunkSize || 4000;
  for (var offset = 0; offset < values.length; offset += chunkSize) {
    var chunk = values.slice(offset, offset + chunkSize);
    sheet.getRange(startRow + offset, startCol, chunk.length, chunk[0].length).setValues(chunk);
    // 청크마다의 flush는 비싼 연산이라 기본적으로 호출하지 않습니다.
    // 호출부에서 종료 후 1회만 flush합니다.
    if (flushEachChunk) SpreadsheetApp.flush();
  }
}

function readProcessedSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 17);
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var orders = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;

    orders.push({
      id: text_(row[0]),
      item: text_(row[1]),
      qty: num_(row[2]),
      revenue: num_(row[3]),
      ship: num_(row[4]),
      shop: text_(row[5]),
      dt: String(row[6]),
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
    });
  }

  return {
    orders: orders,
    meta: {
      source: '가공_데이터',
      rows: orders.length,
      timestamp: new Date().toISOString(),
      version: 'v3.8'
    }
  };
}

function buildFromRawFallback(ss) {
  var rawSheet = ss.getSheetByName('raw_orders');
  if (!rawSheet || rawSheet.getLastRow() < 2) {
    return { orders: [], meta: { source: 'raw_orders', rows: 0 } };
  }

  var rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 7).getValues();
  var rows = buildProcessedRows_(
    rawData, loadChannelMap(ss), loadProductMap(ss), loadCostMap(ss), loadShipCost(ss), {}
  );
  var orders = rows.map(function(row, i) { return processedRowToOrder_(row, i); });

  return {
    orders: orders,
    meta: {
      source: 'raw_orders+매핑(폴백)',
      rows: orders.length,
      timestamp: new Date().toISOString(),
      version: 'v3.8'
    }
  };
}

// ============================================================
// 대시보드 요약 캐시
// ============================================================
function refreshDashboardCacheMenu() {
  runWithLock_('대시보드 캐시 갱신', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stats = refreshDashboardCache(ss);
    return '✅ 대시보드 캐시 갱신 완료\n\n' +
      '원본 행 수: ' + stats.rawRows + '건\n' +
      'baseRows: ' + stats.baseRows + '행\n' +
      'productRows: ' + stats.productRows + '행\n' +
      'qualityRows: ' + stats.qualityRows + '행\n' +
      '품질 신호: ' + stats.issueText + '\n' +
      '소요 시간: ' + stats.timingText + '\n\n' +
      '갱신 시각: ' + new Date().toLocaleString('ko-KR');
  });
}

// 객체가 필요한 내부 용도(testGetData 등). 웹 응답은 getDashboardCacheText_를 씁니다.
function getDashboardCache() {
  return JSON.parse(getDashboardCacheText_('all'));
}

// 캐시 섹션 이름과 시트/스크립트 캐시에 저장되는 순서
var CACHE_SECTIONS = ['meta', 'baseRows', 'productRows', 'qualityRows'];

// part → 응답에 실을 섹션. 알 수 없는 값은 전체 응답(구 대시보드 호환).
function cacheSectionsForPart_(part) {
  if (part === 'summary') return ['meta', 'baseRows', 'productRows'];
  if (part === 'quality') return ['meta', 'qualityRows'];
  if (part === 'meta') return ['meta'];
  return CACHE_SECTIONS.slice();
}

// 요청한 섹션의 JSON 텍스트를 구해 응답 문자열을 이어 붙입니다.
// 큰 배열을 JSON.parse/JSON.stringify 하지 않는 것이 핵심입니다.
function getDashboardCacheText_(part) {
  var t0 = Date.now();
  var sections = cacheSectionsForPart_(part);
  var servedFrom = 'script_cache';
  var texts = readCacheTextsFromScriptCache_(sections);

  if (!texts) {
    servedFrom = 'sheet';
    texts = readCacheTextsFromSheet_(sections);
  }

  if (!texts) {
    // dashboard_cache 시트가 없음 → 즉석 생성 (느림. refreshProcessedData 1회 실행이 정답)
    servedFrom = 'built_on_demand';
    var processed = getProcessedData();
    var built = buildDashboardCacheFromOrders_(processed.orders || []);
    built.meta.source = 'built_on_demand';
    built.meta.warning = 'dashboard_cache sheet was missing; run refreshProcessedData once';
    texts = {};
    for (var b = 0; b < sections.length; b++) {
      texts[sections[b]] = JSON.stringify(built[sections[b]] || null);
    }
  }

  // meta만 작으니 파싱해서 서빙 정보를 덧붙입니다.
  var meta = {};
  try { meta = JSON.parse(texts.meta || '{}') || {}; } catch (err) { meta = {}; }
  meta.part = part || 'all';
  meta.servedFrom = servedFrom;
  meta.servedAt = new Date().toISOString();
  meta.servedMs = Date.now() - t0;
  var metaText = JSON.stringify(meta);

  var inner = ['"meta":' + metaText];
  for (var i = 0; i < sections.length; i++) {
    var name = sections[i];
    if (name === 'meta') continue;
    inner.push('"' + name + '":' + (texts[name] || '[]'));
  }

  return '{"dashboardCache":{' + inner.join(',') + '},"meta":' + metaText + '}';
}

// ---------- 스크립트 캐시 (CacheService) ----------
// 섹션별로 'dc:<section>' 인덱스 키(생성 시각·조각 수)와 'dc:<section>:<gen>:<i>' 조각 키를 둡니다.
// 조각을 먼저 쓰고 인덱스를 마지막에 써서, 읽는 쪽이 반쯤 갱신된 상태를 보지 않게 합니다.
var SCRIPT_CACHE_TTL_SEC = 21600;     // 6시간 (CacheService 최대)
var SCRIPT_CACHE_CHUNK_CHARS = 30000; // 한글 3바이트 기준 90KB < 100KB 제한

function scriptCache_() {
  try { return CacheService.getScriptCache(); } catch (err) { return null; }
}

function readCacheTextsFromScriptCache_(sections) {
  var cache = scriptCache_();
  if (!cache) return null;

  try {
    var indexKeys = sections.map(function(s) { return 'dc:' + s; });
    var indexes = cache.getAll(indexKeys);
    var chunkKeys = [];
    var plan = {};

    for (var i = 0; i < sections.length; i++) {
      var raw = indexes['dc:' + sections[i]];
      if (!raw) return null;
      var idx = JSON.parse(raw);
      var keys = [];
      for (var p = 0; p < idx.parts; p++) keys.push('dc:' + sections[i] + ':' + idx.gen + ':' + p);
      plan[sections[i]] = keys;
      chunkKeys = chunkKeys.concat(keys);
    }

    var chunks = {};
    for (var off = 0; off < chunkKeys.length; off += 100) {
      var got = cache.getAll(chunkKeys.slice(off, off + 100));
      for (var k in got) chunks[k] = got[k];
    }

    var texts = {};
    for (var s = 0; s < sections.length; s++) {
      var parts = plan[sections[s]];
      var buf = [];
      for (var c = 0; c < parts.length; c++) {
        if (chunks[parts[c]] === undefined || chunks[parts[c]] === null) return null; // 일부 만료 → 시트로
        buf.push(chunks[parts[c]]);
      }
      texts[sections[s]] = buf.join('');
    }
    return texts;
  } catch (err) {
    Logger.log('script cache read skipped: ' + err.message);
    return null;
  }
}

function writeCacheTextsToScriptCache_(texts) {
  var cache = scriptCache_();
  if (!cache) return;

  try {
    var gen = String(Date.now());
    for (var name in texts) {
      var text = texts[name] || '';
      var parts = Math.max(1, Math.ceil(text.length / SCRIPT_CACHE_CHUNK_CHARS));
      var batch = {};
      var count = 0;
      for (var p = 0; p < parts; p++) {
        batch['dc:' + name + ':' + gen + ':' + p] =
          text.substring(p * SCRIPT_CACHE_CHUNK_CHARS, (p + 1) * SCRIPT_CACHE_CHUNK_CHARS);
        count++;
        if (count >= 50) { cache.putAll(batch, SCRIPT_CACHE_TTL_SEC); batch = {}; count = 0; }
      }
      if (count > 0) cache.putAll(batch, SCRIPT_CACHE_TTL_SEC);
      cache.put('dc:' + name, JSON.stringify({ gen: gen, parts: parts, chars: text.length }), SCRIPT_CACHE_TTL_SEC);
    }
  } catch (err) {
    // 스크립트 캐시는 가속용일 뿐입니다. 실패해도 시트 경로가 그대로 동작합니다.
    Logger.log('script cache write skipped: ' + err.message);
  }
}

// ---------- 시트 읽기 (요청한 섹션의 행만) ----------
function readCacheTextsFromSheet_(sections) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('dashboard_cache');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var lastRow = sheet.getLastRow();
  // A:B(section, part)만 먼저 읽습니다. json 열은 필요한 행 범위만 읽습니다.
  var head = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var wanted = {};
  for (var w = 0; w < sections.length; w++) wanted[sections[w]] = { first: -1, last: -1 };

  for (var i = 0; i < head.length; i++) {
    var name = String(head[i][0] || '').trim();
    if (!wanted[name]) continue;
    if (wanted[name].first < 0) wanted[name].first = i;
    wanted[name].last = i;
  }

  var texts = {};
  var filled = {};
  for (var s = 0; s < sections.length; s++) {
    var range = wanted[sections[s]];
    if (range.first < 0) { texts[sections[s]] = ''; continue; }
    var rows = sheet.getRange(range.first + 2, 1, range.last - range.first + 1, 3).getValues();
    var pieces = [];
    for (var r = 0; r < rows.length; r++) {
      if (String(rows[r][0] || '').trim() !== sections[s]) continue;
      pieces.push({ part: Number(rows[r][1]) || 0, text: String(rows[r][2] || '') });
    }
    pieces.sort(function(a, b) { return a.part - b.part; });
    texts[sections[s]] = pieces.map(function(p) { return p.text; }).join('');
    filled[sections[s]] = texts[sections[s]];
  }

  // 다음 요청부터는 시트를 읽지 않도록 스크립트 캐시를 채워 둡니다.
  writeCacheTextsToScriptCache_(filled);
  return texts;
}

function refreshDashboardCache(ss, processedRows) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var t0 = Date.now();
  var orders = processedRows
    ? processedRows.map(function(row, i) { return processedRowToOrder_(row, i); })
    : readProcessedOrdersForCache_(ss);
  var t1 = Date.now();

  var cache = buildDashboardCacheFromOrders_(orders);
  var t2 = Date.now();

  writeDashboardCache_(ss, cache);
  var t3 = Date.now();

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

function qualityCountText_(counts) {
  var labels = {
    channelUnmapped: '채널 미매핑',
    productUnmapped: '상품 미매핑',
    managerMissing: '담당자 미지정',
    costMissing: '원가 0/1',
    zeroRevenue: '매출 0원(주문 전체)',
    negativeMargin: '마진 음수(주문 단위)',
    zeroRevenueSplit: '합배송 분할 0원(참고)'
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

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PROC_HEADERS.length).getValues();
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
    zeroRevenueSplit: 0,
    negativeMargin: 0
  };

  // 주문 단위 합계 (합배송 분할 행 판정용): 같은 주문번호의 매출·마진을 먼저 모읍니다.
  var orderTotals = {};
  for (var t = 0; t < orders.length; t++) {
    var o = orders[t];
    if (!o || !o.date || !o.id) continue;
    if (!orderTotals[o.id]) orderTotals[o.id] = { revenue: 0, margin: 0 };
    orderTotals[o.id].revenue += o.revenue || 0;
    orderTotals[o.id].margin += o.margin || 0;
  }

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

    var issues = qualityIssueKeys_(order, orderTotals);
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
    version: 'v3.8',
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
  var key = keyParts.join(KEY_SEP);
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
    if ((a.channel || '') !== (b.channel || '')) {
      return (a.channel || '').localeCompare(b.channel || '');
    }
    if ((a.category || '') !== (b.category || '')) {
      return (a.category || '').localeCompare(b.category || '');
    }
    return (a.manager || '').localeCompare(b.manager || '');
  });
}

// orderTotals가 있으면 매출 0원 / 마진 음수를 주문 단위로 판정합니다.
// - 행은 0원인데 같은 주문의 다른 행에 매출이 있으면 → zeroRevenueSplit (합배송 분할, 참고용)
// - 주문 전체가 0원이면 → zeroRevenue (확인 필요)
// - 마진 음수는 행이 아니라 주문 합계 마진이 음수일 때만
function qualityIssueKeys_(order, orderTotals) {
  var keys = [];
  var total = (orderTotals && order.id && orderTotals[order.id]) || null;
  var orderRevenue = total ? total.revenue : order.revenue;
  var orderMargin = total ? total.margin : order.margin;

  if (order.channel === '미매핑') keys.push('channelUnmapped');
  if (order.product === '미매핑' || order.category === '미매핑') keys.push('productUnmapped');
  if (!order.manager || order.manager === '미지정') keys.push('managerMissing');
  if (order.cost <= 1 && order.revenue > 0) keys.push('costMissing');
  if (order.revenue === 0) keys.push(orderRevenue > 0 ? 'zeroRevenueSplit' : 'zeroRevenue');
  if (order.margin < 0 && orderMargin < 0) keys.push('negativeMargin');
  return keys;
}

// 대시보드 품질 패널이 실제로 쓰는 필드만 담습니다. (ship/settlement/shipCost 제외 → 전송량 축소)
function compactQualityRow_(order, issues) {
  return {
    sheetRow: order.sheetRow,
    id: order.id,
    item: order.item,
    qty: order.qty,
    revenue: order.revenue,
    shop: order.shop,
    date: order.date,
    channel: order.channel,
    product: order.product,
    category: order.category,
    cost: order.cost,
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

  var texts = {};
  var rows = [['section', 'part', 'json']];
  for (var s = 0; s < CACHE_SECTIONS.length; s++) {
    var name = CACHE_SECTIONS[s];
    texts[name] = JSON.stringify(cache[name] || null);
    appendJsonChunks_(rows, name, texts[name]);
  }

  setValuesChunked_(sheet, 1, 1, rows, 1000);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF');

  // autoResizeColumns는 json 열(셀당 최대 45,000자)의 폭을 측정하느라
  // 수 분이 걸리므로 사용하지 않고 고정 폭을 지정합니다.
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 400);

  try {
    sheet.hideSheet();
  } catch (err) {
    // 숨김 실패는 대시보드 기능에 영향이 없어 무시합니다.
  }

  // 웹 요청이 시트 대신 바로 읽도록 스크립트 캐시도 같은 내용으로 갱신합니다.
  writeCacheTextsToScriptCache_(texts);
}

function appendJsonChunks_(rows, section, text) {
  text = String(text || '');
  var chunkSize = 45000;
  var part = 0;

  for (var i = 0; i < text.length; i += chunkSize) {
    rows.push([section, part, text.substring(i, i + chunkSize)]);
    part++;
  }

  if (text.length === 0) {
    rows.push([section, 0, '']);
  }
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

// ============================================================
// v3.8 매핑 자동 채움
// ============================================================
var FILL_COLOR_AUTO = '#DCFCE7';   // 규칙으로 확정
var FILL_COLOR_INFER = '#DBEAFE';  // 추론 — 검토 필요
var FILL_COLOR_OPEN = '#FEF3C7';   // 못 채움 — 사람이 입력
var FILL_COLOR_WARN = '#FEE2E2';   // 원가 0/1 등 위험
var REPORT_SHEET_NAME = '점검_리포트';
var REPORT_RECENT_DAYS = 30;

// 품목명 정규화: 선행 [접두어] 제거, (N개입) → N개입, 끝의 (1EA) 제거, 공백 제거, 소문자
function normItem_(s) {
  s = String(s || '').toLowerCase();
  var prev;
  do { prev = s; s = s.replace(/^\s*\[[^\]]*\]\s*/, ''); } while (s !== prev);
  s = s.replace(/\(\s*(\d+)\s*개입\s*\)/g, ' $1개입');
  s = s.replace(/\(\s*1\s*ea\s*\)\s*$/i, '');
  s = s.replace(/\s+/g, '');
  return s;
}

// 표준품목명이 있는 map_product 행 + 표준품목명 자체를 정규화 키로 인덱싱합니다.
function buildProductIndex_(prMap) {
  var index = {};
  var name;
  for (name in prMap) {
    var p = prMap[name];
    if (!p.std) continue;
    var key = normItem_(name);
    if (key && !index[key]) index[key] = { std: p.std, cat: p.cat, parts: p.parts };
  }
  for (name in prMap) {
    var q = prMap[name];
    if (!q.std) continue;
    var stdKey = normItem_(q.std);
    if (stdKey && !index[stdKey]) index[stdKey] = { std: q.std, cat: q.cat, parts: q.parts };
  }
  return index;
}

function repeatParts_(base, n) {
  var unit = base.parts ? base.parts.split(',').map(function(x) { return x.trim(); }).filter(Boolean) : [base.std];
  var out = [];
  for (var i = 0; i < n; i++) out = out.concat(unit);
  return out;
}

// 단일 품목(세트 아님) 추론: 일치 → 배수 → 포함
function inferSingleProduct_(name, index, allowFuzzy) {
  var n = normItem_(name);
  if (!n) return null;
  if (index[n]) {
    var hit = index[n];
    return { std: hit.std, cat: hit.cat, parts: hit.parts, source: '자동·일치' };
  }

  // 배수: (NEA) 한 번, '+' 없음
  var m = name.match(/^(.*?)\(\s*(\d+)\s*EA\s*\)(.*)$/i);
  if (m && name.indexOf('+') < 0) {
    var count = parseInt(m[2], 10) || 1;
    var body = (m[1] + ' ' + m[3]).trim();
    var base = inferSingleProduct_(body, index, false);
    if (base && base.std) {
      if (count === 1) return { std: base.std, cat: base.cat, parts: base.parts, source: '자동·일치' };
      return {
        std: base.std + ' * ' + count,
        cat: base.cat,
        parts: repeatParts_(base, count).join(','),
        source: '자동·배수'
      };
    }
  }

  // 포함: 정규화 문자열이 기존 키를 포함하고 남는 글자가 6자 이하
  if (allowFuzzy) {
    var best = '';
    for (var key in index) {
      if (key.length >= 4 && n.indexOf(key) >= 0 && n.length - key.length <= 6 && key.length > best.length) best = key;
    }
    if (best) {
      var f = index[best];
      return { std: f.std, cat: f.cat, parts: f.parts, source: '추론·유사' };
    }
  }
  return null;
}

// 품목 추론 진입점. 세트(+)는 조각마다 풀어 구성단품을 나열합니다.
function inferProduct_(name, index) {
  name = String(name || '').trim();
  if (!name) return null;

  if (name.indexOf('+') >= 0) {
    var segments = name.split('+');
    var parts = [];
    for (var i = 0; i < segments.length; i++) {
      var seg = segments[i].trim();
      var count = 1;
      var em = seg.match(/^(.*?)\(\s*(\d+)\s*EA\s*\)\s*$/i);
      if (em) { seg = em[1].trim(); count = parseInt(em[2], 10) || 1; }
      var base = inferSingleProduct_(seg, index, true);
      if (!base || !base.std) return null;
      parts = parts.concat(repeatParts_(base, count));
    }
    return { std: name, cat: '세트', parts: parts.join(','), source: '자동·세트' };
  }

  return inferSingleProduct_(name, index, true);
}

// 채널 추론: 기존 쇼핑몰명을 포함하면 그 행 복사 → 채널그룹명을 포함하면 그 그룹의 최빈 수수료·담당자
function inferChannel_(shop, chMap) {
  shop = String(shop || '').trim();
  if (!shop) return null;

  var bestName = '';
  for (var name in chMap) {
    if (name !== shop && name.length >= 2 && shop.indexOf(name) >= 0 && name.length > bestName.length) bestName = name;
  }
  if (bestName) {
    var c = chMap[bestName];
    return { group: c.group, fee: c.fee, manager: c.manager, source: '추론·포함', basis: bestName };
  }

  var groups = {};
  for (var n2 in chMap) {
    var g = chMap[n2].group;
    if (!g || g === '미매핑') continue;
    if (!groups[g]) groups[g] = { fees: {}, managers: {} };
    groups[g].fees[chMap[n2].fee] = (groups[g].fees[chMap[n2].fee] || 0) + 1;
    groups[g].managers[chMap[n2].manager] = (groups[g].managers[chMap[n2].manager] || 0) + 1;
  }
  var bestGroup = '';
  for (var gName in groups) {
    if (gName.length >= 2 && shop.indexOf(gName) >= 0 && gName.length > bestGroup.length) bestGroup = gName;
  }
  if (bestGroup) {
    return {
      group: bestGroup,
      fee: num_(mostCommonKey_(groups[bestGroup].fees)),
      manager: mostCommonKey_(groups[bestGroup].managers) || '미지정',
      source: '추론·그룹',
      basis: bestGroup
    };
  }
  return null;
}

function mostCommonKey_(counts) {
  var best = '', bestN = -1;
  for (var k in counts) { if (counts[k] > bestN) { bestN = counts[k]; best = k; } }
  return best;
}

// raw 신규 행 + 과거 가공 행(채널그룹 미매핑)에서 map_channel에 없는 쇼핑몰명을 모읍니다.
function collectUnknownShops_(rawData, existingRows, chMap) {
  var seen = {};
  var out = [];
  var push = function(shop) {
    shop = text_(shop);
    if (!shop || chMap[shop] || seen[shop]) return;
    seen[shop] = true;
    out.push(shop);
  };
  for (var i = 0; i < rawData.length; i++) push(rawData[i][5]);
  for (var j = 0; j < existingRows.length; j++) {
    if (text_(existingRows[j][8], '미매핑') === '미매핑') push(existingRows[j][5]);
  }
  return out;
}

// 채운 키가 과거 가공 행에 미매핑으로 남아 있는지
function hasBacklogRows_(existingRows, fill) {
  if (!existingRows.length) return false;
  for (var i = 0; i < existingRows.length; i++) {
    var row = existingRows[i];
    if (fill.filledShops[text_(row[5])] && text_(row[8], '미매핑') === '미매핑') return true;
    if (fill.filledItems[text_(row[1])] && text_(row[9], '미매핑') === '미매핑') return true;
  }
  return false;
}

function ensureHeader_(sheet, col, title) {
  if (sheet.getMaxColumns() < col) sheet.insertColumnsAfter(sheet.getMaxColumns(), col - sheet.getMaxColumns());
  var cell = sheet.getRange(1, col);
  if (String(cell.getValue() || '').trim() === title) return;
  cell.setValue(title).setFontWeight('bold').setBackground('#1F2937').setFontColor('#FFFFFF');
}

// 빈 칸만 채웁니다. 반환값은 리포트·메시지용 요약입니다.
function fillMappings_(ss, unknownShops, chMap, prMap, costMap) {
  var result = {
    changed: false,
    products: [], productsOpen: [], productDupes: [],
    channels: [], channelsOpen: [],
    costsAdded: [],
    filledItems: {}, filledShops: {},
    counts: { match: 0, multiple: 0, set: 0, fuzzy: 0 }
  };
  var index = buildProductIndex_(prMap);
  var costNeeded = {};

  // ---- map_product: 표준품목명이 빈 행 ----
  var prSheet = ss.getSheetByName('map_product');
  if (prSheet && prSheet.getLastRow() > 1) {
    ensureHeader_(prSheet, 7, '매핑출처');
    var prLast = prSheet.getLastRow();
    var prData = prSheet.getRange(2, 1, prLast - 1, 7).getValues();
    var nameCount = {};
    var i;
    for (i = 0; i < prData.length; i++) {
      var nm = text_(prData[i][0]);
      if (nm && nm !== '-') nameCount[nm] = (nameCount[nm] || 0) + 1;
    }
    for (i = 0; i < prData.length; i++) {
      var origName = text_(prData[i][0]);
      if (!origName || origName === '-') continue;
      var rowNo = i + 2;
      if (nameCount[origName] > 1 && result.productDupes.indexOf(origName) < 0) result.productDupes.push(origName);
      if (text_(prData[i][1])) continue; // 사람이 채운 행은 건드리지 않음

      var guess = inferProduct_(origName, index);
      if (guess) {
        prSheet.getRange(rowNo, 2, 1, 2).setValues([[guess.std, guess.cat || '']]);
        prSheet.getRange(rowNo, 5).setValue(guess.parts || '');
        prSheet.getRange(rowNo, 7).setValue(guess.source);
        prSheet.getRange(rowNo, 1, 1, 7).setBackground(guess.source.indexOf('추론') === 0 ? FILL_COLOR_INFER : FILL_COLOR_AUTO);
        result.products.push({ name: origName, std: guess.std, cat: guess.cat, parts: guess.parts, source: guess.source });
        result.filledItems[origName] = true;
        var newKey = normItem_(origName);
        if (newKey && !index[newKey]) index[newKey] = { std: guess.std, cat: guess.cat, parts: guess.parts };
        result.changed = true;
        if (guess.source === '자동·일치') result.counts.match++;
        else if (guess.source === '자동·배수') result.counts.multiple++;
        else if (guess.source === '자동·세트') result.counts.set++;
        else result.counts.fuzzy++;
        // 원가가 필요한 단품 이름 수집
        var baseNames = guess.parts ? guess.parts.split(',') : [guess.std.split(' * ')[0]];
        for (var b = 0; b < baseNames.length; b++) {
          var bn = baseNames[b].trim();
          if (bn && !costMap[bn]) costNeeded[bn] = true;
        }
      } else {
        prSheet.getRange(rowNo, 7).setValue('미해결');
        prSheet.getRange(rowNo, 1, 1, 7).setBackground(FILL_COLOR_OPEN);
        result.productsOpen.push(origName);
      }
    }
  }

  // ---- map_channel: 없는 쇼핑몰명 추가 ----
  var chSheet = ss.getSheetByName('map_channel');
  if (chSheet && unknownShops.length > 0) {
    ensureHeader_(chSheet, 6, '매핑출처');
    var chRows = [];
    var chColors = [];
    for (var u = 0; u < unknownShops.length; u++) {
      var shop = unknownShops[u];
      var cg = inferChannel_(shop, chMap);
      if (cg) {
        chRows.push([shop, cg.group, 'Y', cg.fee, cg.manager, cg.source + ' (' + cg.basis + ')']);
        chColors.push(FILL_COLOR_INFER);
        result.channels.push({ shop: shop, group: cg.group, fee: cg.fee, manager: cg.manager, source: cg.source, basis: cg.basis });
        result.filledShops[shop] = true;
        result.changed = true;
      } else {
        chRows.push([shop, '', 'Y', '', '', '미해결']);
        chColors.push(FILL_COLOR_OPEN);
        result.channelsOpen.push(shop);
        result.changed = true; // 행이 추가되므로 재로드 필요(값은 미매핑 그대로)
      }
    }
    var chStart = chSheet.getLastRow() + 1;
    chSheet.getRange(chStart, 1, chRows.length, 6).setValues(chRows);
    for (var c = 0; c < chRows.length; c++) chSheet.getRange(chStart + c, 1, 1, 6).setBackground(chColors[c]);
    chSheet.getRange(chStart, 4, chRows.length, 1).setNumberFormat('0.0%'); // 0.15 → 15.0% (기존 행과 같은 표기)
  }

  // ---- map_cost: 완전히 새 단품만 빈 행으로 ----
  var costSheet = ss.getSheetByName('map_cost');
  var costNames = Object.keys(costNeeded).filter(function(n) { return n.indexOf(' * ') < 0 && n.indexOf('+') < 0; });
  if (costSheet && costNames.length > 0) {
    var costStart = costSheet.getLastRow() + 1;
    var costRows = costNames.map(function(n) { return [n, '', '', '']; });
    costSheet.getRange(costStart, 1, costRows.length, 4).setValues(costRows);
    costSheet.getRange(costStart, 1, costRows.length, 4).setBackground(FILL_COLOR_OPEN);
    result.costsAdded = costNames;
  }

  return result;
}

function fillSummaryText_(fill, fillError) {
  if (fillError) return '매핑 자동 채움: 실패 (' + fillError + ')';
  if (!fill) return '매핑 자동 채움: 실행 안 됨';
  var c = fill.counts;
  return '매핑 자동: 상품 ' + fill.products.length +
    '(일치 ' + c.match + '·배수 ' + c.multiple + '·세트 ' + c.set + '·유사 ' + c.fuzzy + ')' +
    ' · 채널 ' + fill.channels.length +
    ' · 미해결 ' + (fill.productsOpen.length + fill.channelsOpen.length) +
    ' · 원가 입력 필요 ' + fill.costsAdded.length +
    ' → ' + REPORT_SHEET_NAME + ' 탭';
}

// 최근 N일 가공 행에서 쇼핑몰명 / 품목명 / 표준품목명별 행 수·매출을 집계합니다.
function recentImpact_(rows) {
  var maxDate = '';
  var i;
  for (i = 0; i < rows.length; i++) {
    var d = normalizeDateOnly_(rows[i][7] || rows[i][6]);
    if (d > maxDate) maxDate = d;
  }
  var cutoff = '';
  if (maxDate) {
    var dt = new Date(maxDate + 'T00:00:00');
    dt.setDate(dt.getDate() - REPORT_RECENT_DAYS);
    cutoff = normalizeDateOnly_(dt);
  }
  var byShop = {}, byItem = {}, byStd = {};
  var add = function(map, key, rev) {
    if (!key) return;
    if (!map[key]) map[key] = { rows: 0, rev: 0 };
    map[key].rows++;
    map[key].rev += rev;
  };
  for (i = 0; i < rows.length; i++) {
    var r = rows[i];
    var date = normalizeDateOnly_(r[7] || r[6]);
    if (cutoff && date < cutoff) continue;
    var rev = num_(r[3]);
    add(byShop, text_(r[5]), rev);
    add(byItem, text_(r[1]), rev);
    add(byStd, text_(r[9]), rev);
  }
  return { cutoff: cutoff, maxDate: maxDate, byShop: byShop, byItem: byItem, byStd: byStd };
}

function impactOf_(map, key) {
  var v = map[key];
  return v ? v : { rows: 0, rev: 0 };
}

// '점검_리포트' 탭. 매 실행 덮어씁니다.
function writeCheckReport_(ss, ctx) {
  var sheet = ss.getSheetByName(REPORT_SHEET_NAME) || ss.insertSheet(REPORT_SHEET_NAME);
  sheet.clear();
  var fill = ctx.fill;
  var impact = recentImpact_(ctx.rows || []);
  var period = impact.cutoff ? (impact.cutoff + ' ~ ' + impact.maxDate) : '-';

  var review = []; // 검토 필요 (파랑)
  var todo = [];   // 사람이 채울 것 (노랑)
  var i;
  if (fill) {
    for (i = 0; i < fill.products.length; i++) {
      var p = fill.products[i];
      if (p.source.indexOf('추론') !== 0) continue;
      var pi = impactOf_(impact.byItem, p.name);
      review.push(['map_product', p.name, '표준 ' + p.std + ' / ' + (p.cat || '') + (p.parts ? ' / ' + p.parts : ''), p.source, pi.rows, pi.rev]);
    }
    for (i = 0; i < fill.channels.length; i++) {
      var ch = fill.channels[i];
      var ci = impactOf_(impact.byShop, ch.shop);
      review.push(['map_channel', ch.shop, ch.group + ' / 수수료 ' + Math.round(ch.fee * 1000) / 10 + '% / ' + ch.manager + ' — 수수료 확인', ch.source + ' (' + ch.basis + ')', ci.rows, ci.rev]);
    }
    for (i = 0; i < fill.productsOpen.length; i++) {
      var oi = impactOf_(impact.byItem, fill.productsOpen[i]);
      todo.push(['map_product', fill.productsOpen[i], 'B 표준품목명 · C 상품군 (· E 구성단품)', oi.rows, oi.rev]);
    }
    for (i = 0; i < fill.channelsOpen.length; i++) {
      var si = impactOf_(impact.byShop, fill.channelsOpen[i]);
      todo.push(['map_channel', fill.channelsOpen[i], 'B 채널그룹 · D 수수료율 · E 담당자', si.rows, si.rev]);
    }
    for (i = 0; i < fill.costsAdded.length; i++) {
      var ki = impactOf_(impact.byStd, fill.costsAdded[i]);
      todo.push(['map_cost', fill.costsAdded[i], 'B 단품원가 (새 단품)', ki.rows, ki.rev]);
    }
    for (i = 0; i < fill.productDupes.length; i++) {
      var di = impactOf_(impact.byItem, fill.productDupes[i]);
      todo.push(['map_product', fill.productDupes[i], '원본 품목명 중복 행 — 하나만 남기기', di.rows, di.rev]);
    }
  }
  // 원가 0/1인 표준품목명 (매출이 있는 것만)
  var costMap = ctx.costMap || {};
  for (var std in impact.byStd) {
    if (!std || std === '미매핑') continue;
    var cost = costMap[std];
    if (cost !== undefined && cost <= 1 && impact.byStd[std].rev > 0) {
      todo.push(['map_cost', std, 'B 단품원가가 ' + cost + ' — 실제 원가 입력', impact.byStd[std].rows, impact.byStd[std].rev]);
    }
  }
  review.sort(function(a, b) { return b[5] - a[5]; });
  todo.sort(function(a, b) { return b[4] - a[4]; });

  var out = [];
  out.push(['점검 리포트', new Date().toLocaleString('ko-KR'), '', '', '', '']);
  out.push(['갱신 모드', ctx.mode, '', '', '', '']);
  out.push(['자동 채움', fill ? ('상품 ' + fill.products.length + ' · 채널 ' + fill.channels.length) : (ctx.fillError ? '실패: ' + ctx.fillError : '-'), '', '', '', '']);
  out.push(['검토 필요(파랑)', review.length, '사람이 채울 것(노랑)', todo.length, '', '']);
  out.push(['영향 집계 기간', period, '최근 ' + REPORT_RECENT_DAYS + '일 · 가공_데이터 기준', '', '', '']);
  out.push(['', '', '', '', '', '']);
  var reviewHeaderRow = out.length + 1;
  out.push(['[검토 필요] 시트', '키', '채운 값', '출처', '최근 행수', '최근 매출']);
  if (!review.length) out.push(['-', '없음', '', '', '', '']);
  for (i = 0; i < review.length; i++) out.push(review[i]);
  out.push(['', '', '', '', '', '']);
  var todoHeaderRow = out.length + 1;
  out.push(['[사람이 채울 것] 시트', '키', '비어 있는 열 / 조치', '최근 행수', '최근 매출', '']);
  if (!todo.length) out.push(['-', '없음', '', '', '', '']);
  for (i = 0; i < todo.length; i++) out.push(todo[i].concat(['']));

  sheet.getRange(1, 1, out.length, 6).setValues(out);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  [reviewHeaderRow, todoHeaderRow].forEach(function(r) {
    sheet.getRange(r, 1, 1, 6).setFontWeight('bold').setBackground('#1F2937').setFontColor('#FFFFFF');
  });
  if (review.length) sheet.getRange(reviewHeaderRow + 1, 1, review.length, 6).setBackground(FILL_COLOR_INFER);
  if (todo.length) sheet.getRange(todoHeaderRow + 1, 1, todo.length, 6).setBackground(FILL_COLOR_OPEN);
  if (review.length) sheet.getRange(reviewHeaderRow + 1, 6, review.length, 1).setNumberFormat('#,##0');
  if (todo.length) sheet.getRange(todoHeaderRow + 1, 5, todo.length, 1).setNumberFormat('#,##0');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 360);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 120);
}

// ============================================================
// 매핑 테이블
// ============================================================
function loadChannelMap(ss) {
  var m = {};
  var s = ss.getSheetByName('map_channel');
  if (!s || s.getLastRow() < 2) return m;

  var colCount = Math.max(s.getLastColumn(), 5);
  var d = s.getRange(2, 1, s.getLastRow() - 1, colCount).getValues();

  for (var i = 0; i < d.length; i++) {
    var n = String(d[i][0]).trim();
    var g = String(d[i][1]).trim();
    var f = d[i][3] ? num_(d[i][3]) : 0;
    if (f > 1) f = f / 100;
    var mgr = d[i][4] ? String(d[i][4]).trim() : '미지정';
    if (n) m[n] = { group: g, fee: f, manager: mgr };
  }

  return m;
}

function loadProductMap(ss) {
  var m = {};
  var s = ss.getSheetByName('map_product');
  if (!s || s.getLastRow() < 2) return m;

  var d = s.getRange(2, 1, s.getLastRow() - 1, Math.max(s.getLastColumn(), 5)).getValues();
  for (var i = 0; i < d.length; i++) {
    var o = String(d[i][0]).trim();
    var st = String(d[i][1]).trim();
    var c = String(d[i][2]).trim();
    var p = d[i][4] ? String(d[i][4]).trim() : '';
    if (o) m[o] = { std: st, cat: c, parts: p };
  }

  return m;
}

function loadCostMap(ss) {
  var m = {};
  var s = ss.getSheetByName('map_cost');
  if (!s || s.getLastRow() < 2) return m;

  var d = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < d.length; i++) {
    var n = String(d[i][0]).trim();
    var c = num_(d[i][1]);
    if (n) m[n] = c;
  }

  return m;
}

function loadShipCost(ss) {
  var s = ss.getSheetByName('map_cost');
  if (!s) return 4521;
  var v = s.getRange('D2').getValue();
  return num_(v) || 4521;
}

function calcProductCost(stdName, partsStr, costMap, qty) {
  if (partsStr && partsStr.length > 0) {
    var parts = partsStr.split(',');
    var total = 0;
    for (var i = 0; i < parts.length; i++) {
      total += (costMap[parts[i].trim()] || 0);
    }
    return total * qty;
  }
  return (costMap[stdName] || 0) * qty;
}

// ============================================================
// 날짜 파서
// ============================================================
function parseAnyDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return fmtDate(val);
  }
  if (typeof val === 'number' && val > 30000 && val < 60000) {
    var d = new Date((val - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return fmtDate(d);
    return '';
  }

  var s = String(val).trim();
  if (!s) return '';

  var m = s.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})(.*)/);
  if (m) {
    var o = m[1] + '-' + p2(m[2]) + '-' + p2(m[3]);
    var tp = m[4] ? m[4].trim() : '';
    if (tp) {
      var t = tp.match(/(\d{1,2})[:\.](\d{2})(?:[:\.](\d{2}))?/);
      if (t) o += ' ' + p2(t[1]) + ':' + p2(t[2]) + ':' + p2(t[3] || '0');
    }
    return o;
  }

  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(.*)/);
  if (m) {
    var o2 = m[1] + '-' + p2(m[2]) + '-' + p2(m[3]);
    var t2 = m[4] ? m[4].trim() : '';
    if (t2) {
      var tt = t2.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (tt) o2 += ' ' + p2(tt[1]) + ':' + p2(tt[2]) + ':' + p2(tt[3] || '0');
    }
    return o2;
  }

  m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) return m[1] + '-' + p2(m[2]) + '-' + p2(m[3]);

  var last = new Date(s);
  if (!isNaN(last.getTime())) return fmtDate(last);
  return '';
}

function p2(n) {
  var s = String(n);
  return s.length < 2 ? '0' + s : s;
}

function fmtDate(d) {
  return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()) +
    ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()) + ':' + p2(d.getSeconds());
}

function testGetData() {
  var r = getProcessedData();
  var lines = [];
  lines.push('총 행: ' + r.orders.length);

  var ids = {};
  for (var i = 0; i < r.orders.length; i++) {
    ids[r.orders[i].id] = true;
  }
  lines.push('고유 주문: ' + Object.keys(ids).length);

  if (r.orders.length > 0) {
    var o = r.orders[0];
    lines.push(
      '첫번째: 매출=' + o.revenue +
      ' 정산=' + o.settlement +
      ' 원가=' + o.cost +
      ' 배송=' + o.shipCost +
      ' 마진=' + o.margin +
      ' 담당자=' + o.manager +
      ' 날짜=' + o.date
    );
  }

  lines.push('미매핑: ' + r.unmapped.length + '건');
  if (r.unmapped.length > 0) {
    Logger.log('미매핑 예시: ' + JSON.stringify(r.unmapped.slice(0, 3)));
  }

  var cache = getDashboardCache();
  lines.push('캐시 meta: ' + JSON.stringify(cache.meta));

  var state = loadProcState_();
  lines.push('증분 상태: ' + (state ? JSON.stringify(state) : '없음(다음 갱신은 전체 재생성)'));

  var mgrAgg = {};
  for (var j = 0; j < r.orders.length; j++) {
    var mgr = r.orders[j].manager || '미지정';
    if (!mgrAgg[mgr]) mgrAgg[mgr] = { rev: 0, cnt: 0 };
    mgrAgg[mgr].rev += r.orders[j].revenue;
    mgrAgg[mgr].cnt++;
  }
  lines.push('담당자별: ' + JSON.stringify(mgrAgg));

  notifyUser_('📋 데이터 확인\n\n' + lines.join('\n'));
}
