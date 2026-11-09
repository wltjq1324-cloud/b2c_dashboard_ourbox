// 채널 멀티선택(a+b+c 합산) 브라우저 시험.
// dashboard.src.html을 실제 Chromium에 띄우고 Apps Script 응답을 가짜 데이터로 대체한 뒤
// 채널 행 클릭·칩 해제·드롭다운·필터 초기화·담당자 페이지 전환을 눌러 보며 KPI 합계와 선택 상태를 대조하고
// 페이지 오류·콘솔 error가 0건인지 확인한다.
//
// Playwright가 필요하다(`node --test test/*.test.mjs`의 캐시 시험과 달리 설치가 필요해 이름을 *.e2e.mjs로 분리).
//   npm i -D playwright && npx playwright install chromium
//   node test/channel-multi.e2e.mjs [스크린샷.png]
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';

let chromium;
try { ({ chromium } = createRequire(import.meta.url)('playwright')); }
catch { console.log('SKIP: playwright가 없어 채널 멀티선택 E2E를 건너뜁니다 (npm i -D playwright && npx playwright install chromium)'); process.exit(0); }

const SRC = new URL('../dashboard.src.html', import.meta.url).href;
const SHOT = process.argv[2] || '';

// 가짜 데이터: 2026-11-01~05, 채널 3개. 일 매출 카카오 10만 / 네이버 20만 / 자사몰 5만 → 합 175만
const days = ['2026-11-01', '2026-11-02', '2026-11-03', '2026-11-04', '2026-11-05'];
const chans = [
  { channel: '카카오', rev: 100000, category: '단품', manager: '담당A' },
  { channel: '네이버', rev: 200000, category: '세트', manager: '담당B' },
  { channel: '자사몰', rev: 50000, category: '단품', manager: '담당A' },
];
const mk = (d, c, extra = {}) => ({
  date: d, channel: c.channel, category: c.category, manager: c.manager,
  qty: 1, revenue: c.rev, settlement: c.rev * 0.9, cost: c.rev * 0.3, shipCost: 0,
  margin: c.rev * 0.6, orders: 1, ...extra,
});
const baseRows = days.flatMap((d) => chans.map((c) => mk(d, c)));
const productRows = days.flatMap((d) => chans.map((c) => mk(d, c, { product: '감자빵' })));
const fixture = { dashboardCache: { baseRows, productRows, qualityRows: [], meta: { servedFrom: 'cache' } }, meta: {} };
const TOTAL = 5 * (100000 + 200000 + 50000);

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1400, height: 1000 }, deviceScaleFactor: 2 });
const pageErrors = [], consoleErrors = [];
page.on('pageerror', (e) => pageErrors.push(String(e)));
page.on('console', (m) => {
  if (m.type() !== 'error') return;
  const loc = (m.location && m.location().url) || '';
  if (/cdn\.jsdelivr\.net/.test(loc)) return; // 아래 route.abort()로 일부러 막은 Chart.js CDN 로드 실패는 예외
  consoleErrors.push(m.text() + (loc ? ' @ ' + loc : ''));
});
await page.route('**/script.google.com/**', (r) => r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(fixture) }));
await page.route('**/cdn.jsdelivr.net/**', (r) => r.abort()); // Chart.js 없이도 돌아가야 한다(기존 폴백 chartLibMissing)

await page.goto(SRC);
await page.waitForFunction(() => typeof ALL !== 'undefined' && ALL.length > 0 && document.querySelectorAll('#channelProgress [data-channel]').length > 0, null, { timeout: 15000 });

const state = () => page.evaluate(() => {
  const el = document.getElementById('channelFilter');
  return {
    sel: [...CHANNEL_SEL].sort(),
    rev: sum(filtered(), 'revenue'),
    selVal: el.value,
    selText: el.options[el.selectedIndex] ? el.options[el.selectedIndex].textContent : '',
    chips: document.querySelectorAll('[data-channel-pill]').length,
    clear: !!document.querySelector('[data-channel-clear]'),
    hiRows: [...document.querySelectorAll('#channelProgress .selected')].map((e) => e.dataset.channel).sort(),
    hiTable: [...document.querySelectorAll('#channelTable tr.selected')].map((e) => e.dataset.channel).sort(),
    kpi: document.getElementById('kRevenue').textContent,
    shown: document.querySelectorAll('#channelProgress [data-channel]').length,
    dimRows: [...document.querySelectorAll('#channelProgress .dim')].map((e) => e.dataset.channel).sort(),
    totalsLabel: (document.querySelector('#profitStructureTable tbody tr:last-child td') || {}).textContent || '',
    totalsRev: (document.querySelectorAll('#profitStructureTable tbody tr:last-child td')[2] || {}).textContent || '',
  };
});
const krw = (n) => page.evaluate((v) => KRW(v), n);
const click = (sel) => page.click(sel, { timeout: 5000 });
let s;

// 0) 초기 상태: 전체
s = await state();
assert.deepEqual(s.sel, []); assert.equal(s.rev, TOTAL); assert.equal(s.selVal, 'ALL'); assert.equal(s.chips, 0);
assert.equal(s.kpi, await krw(TOTAL));
console.log('✔ 0 초기: 전체', s.rev);

// 1) 막대 행 클릭 → 단일 선택. 채널 목록은 전부 남고(다음 채널을 더 누를 수 있어야 함) 나머지는 흐림
await click('#channelProgress [data-channel="카카오"]'); s = await state();
assert.deepEqual(s.sel, ['카카오']); assert.equal(s.rev, 500000); assert.equal(s.selVal, '카카오'); assert.equal(s.chips, 1); assert.equal(s.clear, false);
assert.deepEqual(s.hiRows, ['카카오']); assert.equal(s.kpi, await krw(500000));
assert.equal(s.shown, 3, '선택 후에도 채널 목록은 전부 보여야 다음 채널을 더 누를 수 있다'); assert.deepEqual(s.dimRows, ['네이버', '자사몰']);
assert.equal(s.totalsLabel, '합계 (선택 1개 채널)'); assert.equal(s.totalsRev, await krw(500000));
console.log('✔ 1 카카오 클릭 → 단일', s.rev, '| 목록 유지', s.shown, '| 흐림', s.dimRows);

// 2) 다른 패널(채널별 수익 구조 표) 행 클릭 → 누적(a+b)
await click('#channelTable [data-channel="네이버"]'); s = await state();
assert.deepEqual(s.sel, ['네이버', '카카오']); assert.equal(s.rev, 1500000); assert.equal(s.selVal, '__MULTI'); assert.equal(s.selText, '2개 채널 합산');
assert.equal(s.chips, 2); assert.equal(s.clear, true); assert.deepEqual(s.hiRows, ['네이버', '카카오']); assert.deepEqual(s.hiTable, ['네이버', '카카오']);
assert.equal(s.kpi, await krw(1500000));
console.log('✔ 2 +네이버 → 2개 합산', s.rev);
if (SHOT) { await page.evaluate(() => window.scrollTo(0, 0)); await page.screenshot({ path: SHOT }); }

// 3) 수익구조(상단 표) 행 클릭 → a+b+c
await click('#profitStructureTable [data-channel="자사몰"]'); s = await state();
assert.deepEqual(s.sel, ['네이버', '자사몰', '카카오']); assert.equal(s.rev, TOTAL); assert.equal(s.selText, '3개 채널 합산'); assert.equal(s.chips, 3);
console.log('✔ 3 +자사몰 → 3개 합산', s.rev);

// 4) 칩 ✕ 로 하나 해제
await click('[data-channel-pill="카카오"]'); s = await state();
assert.deepEqual(s.sel, ['네이버', '자사몰']); assert.equal(s.rev, 1250000); assert.equal(s.chips, 2);
console.log('✔ 4 칩으로 카카오 해제', s.rev);

// 5) 전체 해제 칩
await click('[data-channel-clear]'); s = await state();
assert.deepEqual(s.sel, []); assert.equal(s.rev, TOTAL); assert.equal(s.selVal, 'ALL'); assert.equal(s.chips, 0); assert.equal(s.clear, false);
assert.deepEqual(s.dimRows, []); assert.equal(s.totalsLabel, '합계 (미매핑 포함)'); assert.equal(s.totalsRev, await krw(TOTAL));
console.log('✔ 5 전체 해제 칩', s.rev);

// 6) 드롭다운 = 단일 교체
await page.selectOption('#channelFilter', '네이버'); s = await state();
assert.deepEqual(s.sel, ['네이버']); assert.equal(s.rev, 1000000);
await click('#channelProgress [data-channel="카카오"]'); s = await state();
assert.deepEqual(s.sel, ['네이버', '카카오']);
await page.selectOption('#channelFilter', '자사몰'); s = await state();
assert.deepEqual(s.sel, ['자사몰'], '드롭다운 선택은 기존 선택을 교체해야 한다'); assert.equal(s.rev, 250000);
await page.selectOption('#channelFilter', 'ALL'); s = await state();
assert.deepEqual(s.sel, []); assert.equal(s.rev, TOTAL);
console.log('✔ 6 드롭다운 교체/전체');

// 7) 같은 행 재클릭 → 해제(마지막 하나면 전체로)
await click('#channelProgress [data-channel="카카오"]');
await click('#channelProgress [data-channel="카카오"]'); s = await state();
assert.deepEqual(s.sel, []); assert.equal(s.rev, TOTAL);
console.log('✔ 7 재클릭 해제');

// 8) 상품군 필터와 상호작용: 세트는 네이버만 → 카카오는 범위 밖이라 선택에서 빠짐(기존 단일 선택이 ALL로 돌아가던 동작과 동일)
await click('#channelProgress [data-channel="카카오"]');
await click('#channelProgress [data-channel="네이버"]');
await page.selectOption('#categoryFilter', '세트'); s = await state();
assert.deepEqual(s.sel, ['네이버']); assert.equal(s.rev, 1000000);
await page.selectOption('#categoryFilter', 'ALL'); s = await state();
assert.deepEqual(s.sel, ['네이버']); assert.equal(s.rev, 1000000);
console.log('✔ 8 상품군 필터 연동(범위 밖 채널 자동 제거)');

// 9) 필터 초기화 버튼
await click('button.btn.secondary'); s = await state();
assert.deepEqual(s.sel, []); assert.equal(s.rev, TOTAL); assert.equal(s.chips, 0);
console.log('✔ 9 필터 초기화');

// 10) 담당자 페이지 전환 중에도 멀티선택 유지·오류 없음
await click('#channelProgress [data-channel="카카오"]');
await click('#channelProgress [data-channel="자사몰"]');
await page.evaluate(() => switchPage('manager'));
const mgr = await page.evaluate(() => ({ cards: document.querySelectorAll('#managerCards .manager-card').length, sel: [...CHANNEL_SEL].sort(), rev: sum(filtered(), 'revenue') }));
assert.equal(mgr.cards, 1, '카카오+자사몰은 모두 담당A → 카드 1장'); assert.deepEqual(mgr.sel, ['자사몰', '카카오']); assert.equal(mgr.rev, 750000);
await page.evaluate(() => switchPage('overview')); s = await state();
assert.deepEqual(s.sel, ['자사몰', '카카오']); assert.equal(s.rev, 750000); assert.equal(s.chips, 2);
console.log('✔ 10 담당자 페이지 왕복', mgr);

// 11) 전월 비교(sameFilterRows)도 같은 채널 집합을 쓰는지
const prev = await page.evaluate(() => sameFilterRows('2026-11-01', '2026-11-05').length);
assert.equal(prev, 10, '카카오+자사몰 5일 → 10행');
console.log('✔ 11 sameFilterRows 멀티 반영');

assert.deepEqual(pageErrors, [], '페이지 오류 없어야 함');
assert.deepEqual(consoleErrors, [], '콘솔 error 없어야 함');
console.log('✔ 페이지 오류 0건 · 콘솔 error 0건');
await browser.close();
console.log('ALL PASSED');
