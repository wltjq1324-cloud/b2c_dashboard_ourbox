// Apps Script 캐시 경로 모의 시험: node --test test/
// apps-script-v*.js(가장 높은 버전)를 vm에 올리고 SpreadsheetApp·CacheService·LockService·PropertiesService를 흉내 낸다.
// 데이터는 전부 가짜(주문번호 T0001…, 담당A·B)다.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import vm from 'node:vm';

const ROOT = new URL('../', import.meta.url);
const FILE = readdirSync(ROOT).filter((f) => /^apps-script-v[\d.]+\.js$/.test(f))
  .sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).pop();
const CODE = readFileSync(new URL(FILE, ROOT), 'utf8');

class Range {
  constructor(sheet, r, c, nr = 1, nc = 1) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const src = this.sheet.values[this.r - 1 + i] || [];
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(src[this.c - 1 + j] ?? '');
      out.push(row);
    }
    return out;
  }
  setValues(v) {
    if (this.sheet.writeHook) this.sheet.writeHook(this, v);
    v.forEach((row, i) => {
      const at = this.r - 1 + i;
      while (this.sheet.values.length <= at) this.sheet.values.push([]);
      row.forEach((cell, j) => { this.sheet.values[at][this.c - 1 + j] = cell; });
    });
    return this;
  }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
}

class Sheet {
  constructor(ss, name) { Object.assign(this, { ss, name, values: [], hidden: false, writeHook: null }); }
  getName() { return this.name; }
  setName(n) {
    if (this.ss.renameHook) this.ss.renameHook(this, n);
    if (this.ss.sheets.some((s) => s !== this && s.name === n)) throw new Error(`이미 있는 이름: ${n}`);
    this.name = n;
    return this;
  }
  getParent() { return this.ss; }
  getLastRow() { return this.values.length; }
  getLastColumn() { return this.values.reduce((m, r) => Math.max(m, r.length), 0); }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr, nc); }
  setColumnWidth() { return this; }
  hideSheet() { this.hidden = true; return this; }
  isSheetHidden() { return this.hidden; }
  getFilter() { return null; }
  clear() { this.values = []; return this; }
}

class Spreadsheet {
  constructor() { this.sheets = []; this.active = null; this.renameHook = null; this.insertHook = null; }
  getId() { return 'test'; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n, idx) {
    if (this.getSheetByName(n)) throw new Error(`이미 있는 이름: ${n}`);
    const s = new Sheet(this, n);
    if (idx == null) this.sheets.push(s); else this.sheets.splice(idx, 0, s);
    this.active = s;
    if (this.insertHook) this.insertHook(s);
    return s;
  }
  deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
  getNumSheets() { return this.sheets.length; }
  getActiveSheet() { return this.active || this.sheets[0]; }
  setActiveSheet(s) { this.active = s; return s; }
}

function makeEnv() {
  const ss = new Spreadsheet();
  const store = new Map();
  const lockState = { busy: false };
  const props = new Map();
  const ctx = {
    console,
    Logger: { log() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {}, getUi() { throw new Error('UI 없음'); } },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => (store.has(k) ? store.get(k) : null),
        getAll: (keys) => Object.fromEntries(keys.filter((k) => store.has(k)).map((k) => [k, store.get(k)])),
        put: (k, v) => { store.set(k, v); },
        putAll: (o) => { for (const [k, v] of Object.entries(o)) store.set(k, v); },
        remove: (k) => store.delete(k),
      }),
    },
    LockService: { getDocumentLock: () => ({ tryLock: () => !lockState.busy, releaseLock() {} }) },
    PropertiesService: {
      getDocumentProperties: () => ({ getProperty: (k) => props.get(k) ?? null, setProperty: (k, v) => props.set(k, v) }),
    },
    ContentService: { createTextOutput: (t) => ({ text: t, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
  };
  vm.createContext(ctx);
  vm.runInContext(CODE, ctx, { filename: FILE });
  return { ss, store, lockState, props, ctx };
}

// 가짜 가공_데이터 n행(17열). 날짜·채널·상품·담당이 섞여 캐시 섹션이 여러 조각(45,000자)으로 나뉜다
function seedProcessed(env, n, revenueBase = 10000) {
  const sheet = env.ss.insertSheet('가공_데이터');
  const rows = [env.ctx.PROC_HEADERS.slice()];
  const channels = ['채널가', '채널나', '채널다', '채널라'];
  const products = ['감자빵 5개입', '감자빵 10개입', '고구마빵 4개입', '옥수수빵 3개입', '선물세트'];
  for (let i = 0; i < n; i++) {
    const day = String((i % 28) + 1).padStart(2, '0');
    const revenue = i % 3 === 0 ? 0 : revenueBase + (i % 13) * 1000; // 품질 행을 늘려 여러 조각으로 나뉘게
    rows.push([`T${String(i).padStart(5, '0')}`, `상품${i % 40}`, (i % 3) + 1, revenue, 0, `몰${i % 9}`,
      `2026-10-${day} 10:00:00`, `2026-10-${day}`, channels[i % 4], products[i % 5], '빵', 0.05,
      revenue * 0.95, 3000, 2500, revenue * 0.95 - 5500, i % 2 ? '담당A' : '담당B']);
  }
  sheet.values = rows;
  env.ss.setActiveSheet(sheet);
  return sheet;
}

const cacheSheet = (env) => env.ss.getSheetByName('dashboard_cache');
const names = (env) => env.ss.sheets.map((s) => s.name).sort();
const parse = (text) => JSON.parse(text);
const sorted = (arr) => arr.map((r) => JSON.stringify(r, Object.keys(r).sort())).sort();
const expireScriptCache = (env) => env.store.clear();

test('캐시 쓰기: 끝까지 쓴 뒤 바꿔 끼우고 임시·이전본을 남기지 않는다. 스크립트 캐시와 시트 경로가 같은 데이터, 보던 탭 유지', () => {
  const env = makeEnv();
  const proc = seedProcessed(env, 3000);
  const stats = env.ctx.refreshDashboardCache(env.ss);
  assert.equal(stats.rawRows, 3000);
  assert.deepEqual(names(env), ['dashboard_cache', '가공_데이터']);
  assert.equal(cacheSheet(env).isSheetHidden(), true);
  assert.equal(env.ctx.cacheSheetIsComplete_(env.ss), true);
  assert.equal(env.ss.getActiveSheet(), proc, '메뉴로 돌려도 보던 탭이 그대로');
  const sections = new Set(cacheSheet(env).values.slice(1).map((r) => r[0]));
  assert.deepEqual([...sections].sort(), ['baseRows', 'meta', 'productRows', 'qualityRows']);
  assert.ok(cacheSheet(env).values.filter((r) => r[0] === 'qualityRows').length > 1, '여러 조각으로 나뉜 섹션도 시험한다');

  const fromScript = parse(env.ctx.getDashboardCacheText_('summary'));
  assert.equal(fromScript.meta.servedFrom, 'script_cache');
  expireScriptCache(env);
  const fromSheet = parse(env.ctx.getDashboardCacheText_('summary'));
  assert.equal(fromSheet.meta.servedFrom, 'sheet');
  assert.deepEqual(fromSheet.dashboardCache.baseRows, fromScript.dashboardCache.baseRows);
  assert.deepEqual(fromSheet.dashboardCache.productRows, fromScript.dashboardCache.productRows);
  assert.equal(fromSheet.meta.warning, undefined);
});

test('쓰다 멈추면: 쓰던 캐시는 그대로이고 다음 갱신이 남은 임시 시트를 치운다', () => {
  const env = makeEnv();
  const proc = seedProcessed(env, 3000);
  env.ctx.refreshDashboardCache(env.ss);
  const before = JSON.stringify(cacheSheet(env).values);

  proc.values.push(['T99999', '상품x', 1, 999999, 0, '몰x', '2026-10-30 10:00:00', '2026-10-30', '채널가', '감자빵 5개입', '빵', 0.05, 1, 1, 1, 1, '담당A']);
  env.ss.insertHook = (s) => { if (s.name === 'dashboard_cache_next') s.writeHook = () => { throw new Error('실행 시간 초과(시험)'); }; };
  assert.throws(() => env.ctx.refreshDashboardCache(env.ss), /시간 초과/);
  assert.equal(JSON.stringify(cacheSheet(env).values), before, '쓰던 캐시는 한 글자도 안 바뀐다');
  assert.ok(env.ss.getSheetByName('dashboard_cache_next'), '멈춘 임시 시트가 남아 있다');
  expireScriptCache(env);
  assert.equal(parse(env.ctx.getDashboardCacheText_('summary')).meta.servedFrom, 'sheet', '읽는 쪽은 옛 캐시를 그대로 읽는다');

  env.ss.insertHook = null;
  const stats = env.ctx.refreshDashboardCache(env.ss);
  assert.equal(stats.rawRows, 3001);
  assert.deepEqual(names(env), ['dashboard_cache', '가공_데이터']);
  assert.notEqual(JSON.stringify(cacheSheet(env).values), before);
});

test('이름 바꾸기 사이에 멈추면: 읽는 쪽은 직전본(_prev)을 읽고, 점검이 다시 만든다', () => {
  const env = makeEnv();
  seedProcessed(env, 1500);
  env.ctx.refreshDashboardCache(env.ss);
  const old = parse(env.ctx.getDashboardCacheText_('summary')).dashboardCache;

  env.ss.renameHook = (s, n) => { if (n === 'dashboard_cache' && s.values.length) throw new Error('이름 바꾸기 실패(시험)'); };
  assert.throws(() => env.ctx.refreshDashboardCache(env.ss), /이름 바꾸기/);
  assert.equal(env.ss.getSheetByName('dashboard_cache'), null);
  assert.ok(env.ss.getSheetByName('dashboard_cache_prev'));
  expireScriptCache(env);
  const mid = parse(env.ctx.getDashboardCacheText_('summary'));
  assert.equal(mid.meta.servedFrom, 'sheet', '즉석 생성으로 떨어지지 않는다');
  assert.deepEqual(mid.dashboardCache.baseRows, old.baseRows);

  env.ss.renameHook = null;
  assert.match(env.ctx.keepDashboardCacheWarm(), /^rebuilt/);
  assert.deepEqual(names(env), ['dashboard_cache', '가공_데이터']);
});

test('캐시가 깨졌으면: 즉석 생성하고 결과를 스크립트 캐시에 넣어 다음 요청은 빠르다, 내용은 정상 캐시와 같다', () => {
  const env = makeEnv();
  seedProcessed(env, 2000);
  env.ctx.refreshDashboardCache(env.ss);
  expireScriptCache(env);
  const normal = parse(env.ctx.getDashboardCacheText_('all')).dashboardCache;

  const sheet = cacheSheet(env);
  sheet.values = sheet.values.filter((r) => r[0] !== 'productRows'); // 쓰다 끊긴 캐시처럼 섹션 하나가 없다
  assert.equal(env.ctx.cacheSheetIsComplete_(env.ss), false);
  expireScriptCache(env);
  const first = parse(env.ctx.getDashboardCacheText_('summary'));
  assert.equal(first.meta.servedFrom, 'built_on_demand');
  assert.match(first.meta.warning, /missing or incomplete/);
  const second = parse(env.ctx.getDashboardCacheText_('summary'));
  assert.equal(second.meta.servedFrom, 'script_cache');
  const quality = parse(env.ctx.getDashboardCacheText_('quality'));
  assert.equal(quality.meta.servedFrom, 'script_cache', '품질 요청도 같은 즉석본을 캐시에서 받는다');

  // 순서까지 같다(정렬하지 않고 비교). v3.9 즉석 생성은 품질 상세의 sheetRow가 빠져 여기서 걸렸다
  for (const k of ['baseRows', 'productRows']) assert.deepEqual(first.dashboardCache[k], normal[k], k);
  assert.deepEqual(quality.dashboardCache.qualityRows, normal.qualityRows);
  assert.ok(quality.dashboardCache.qualityRows.every((r) => Number.isInteger(r.sheetRow) && r.sheetRow >= 2), '품질 상세마다 시트 행 번호');

  // 스크립트 캐시가 즉석본으로 차 있어도 점검은 시트가 깨진 것을 보고 다시 만든다
  assert.match(env.ctx.keepDashboardCacheWarm(), /^rebuilt/);
  assert.equal(env.ctx.cacheSheetIsComplete_(env.ss), true);
});

test('자동 점검: 할 일 없음·스크립트 캐시만 만료·다른 갱신 중', () => {
  const env = makeEnv();
  seedProcessed(env, 500);
  assert.match(env.ctx.keepDashboardCacheWarm(), /^rebuilt/, '캐시 시트가 없으면 만든다');
  assert.match(env.ctx.keepDashboardCacheWarm(), /^ok/);
  expireScriptCache(env);
  assert.match(env.ctx.keepDashboardCacheWarm(), /^warmed/);
  assert.equal(parse(env.ctx.getDashboardCacheText_('summary')).meta.servedFrom, 'script_cache');
  env.lockState.busy = true;
  assert.match(env.ctx.keepDashboardCacheWarm(), /^skip/);
  const last = JSON.parse(env.props.get('OURBOX_CACHE_KEEPER_LAST'));
  assert.equal(last.result, 'skip');
});

test('지키는 선: 이름만 같은 다른 시트는 지우거나 바꾸지 않는다', () => {
  const env = makeEnv();
  seedProcessed(env, 300);
  const foreignNext = env.ss.insertSheet('dashboard_cache_next');
  foreignNext.values = [['메모'], ['사람이 쓴 내용']];
  assert.throws(() => env.ctx.refreshDashboardCache(env.ss), /캐시 형식이 아니어서 지우지 않았습니다/);
  assert.deepEqual(foreignNext.values, [['메모'], ['사람이 쓴 내용']]);
  env.ss.deleteSheet(foreignNext);

  const foreignMain = env.ss.insertSheet('dashboard_cache');
  foreignMain.values = [['다른', '표', '입니다'], ['1', '2', '3']];
  assert.throws(() => env.ctx.refreshDashboardCache(env.ss), /머리글이 캐시 형식이 아니어서/);
  assert.deepEqual(foreignMain.values, [['다른', '표', '입니다'], ['1', '2', '3']]);
  assert.equal(env.ss.getSheetByName('dashboard_cache_prev'), null);
});

test('doGet: mode=cache는 JSON 텍스트를 그대로 돌려준다', () => {
  const env = makeEnv();
  seedProcessed(env, 200);
  env.ctx.refreshDashboardCache(env.ss);
  const out = env.ctx.doGet({ parameter: { mode: 'cache', part: 'summary' } });
  const json = parse(out.text);
  assert.ok(Array.isArray(json.dashboardCache.baseRows) && json.dashboardCache.baseRows.length > 0);
  assert.equal(json.meta.version, 'v3.10');
});
