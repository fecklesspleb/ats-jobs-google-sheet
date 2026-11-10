/**
 * Local harness for Code.gs. Stubs UrlFetchApp, SpreadsheetApp, PropertiesService,
 * ScriptApp and Utilities. Does not call api.apify.com or any other network.
 *
 * Usage: node test/run.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(HERE, '..');
const CODE = fs.readFileSync(path.join(ROOT, 'Code.gs'), 'utf8');

const SAMPLE_ITEMS = [
  {
    id: '34413f8d-26bf-4bbc-8ade-eb309a0e2245',
    ats: 'ashby',
    companySlug: 'ramp',
    companyName: 'Ramp',
    title: 'Security Engineer, Cloud',
    location: 'New York, NY (HQ)',
    postedAt: '2026-04-07T17:12:35.753000+00:00',
    url: 'https://jobs.ashbyhq.com/ramp/34413f8d-26bf-4bbc-8ade-eb309a0e2245',
    salaryText: '$211.4K – $290.6K • Offers Equity',
    uid: 'ashby:ramp:34413f8d-26bf-4bbc-8ade-eb309a0e2245',
    workArrangement: 'hybrid',
  },
  {
    id: 'JR2001996',
    ats: 'workday',
    companySlug: 'nvidia/NVIDIAExternalCareerSite',
    companyName: 'NVIDIA USA',
    title: 'Principal Firmware Engineer',
    location: 'US, CA, Santa Clara',
    postedAt: '2026-10-28',
    url: 'https://nvidia.wd5.myworkdayjobs.com/NVIDIAExternalCareerSite/job/JR2001996',
    uid: 'workday:nvidia:JR2001996',
  },
  {
    status: 'unresolved',
    input: 'not-a-real-board-xyz',
    uid: 'unresolved:not-a-real-board-xyz',
    reason: 'No supported ATS board found for this entry',
  },
  {
    status: 'expired',
    uid: 'ashby:ramp:00000000-0000-0000-0000-000000000000',
    title: 'Closed role',
    url: 'https://jobs.ashbyhq.com/ramp/00000000-0000-0000-0000-000000000000',
  },
  { type: 'status', status: 'warning', message: 'Sample board could not be read' },
];

class FakeRange {
  constructor(sheet, r, c, nRows = 1, nCols = 1) {
    this.sheet = sheet;
    this.r = r;
    this.c = c;
    this.nRows = nRows;
    this.nCols = nCols;
  }
  setValues(values) {
    for (let i = 0; i < values.length; i++) {
      const rowIndex = this.r + i - 1;
      while (this.sheet.cells.length <= rowIndex) this.sheet.cells.push([]);
      for (let j = 0; j < values[i].length; j++) {
        const colIndex = this.c + j - 1;
        this.sheet.cells[rowIndex][colIndex] = values[i][j];
      }
    }
    return this;
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nRows; i++) {
      const row = this.sheet.cells[this.r + i - 1] || [];
      const line = [];
      for (let j = 0; j < this.nCols; j++) line.push(row[this.c + j - 1] ?? '');
      out.push(line);
    }
    return out;
  }
}

class FakeSheet {
  constructor(name, cells = []) {
    this.name = name;
    this.cells = cells;
  }
  getLastRow() {
    return this.cells.length;
  }
  getDataRange() {
    const nRows = Math.max(1, this.cells.length);
    const nCols = this.cells.reduce((m, r) => Math.max(m, r.length), 1);
    return new FakeRange(this, 1, 1, nRows, nCols);
  }
  getRange(r, c, nRows = 1, nCols = 1) {
    return new FakeRange(this, r, c, nRows, nCols);
  }
  appendRow(row) {
    this.cells.push(row.slice());
  }
}

class FakeSpreadsheet {
  constructor() {
    this.sheets = {};
  }
  getSheetByName(name) {
    return this.sheets[name] || null;
  }
  insertSheet(name) {
    const sheet = new FakeSheet(name);
    this.sheets[name] = sheet;
    return sheet;
  }
}

class FakeTrigger {
  constructor(handler) {
    this.handler = handler;
  }
  getHandlerFunction() {
    return this.handler;
  }
}

function makeWorld(opts = {}) {
  const props = { ...(opts.props || {}) };
  const ss = new FakeSpreadsheet();
  const fetches = [];
  const triggers = [];
  let fetchImpl = opts.fetchImpl;
  const now = opts.now || new Date('2026-11-10T16:00:00Z');

  const world = {
    fetches,
    triggers,
    ss,
    props,
    Date: class extends Date {
      constructor(...args) {
        if (args.length === 0) return super(now.getTime());
        return super(...args);
      }
      static now() {
        return now.getTime();
      }
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) {
            return Object.prototype.hasOwnProperty.call(props, key) ? props[key] : null;
          },
        };
      },
    },
    UrlFetchApp: {
      fetch(url, options) {
        fetches.push({ url, options: options || {} });
        if (typeof fetchImpl === 'function') return fetchImpl(url, options || {});
        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify(SAMPLE_ITEMS),
        };
      },
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return ss;
      },
    },
    ScriptApp: {
      getProjectTriggers() {
        return triggers.slice();
      },
      deleteTrigger(t) {
        const i = triggers.indexOf(t);
        if (i >= 0) triggers.splice(i, 1);
      },
      newTrigger(name) {
        const state = { name, daily: false, hour: null };
        const builder = {
          timeBased() {
            return builder;
          },
          everyDays(n) {
            state.daily = n;
            return builder;
          },
          atHour(h) {
            state.hour = h;
            return builder;
          },
          create() {
            const t = new FakeTrigger(state.name);
            t.daily = state.daily;
            t.hour = state.hour;
            triggers.push(t);
            return t;
          },
        };
        return builder;
      },
    },
    Utilities: {
      formatDate(date, _tz, _format) {
        return new Date(date).toISOString().replace(/\.\d{3}Z$/, 'Z');
      },
    },
    console,
    JSON,
    String,
    Array,
    Object,
    Error,
    encodeURIComponent,
    Math,
  };
  vm.createContext(world);
  vm.runInContext(CODE, world, { filename: 'Code.gs' });
  return world;
}

function httpResponse(code, body) {
  return {
    getResponseCode: () => code,
    getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)),
  };
}

let passed = 0;
function test(name, fn) {
  fn();
  passed += 1;
  console.log('ok  ' + name);
}

// --- tests

test('Code.gs has no hardcoded token or email', () => {
  assert.doesNotMatch(CODE, /apify_api_/i);
  assert.doesNotMatch(CODE, /Bearer [A-Za-z0-9_-]{8,}/);
  assert.doesNotMatch(CODE, /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
  assert.match(CODE, /APIFY_TOKEN/);
  assert.match(CODE, /PropertiesService/);
});

test('missing token writes a log row and throws', () => {
  const w = makeWorld({ props: { BOARDS: 'ashby:ramp' } });
  assert.throws(() => w.runDaily(), /APIFY_TOKEN/);
  const log = w.ss.sheets['Run log'];
  assert.ok(log);
  assert.equal(log.cells[0][0], 'time');
  assert.equal(log.cells[1][1], 'error');
  assert.match(String(log.cells[1][3]), /APIFY_TOKEN/);
  assert.equal(w.fetches.length, 0);
});

test('missing BOARDS writes a log row and throws, no fetch', () => {
  const w = makeWorld({ props: { APIFY_TOKEN: 'test-token' } });
  assert.throws(() => w.runDaily(), /BOARDS/);
  assert.equal(w.fetches.length, 0);
  assert.equal(w.ss.sheets['Run log'].cells[1][1], 'error');
});

test('successful run appends only real jobs and records the request shape', () => {
  const w = makeWorld({
    props: {
      APIFY_TOKEN: 'test-token',
      BOARDS: 'https://boards.greenhouse.io/airbnb, lever:spotify, ashby:ramp',
    },
  });
  w.runDaily();
  assert.equal(w.fetches.length, 1);
  const call = w.fetches[0];
  assert.match(call.url, /\/acts\/fecklesspleb~ats-jobs-scraper\/run-sync-get-dataset-items\?/);
  assert.match(call.url, /maxTotalChargeUsd=0\.1\b/);
  assert.match(call.url, /timeout=280/);
  assert.match(call.url, /clean=true/);
  assert.equal(call.options.headers.Authorization, 'Bearer test-token');
  assert.ok(!call.url.includes('test-token'));
  const body = JSON.parse(call.options.payload);
  assert.deepEqual(body.companies, [
    'https://boards.greenhouse.io/airbnb',
    'lever:spotify',
    'ashby:ramp',
  ]);
  assert.equal(body.onlyNewJobs, true);
  assert.equal(body.includeDescription, false);
  assert.equal(body.maxTotalJobs, 50);
  assert.equal(body.stateStoreName, 'ats-jobs-google-sheet');

  const jobs = w.ss.sheets.Jobs.cells;
  assert.deepEqual(jobs[0], [
    'uid', 'ats', 'board', 'company', 'title', 'location',
    'workplace', 'posted_at', 'salary', 'url', 'first_seen',
  ]);
  assert.equal(jobs.length, 3); // header + 2 jobs; unresolved/expired/status dropped
  assert.equal(jobs[1][0], 'ashby:ramp:34413f8d-26bf-4bbc-8ade-eb309a0e2245');
  assert.equal(jobs[1][3], 'Ramp');
  assert.equal(jobs[1][8], '$211.4K – $290.6K • Offers Equity');
  assert.equal(jobs[2][0], 'workday:nvidia:JR2001996');
  assert.equal(jobs[2][3], 'NVIDIA USA');
  const log = w.ss.sheets['Run log'].cells;
  assert.equal(log[1][1], 'ok');
  assert.equal(log[1][2], 2);
  assert.ok(!JSON.stringify(log).includes('test-token'));
});

test('second run with the same jobs appends nothing (uid dedupe)', () => {
  const w = makeWorld({
    props: { APIFY_TOKEN: 'test-token', BOARDS: 'ashby:ramp' },
  });
  w.runDaily();
  w.runDaily();
  assert.equal(w.ss.sheets.Jobs.cells.length, 3); // still header + 2
  const log = w.ss.sheets['Run log'].cells;
  assert.equal(log[2][1], 'ok');
  assert.equal(log[2][2], 0);
});

test('HTTP 408 writes a log row and does not add jobs', () => {
  const w = makeWorld({
    props: { APIFY_TOKEN: 'test-token', BOARDS: 'ashby:ramp' },
    fetchImpl: () => httpResponse(408, ''),
  });
  assert.throws(() => w.runDaily(), /300 s/);
  assert.ok(!w.ss.sheets.Jobs || w.ss.sheets.Jobs.cells.length <= 1);
  assert.equal(w.ss.sheets['Run log'].cells[1][1], 'error');
});

test('HTTP 401 writes a log row and redacts the token if it appeared', () => {
  const w = makeWorld({
    props: { APIFY_TOKEN: 'test-token', BOARDS: 'ashby:ramp' },
    fetchImpl: () => httpResponse(401, 'invalid token test-token'),
  });
  assert.throws(() => w.runDaily(), /401/);
  const msg = String(w.ss.sheets['Run log'].cells[1][3]);
  assert.ok(!msg.includes('test-token'));
  assert.match(msg, /\[redacted\]|401/);
});

test('installTrigger replaces earlier runDaily triggers', () => {
  const w = makeWorld();
  w.installTrigger();
  w.installTrigger();
  assert.equal(w.triggers.length, 1);
  assert.equal(w.triggers[0].getHandlerFunction(), 'runDaily');
  assert.equal(w.triggers[0].daily, 1);
  assert.equal(w.triggers[0].hour, 8);
});

test('cap constant is 0.10 and actor id is the Store name', () => {
  const w = makeWorld();
  assert.equal(w.MAX_TOTAL_CHARGE_USD, 0.10);
  assert.equal(w.ACTOR_ID, 'fecklesspleb~ats-jobs-scraper');
  assert.equal(w.MAX_TOTAL_JOBS, 50);
});

console.log('\n' + passed + ' tests passed (local stubs only; no network)');
