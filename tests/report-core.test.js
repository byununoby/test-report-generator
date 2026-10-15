// GitHub Actions에서 자동 실행되는 판정 로직 테스트 (node --test)
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const core = require('../js/report-core.js');
const config = require('../config/specs.js');

const csv = fs.readFileSync(path.join(__dirname, '..', 'sample-data', 'sample_test_data.csv'), 'utf8');

test('샘플 CSV의 헤더와 행을 정확히 읽는다', () => {
  const { rows, missing } = core.mapRows(core.parseCSV(csv));
  assert.strictEqual(rows.length, 12);
  assert.deepStrictEqual(missing, []);
  assert.strictEqual(rows[0].att1310, 0.379);
});

test('규격을 벗어난 1550nm 손실을 불합격으로 판정한다', () => {
  const { rows } = core.mapRows(core.parseCSV(csv));
  const ev = core.evaluate(rows, 'Type 1', config);
  assert.strictEqual(ev.failCount, 1);
  assert.strictEqual(ev.results[8].checks.att1550, false);
  assert.strictEqual(ev.qaResult, 'FAIL');
});

test('모든 값이 규격 안이면 PASS', () => {
  const { rows } = core.mapRows(core.parseCSV(csv));
  rows[8].att1550 = 0.25;
  const ev = core.evaluate(rows, 'Type 1', config);
  assert.strictEqual(ev.qaResult, 'PASS');
  assert.strictEqual(ev.drums, 12);
  assert.deepStrictEqual(ev.batches, ['B2610-01', 'B2610-02']);
});

test('규격 경계값은 합격으로 본다', () => {
  const spec = config.specs['2.9mm'];
  const r = core.judgeRow({ otdr: 1000, odMax: 3.0, odMin: 2.8, tightBuffer: 950, att1310: 0.5, att1550: 0.4, marking: 'ok' }, spec, config);
  assert.strictEqual(r.pass, true);
});

test('빈 값과 마킹 불량은 불합격', () => {
  const spec = config.specs['2.9mm'];
  const r = core.judgeRow({ otdr: null, odMax: 2.9, odMin: 2.9, tightBuffer: 900, att1310: 0.3, att1550: 0.2, marking: 'NG' }, spec, config);
  assert.strictEqual(r.checks.otdr, false);
  assert.strictEqual(r.checks.marking, false);
});

test('모든 제품 종류에 규격이 연결되어 있다', () => {
  config.cableTypes.forEach(t => assert.ok(config.specs[t.spec], t.id + ' 규격 없음'));
});

test('헤더를 못 찾으면 명확한 오류를 낸다', () => {
  assert.throws(() => core.mapRows([['a', 'b'], ['1', '2']]), /헤더를 찾지 못했습니다/);
});

test('핵심 로직에 외부 통신 코드가 없다 (보안 점검)', () => {
  const files = ['js/report-core.js', 'js/app.js', 'config/specs.js'];
  files.forEach(f => {
    const src = fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
    assert.ok(!/\bfetch\s*\(|XMLHttpRequest|sendBeacon|WebSocket/.test(src), f + '에 통신 코드가 있음');
  });
});
