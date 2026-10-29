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
  assert.strictEqual(ev.batches.length, 12);
  assert.strictEqual(ev.batches[0], '01001');  // 앞자리 0 유지
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

test('숫자 데이터가 없으면 명확한 오류를 낸다', () => {
  assert.throws(() => core.mapRows([['a', 'b'], ['1', '2']]), /숫자 데이터 행을 찾지 못했습니다/);
});

// 실제 원본과 같은 구조(가짜 값): 3줄 제목, 병합 셀, 줄바꿈 제목, 외경 2칸 + 두께 2칸, 추가 열
const ORIGINAL_LIKE = [
  ['No.', 'ORDER NO.', '수주 번호', 'LOT NO.', 'Drum No.', 'Quantity(m)', 'OTDR(m)', '"Cable O.D.\n(mm)"', '', '', '', '"Tight\nBuffer"', 'Attenuation(dB/Km)', '', '"인쇄 간격\n1000mm"', 'Pallet No.', '검사원'],
  ['', '', '', '', '', '', '', '', '', '', '', '', '1310', '1550', '', '', ''],
  ['', '', '', '', '', '', '', '2.80 ~ 3.00', '', '0.70 ~ 0.80', '', '850 ~ 950', 'Max. 0.50', 'Max. 0.40', '990 ~ 1010', '', ''],
  ['1', 'X1001', 'S-001', 'LOT-A01', '00101', '2,000', '2,011', '2.90', '2.95', '0.72', '0.75', '890', '0.330', '0.190', '1004', '1', '검사A'],
  ['2', '', 'S-002', 'LOT-A02', '00102', '2,000', '2,013', '2.88', '2.93', '0.71', '0.74', '900', '0.331', '0.191', '1003', '', '검사A'],
  ['3', '', 'S-003', 'LOT-A03', '00103', '2,000', '2,012', '2.91', '3.05', '0.72', '0.76', '880', '0.329', '0.189', '1020', '', '검사B']
].map(r => r.join('\t')).join('\r\n').replace(/\\n/g, '\n');

test('원본 양식 구조(여러 줄 제목, 병합 셀)를 자동 인식한다', () => {
  const { rows, mapping, analysis } = core.mapRows(core.parseTSV(ORIGINAL_LIKE));
  assert.strictEqual(rows.length, 3);
  assert.strictEqual(analysis.labels[mapping.odA].startsWith('Cable O.D.'), true);
  assert.strictEqual(mapping.odA, 7);
  assert.strictEqual(mapping.odB, 8);
  assert.strictEqual(mapping.batchNo, 4);   // Drum No.
  assert.strictEqual(mapping.salesNo, 2);   // 수주 번호
  assert.strictEqual(mapping.qty, 5);       // Quantity(m)
  assert.strictEqual(mapping.fiberNo, -1);  // 원본에 없음
  assert.strictEqual(rows[0].fiberNo, '1'); // 기본값
  assert.strictEqual(mapping.att1310, 12);
  assert.strictEqual(mapping.att1550, 13);
  assert.strictEqual(mapping.marking, 14);
  assert.strictEqual(mapping.tightBuffer, 11);
  assert.strictEqual(rows[1].orderNo, 'X1001'); // 병합된 ORDER NO. 채움
  assert.strictEqual(rows[0].otdr, 2011);       // 천 단위 쉼표 처리
  assert.strictEqual(rows[0].odMax, 2.95);
  assert.strictEqual(rows[0].odMin, 2.90);
});

test('원본 양식: 외경 초과와 인쇄간격 초과를 잡아낸다', () => {
  const { rows } = core.mapRows(core.parseTSV(ORIGINAL_LIKE));
  const ev = core.evaluate(rows, 'Type 1', config);
  assert.strictEqual(ev.results[0].pass, true);
  assert.strictEqual(ev.results[2].checks.odMax, false);
  assert.strictEqual(ev.results[2].checks.marking, false);
  assert.strictEqual(ev.qaResult, 'FAIL');
});

test('수주번호 앞부분과 미국식 날짜 변환', () => {
  assert.strictEqual(core.salesPrefix('2026-S50-0144-001-001'), '2026-S50-0144');
  assert.strictEqual(core.usDate('2026-07-13'), '07/13/26');
});

test('열 제목 없이 데이터만 붙여넣으면 입력 양식 순서로 읽는다', () => {
  const body = csv.trim().split(/\r?\n/).slice(1).join('\n');
  const { rows } = core.mapRows(core.parseCSV(body));
  assert.strictEqual(rows.length, 12);
  assert.strictEqual(rows[0].batchNo, '01001');
  assert.strictEqual(rows[0].fiberNo, '1');
});

test('핵심 로직에 외부 통신 코드가 없다 (보안 점검)', () => {
  const files = ['js/report-core.js', 'js/app.js', 'config/specs.js'];
  files.forEach(f => {
    const src = fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
    assert.ok(!/\bfetch\s*\(|XMLHttpRequest|sendBeacon|WebSocket/.test(src), f + '에 통신 코드가 있음');
  });
});
