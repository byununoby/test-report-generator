/*
 * Test Report 핵심 로직 (브라우저와 Node 테스트에서 공통 사용)
 * 네트워크 통신 코드가 없습니다. 모든 처리는 실행 중인 PC 안에서만 이루어집니다.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReportCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  // 입력 데이터 컬럼 정의. aliases에 있는 이름이면 헤더로 인식
  var COLUMNS = [
    { key: 'orderNo',     label: 'Order No.',          aliases: ['order no', 'order no.', 'customer order no', 'customer order no.', 'corning order no', 'corning order no.', '오더번호'] },
    { key: 'batchNo',     label: 'Batch No.',          aliases: ['batch no', 'batch no.', 'batch', '배치번호', '로트번호', 'lot no', 'lot no.'] },
    { key: 'otdr',        label: 'OTDR (m)',           aliases: ['otdr', 'otdr (m)', 'otdr(m)', 'length', '길이'], numeric: true },
    { key: 'fiberNo',     label: 'Fiber No.',          aliases: ['fiber no', 'fiber no.', 'fiber', '심선번호'] },
    { key: 'odMax',       label: 'Cable O.D. Max',     aliases: ['cable od max', 'cable o.d. max', 'od max', 'odmax', '외경 max', '외경최대'], numeric: true },
    { key: 'odMin',       label: 'Cable O.D. Min',     aliases: ['cable od min', 'cable o.d. min', 'od min', 'odmin', '외경 min', '외경최소'], numeric: true },
    { key: 'tightBuffer', label: 'Tight Buffer O.D.',  aliases: ['tight buffer', 'tight buffer od', 'tight buffer o.d.', 'tight buffer o.d. (um)', 'tb od', '타이트버퍼'], numeric: true },
    { key: 'att1310',     label: 'Att. 1310',          aliases: ['att1310', 'att 1310', 'ose1310', 'attenuation 1310', '1310', '손실1310'], numeric: true },
    { key: 'att1550',     label: 'Att. 1550',          aliases: ['att1550', 'att 1550', 'ose1550', 'attenuation 1550', '1550', '손실1550'], numeric: true },
    { key: 'marking',     label: 'Marking',            aliases: ['marking', '마킹', '인쇄'] }
  ];

  function norm(s) {
    return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  }

  // 간단한 CSV 파서 (따옴표, 쉼표 포함 값 처리)
  function parseCSV(text) {
    var rows = [], row = [], cur = '', inQ = false;
    text = String(text).replace(/^\uFEFF/, '');
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') { cur += '"'; i++; } else inQ = false;
        } else cur += c;
      } else if (c === '"') inQ = true;
      else if (c === ',') { row.push(cur); cur = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cur); rows.push(row); row = []; cur = '';
      } else cur += c;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ''; }); });
  }

  // 2차원 배열(헤더 포함)을 표준 행 객체 배열로 변환
  function mapRows(table) {
    if (!table || !table.length) throw new Error('데이터가 비어 있습니다.');
    var headerIdx = -1, map = {};
    for (var h = 0; h < Math.min(table.length, 10); h++) {
      var m = {}, hits = 0;
      table[h].forEach(function (cell, ci) {
        var n = norm(cell);
        COLUMNS.forEach(function (col) {
          if (m[col.key] === undefined && col.aliases.indexOf(n) !== -1) { m[col.key] = ci; hits++; }
        });
      });
      if (hits >= 3) { headerIdx = h; map = m; break; }
    }
    if (headerIdx === -1) throw new Error('헤더를 찾지 못했습니다. 입력 양식(템플릿)의 컬럼명을 사용해 주세요.');
    var missing = COLUMNS.filter(function (c) { return map[c.key] === undefined; }).map(function (c) { return c.label; });

    var rows = [];
    for (var r = headerIdx + 1; r < table.length; r++) {
      var src = table[r], obj = {}, empty = true;
      COLUMNS.forEach(function (col) {
        var v = map[col.key] === undefined ? '' : src[map[col.key]];
        if (v != null && String(v).trim() !== '') empty = false;
        if (col.numeric) {
          var num = parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
          obj[col.key] = isNaN(num) ? null : num;
        } else obj[col.key] = v == null ? '' : String(v).trim();
      });
      // 숫자 항목이 하나도 없는 행(규격행, 소제목행 등)은 제외
      var hasNum = COLUMNS.some(function (col) { return col.numeric && obj[col.key] != null; });
      if (!empty && hasNum) rows.push(obj);
    }
    return { rows: rows, missing: missing };
  }

  function inRange(v, s) {
    if (v == null) return false;
    if (s.min != null && v < s.min) return false;
    if (s.max != null && v > s.max) return false;
    return true;
  }

  // 한 행 판정. 항목별 true(합격) / false(불합격)
  function judgeRow(row, spec, config) {
    var passMarks = (config.markingPassValues || ['OK']).map(norm);
    var c = {
      otdr:        row.otdr != null && row.otdr > 0,
      odMax:       inRange(row.odMax, spec.cableOD),
      odMin:       inRange(row.odMin, spec.cableOD),
      tightBuffer: inRange(row.tightBuffer, spec.tightBuffer),
      att1310:     inRange(row.att1310, spec.att1310),
      att1550:     inRange(row.att1550, spec.att1550),
      marking:     passMarks.indexOf(norm(row.marking)) !== -1
    };
    var pass = Object.keys(c).every(function (k) { return c[k]; });
    return { checks: c, pass: pass };
  }

  function stats(values) {
    var v = values.filter(function (x) { return x != null; });
    if (!v.length) return null;
    var sum = v.reduce(function (a, b) { return a + b; }, 0), avg = sum / v.length;
    var sd = v.length > 1 ? Math.sqrt(v.reduce(function (a, b) { return a + (b - avg) * (b - avg); }, 0) / (v.length - 1)) : 0;
    return { n: v.length, avg: avg, min: Math.min.apply(null, v), max: Math.max.apply(null, v), sd: sd };
  }

  // 전체 판정 및 요약
  function evaluate(rows, cableTypeId, config) {
    var type = config.cableTypes.filter(function (t) { return t.id === cableTypeId; })[0];
    if (!type) throw new Error('제품 종류를 선택해 주세요.');
    var spec = config.specs[type.spec];
    if (!spec) throw new Error(type.spec + ' 규격이 설정 파일에 없습니다.');

    var results = rows.map(function (r) { return judgeRow(r, spec, config); });
    var failCount = results.filter(function (x) { return !x.pass; }).length;
    var totalLength = rows.reduce(function (a, r) { return a + (r.otdr || 0); }, 0);
    var batches = [];
    rows.forEach(function (r) { if (r.batchNo && batches.indexOf(r.batchNo) === -1) batches.push(r.batchNo); });

    return {
      type: type,
      spec: spec,
      results: results,
      failCount: failCount,
      qaResult: rows.length && failCount === 0 ? 'PASS' : 'FAIL',
      drums: rows.length,
      totalLength: totalLength,
      batches: batches,
      stats: {
        otdr: stats(rows.map(function (r) { return r.otdr; })),
        odMax: stats(rows.map(function (r) { return r.odMax; })),
        odMin: stats(rows.map(function (r) { return r.odMin; })),
        tightBuffer: stats(rows.map(function (r) { return r.tightBuffer; })),
        att1310: stats(rows.map(function (r) { return r.att1310; })),
        att1550: stats(rows.map(function (r) { return r.att1550; }))
      }
    };
  }

  function specText(s) {
    if (!s) return '';
    var f = function (n) { return s.digits != null ? Number(n).toFixed(s.digits) : String(n); };
    if (s.min != null && s.max != null) return f(s.min) + ' ~ ' + f(s.max);
    if (s.max != null) return '≤ ' + f(s.max);
    if (s.min != null) return '≥ ' + f(s.min);
    return '';
  }

  // 문서번호: TR-YYYYMMDD-오더번호
  function docNumber(dateStr, orderNo) {
    var d = String(dateStr || '').replace(/-/g, '');
    var o = String(orderNo || '').replace(/[^0-9A-Za-z]/g, '');
    return 'TR-' + (d || '00000000') + (o ? '-' + o : '');
  }

  return {
    COLUMNS: COLUMNS, parseCSV: parseCSV, mapRows: mapRows, judgeRow: judgeRow,
    evaluate: evaluate, stats: stats, specText: specText, docNumber: docNumber
  };
});
