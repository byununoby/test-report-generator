/*
 * Test Report 핵심 로직 (브라우저와 Node 테스트에서 공통 사용)
 * 네트워크 통신 코드가 없습니다. 모든 처리는 실행 중인 PC 안에서만 이루어집니다.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReportCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  // 보고서에 필요한 항목. test: 원본 열 제목에서 이 항목을 자동으로 찾는 규칙
  var FIELDS = [
    { key: 'orderNo',     label: 'Order No. (표지 PO No.)',   test: /order|오더/ },
    { key: 'salesNo',     label: '수주번호 (표지 Order No.)',  test: /수주|sales/,              optional: true },
    { key: 'batchNo',     label: 'Batch No. (Drum No.)',      test: /drum|batch|드럼|배치/ },
    { key: 'qty',         label: 'Quantity (m)',              test: /quantity|수량/,           numeric: true, optional: true },
    { key: 'otdr',        label: 'OTDR (m)',                  test: /otdr/,                    numeric: true },
    { key: 'fiberNo',     label: 'Fiber No. (없으면 1)',       test: /fiber|심선/,              optional: true },
    { key: 'odA',         label: 'Cable O.D. 측정값 1',        test: /cable\s*o\.?\s*d|외경/,    numeric: true },
    { key: 'odB',         label: 'Cable O.D. 측정값 2',        test: /cable\s*o\.?\s*d|외경/,    numeric: true },
    { key: 'tightBuffer', label: 'Tight Buffer O.D.',         test: /tight|타이트/,             numeric: true },
    { key: 'att1310',     label: 'Attenuation 1310',          test: /1310/,                    numeric: true },
    { key: 'att1550',     label: 'Attenuation 1550',          test: /1550/,                    numeric: true },
    { key: 'marking',     label: 'Marking (인쇄간격)',         test: /marking|마킹|인쇄/ }
  ];
  // 열 제목 없이 데이터만 붙여넣었을 때 가정하는 순서 (입력 양식 순서)
  var DEFAULT_ORDER = ['orderNo', 'batchNo', 'otdr', 'fiberNo', 'odA', 'odB', 'tightBuffer', 'att1310', 'att1550', 'marking'];

  function norm(s) {
    return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  }
  function clean(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var t = String(v == null ? '' : v).replace(/,/g, '').trim();
    if (t === '' || isNaN(Number(t))) return null;
    return Number(t);
  }
  function numCount(row) { return row.filter(function (v) { return toNum(v) !== null; }).length; }

  // 구분자 텍스트 파서. 엑셀 복사본(탭)과 CSV(쉼표) 공용. 따옴표 안의 줄바꿈 처리
  function parseDelimited(text, delim) {
    var rows = [], row = [], cur = '', inQ = false;
    text = String(text).replace(/^\uFEFF/, '');
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') { cur += '"'; i++; } else inQ = false;
        } else cur += c;
      } else if (c === '"' && cur === '') inQ = true;
      else if (c === delim) { row.push(cur); cur = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cur); rows.push(row); row = []; cur = '';
      } else cur += c;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ''; }); });
  }
  function parseCSV(text) { return parseDelimited(text, ','); }
  function parseTSV(text) { return parseDelimited(text, '\t'); }

  function colName(i) {
    var s = ''; i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  // 붙여넣은 표 분석: 제목 행(여러 줄, 병합 셀 포함)과 데이터 행 분리
  function analyze(table) {
    if (!table || !table.length) throw new Error('데이터가 비어 있습니다.');
    var width = table.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var t = table.map(function (r) { var a = r.slice(); while (a.length < width) a.push(''); return a; });
    var dataStart = -1;
    for (var i = 0; i < t.length; i++) if (numCount(t[i]) >= 4) { dataStart = i; break; }
    if (dataStart === -1) throw new Error('숫자 데이터 행을 찾지 못했습니다. 측정값이 들어 있는 범위까지 복사했는지 확인하세요.');

    var H = t.slice(Math.max(0, dataStart - 4), dataStart).map(function (r) { return r.map(clean); });
    // 병합 셀 복원: 제목 칸이 완전히 비어 있으면 왼쪽 칸 제목을 이어받음
    for (var c = 1; c < width; c++) {
      var allEmpty = H.every(function (r) { return !r[c]; });
      if (allEmpty) { H.forEach(function (r) { r[c] = r[c - 1]; }); continue; }
      for (var h = 0; h < H.length; h++) {
        var below = H.slice(h + 1).some(function (r) { return r[c]; });
        var aboveEmpty = h === 0 || !H[h - 1][c];
        if (!H[h][c] && below && aboveEmpty && H[h][c - 1]) H[h][c] = H[h][c - 1];
      }
    }
    var labels = [];
    for (var k = 0; k < width; k++) {
      var parts = [];
      H.forEach(function (r) { if (r[k] && parts.indexOf(r[k]) === -1) parts.push(r[k]); });
      labels.push(parts.join(' '));
    }
    var data = t.slice(dataStart).filter(function (r) { return numCount(r) >= 3; });
    var samples = labels.map(function (_, k) {
      for (var j = 0; j < data.length; j++) if (clean(data[j][k])) return clean(data[j][k]);
      return '';
    });
    return { labels: labels, data: data, width: width, samples: samples, hasHeader: H.length > 0 };
  }

  // 열 제목으로 항목 자동 연결. 못 찾으면 -1
  function guessMapping(an) {
    var used = {}, map = {}, hits = 0;
    FIELDS.forEach(function (f) {
      map[f.key] = -1;
      for (var c = 0; c < an.width; c++) {
        if (used[c]) continue;
        if (f.test.test(norm(an.labels[c]))) { map[f.key] = c; used[c] = true; hits++; break; }
      }
    });
    // 측정값 2가 제목 없는 바로 옆 칸인 경우
    if (map.odA >= 0 && map.odB === -1 && !an.labels[map.odA + 1] && !used[map.odA + 1]) map.odB = map.odA + 1;
    if (hits < 3 && an.width >= DEFAULT_ORDER.length) {
      DEFAULT_ORDER.forEach(function (k, i) { map[k] = i; });
    }
    return map;
  }

  // 연결 정보대로 보고서 행 생성. 병합 셀(빈 칸)은 위 값으로 채움
  function buildRows(an, map) {
    var prev = {};
    var rows = an.data.map(function (r) {
      var o = {};
      FIELDS.forEach(function (f) {
        var v = map[f.key] >= 0 ? r[map[f.key]] : '';
        o[f.key] = f.numeric ? toNum(v) : clean(v);
      });
      ['orderNo', 'salesNo'].forEach(function (k) { if (!o[k] && prev[k]) o[k] = prev[k]; prev[k] = o[k]; });
      if (!o.fiberNo) o.fiberNo = '1';
      var od = [o.odA, o.odB].filter(function (x) { return x != null; });
      o.odMax = od.length ? Math.max.apply(null, od) : null;
      o.odMin = od.length ? Math.min.apply(null, od) : null;
      return o;
    });
    return rows.filter(function (o) { return o.otdr != null || o.att1310 != null || o.att1550 != null || o.odMax != null; });
  }

  // 한 번에: 표 → 행
  function mapRows(table) {
    var an = analyze(table), map = guessMapping(an);
    var rows = buildRows(an, map);
    var missing = FIELDS.filter(function (f) { return !f.optional && map[f.key] < 0; }).map(function (f) { return f.label; });
    return { rows: rows, missing: missing, mapping: map, analysis: an };
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
      marking:     (spec.marking && toNum(row.marking) !== null)
                     ? inRange(toNum(row.marking), spec.marking)
                     : passMarks.indexOf(norm(row.marking)) !== -1
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
    var hasQty = rows.some(function (r) { return r.qty != null; });
    var totalQty = hasQty ? rows.reduce(function (a, r) { return a + (r.qty || 0); }, 0) : totalLength;
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
      totalQty: totalQty,
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
    var f = function (n) {
      var dec = (String(n).split('.')[1] || '').length;
      return s.digits != null ? Number(n).toFixed(Math.max(s.digits, dec)) : String(n);
    };
    if (s.min != null && s.max != null) return f(s.min) + ' ~ ' + f(s.max);
    if (s.max != null) return '≤ ' + f(s.max);
    if (s.min != null) return '≥ ' + f(s.min);
    return '';
  }

  // 수주번호 2026-S50-0144-001-001 → 2026-S50-0144
  function salesPrefix(v) {
    var p = String(v || '').split('-');
    return p.length > 3 ? p.slice(0, 3).join('-') : String(v || '');
  }
  // 2026-07-13 → 07/13/26
  function usDate(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? m[2] + '/' + m[3] + '/' + m[1].slice(2) : (iso || '');
  }

  // 문서번호: TR-YYYYMMDD-오더번호
  function docNumber(dateStr, orderNo) {
    var d = String(dateStr || '').replace(/-/g, '');
    var o = String(orderNo || '').replace(/[^0-9A-Za-z]/g, '');
    return 'TR-' + (d || '00000000') + (o ? '-' + o : '');
  }

  return {
    FIELDS: FIELDS, parseCSV: parseCSV, parseTSV: parseTSV, analyze: analyze, guessMapping: guessMapping,
    buildRows: buildRows, colName: colName, mapRows: mapRows, judgeRow: judgeRow,
    evaluate: evaluate, stats: stats, salesPrefix: salesPrefix, usDate: usDate, specText: specText, docNumber: docNumber
  };
});
