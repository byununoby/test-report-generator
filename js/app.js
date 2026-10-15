/*
 * 화면 동작 코드
 * 파일은 브라우저 메모리에서만 읽고, 결과 파일도 브라우저에서 바로 만들어 내려받습니다.
 * 서버로 보내는 코드는 없습니다. (index.html의 보안 정책 connect-src 'none'으로 통신 자체를 차단)
 */
(function () {
  var C = window.APP_CONFIG, Core = window.ReportCore;
  var state = { rows: [], ev: null, fileName: '' };
  var $ = function (id) { return document.getElementById(id); };

  // ---------- 초기화 ----------
  function init() {
    if (typeof ExcelJS === 'undefined') {
      setStatus('엑셀 라이브러리(lib/exceljs.min.js)가 없습니다. CSV는 사용할 수 있지만 엑셀 불러오기와 저장은 안 됩니다. GitHub Actions의 오프라인 패키지(test-report-offline)를 받아서 사용하세요.', 'warn');
    }
    var sel = $('cableType');
    C.cableTypes.forEach(function (t) {
      var o = document.createElement('option');
      o.value = t.id; o.textContent = t.id + '  ' + t.name + ' (' + t.color + ')';
      sel.appendChild(o);
    });
    $('date').value = new Date().toISOString().slice(0, 10);
    sel.addEventListener('change', function () { $('color').value = typeOf().color; refresh(); });
    $('color').value = typeOf().color;

    ['date', 'orderNo', 'poNo', 'customer', 'color', 'quantity', 'batchNo', 'productName', 'writer', 'reviewer', 'approver']
      .forEach(function (id) { $(id).addEventListener('input', renderPreview); });

    // 붙여넣기: 파일 선택 없이 클립보드 텍스트로 데이터 입력 (파일 반입 보안 정책 회피 불필요)
    $('pasteBox').addEventListener('paste', function (e) {
      var text = (e.clipboardData || window.clipboardData).getData('text');
      if (!text) return;
      e.preventDefault();
      this.value = text;
      usePasted(text);
    });
    $('pasteBox').addEventListener('input', function () { if (this.value.trim()) usePasted(this.value); });
    $('copyBtn').addEventListener('click', copyResult);

    $('file').addEventListener('change', function (e) { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ''; });
    var drop = $('drop');
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) {
      e.preventDefault(); drop.classList.remove('over');
      if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
    });
    $('sampleBtn').addEventListener('click', function () { useTable(window.SAMPLE_TABLE, '시연용 샘플 데이터'); });
    $('templateBtn').addEventListener('click', downloadTemplate);
    $('excelBtn').addEventListener('click', exportExcel);
    $('printBtn').addEventListener('click', function () { window.print(); });
    $('resetBtn').addEventListener('click', function () {
      state = { rows: [], ev: null, fileName: '' };
      ['orderNo', 'quantity', 'batchNo', 'productName'].forEach(function (id) { $(id).value = ''; });
      $('pasteBox').value = '';
      setStatus('', ''); refresh();
    });
    refresh();
  }

  function typeOf() {
    var id = $('cableType').value;
    return C.cableTypes.filter(function (t) { return t.id === id; })[0];
  }

  function setStatus(msg, kind) {
    var s = $('status'); s.textContent = msg; s.className = 'status ' + (kind || '');
  }

  // ---------- 파일 읽기 (로컬 전용) ----------
  function loadFile(file) {
    var name = file.name.toLowerCase();
    var reader = new FileReader();
    reader.onerror = function () { setStatus('파일을 읽지 못했습니다. 파일이 열려 있으면 닫고 다시 시도하세요.', 'err'); };
    if (name.endsWith('.csv')) {
      reader.onload = function () { useTable(Core.parseCSV(reader.result), file.name); };
      reader.readAsText(file, 'utf-8');
    } else if (name.endsWith('.xlsx')) {
      if (noExcel()) return;
      reader.onload = function () {
        var wb = new ExcelJS.Workbook();
        wb.xlsx.load(reader.result).then(function () {
          // 시트가 여러 개면 데이터 헤더가 있는 첫 시트를 사용
          var tables = wb.worksheets.map(function (ws) {
            var table = [];
            ws.eachRow({ includeEmpty: false }, function (row) {
              var arr = [];
              for (var c = 1; c <= ws.columnCount; c++) arr.push(cellValue(row.getCell(c)));
              table.push(arr);
            });
            return table;
          });
          var pick = tables.filter(function (t) {
            try { return Core.mapRows(t).rows.length > 0; } catch (e) { return false; }
          })[0];
          useTable(pick || tables[0] || [], file.name);
        }).catch(function () { setStatus('엑셀 파일 형식을 읽지 못했습니다. .xlsx로 저장된 파일인지 확인하세요.', 'err'); });
      };
      reader.readAsArrayBuffer(file);
    } else {
      setStatus('.xlsx 또는 .csv 파일만 지원합니다. (.xls는 엑셀에서 .xlsx로 다시 저장)', 'err');
    }
  }

  function cellValue(cell) {
    var v = cell.value;
    if (v == null) return '';
    if (typeof v === 'object') {
      if (v.result !== undefined) return v.result;
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      return cell.text;
    }
    return v;
  }

  // 엑셀에서 복사한 텍스트(탭 구분) 처리
  function parseTSV(text) {
    return String(text).replace(/\r\n?/g, '\n').split('\n')
      .map(function (l) { return l.split('\t'); })
      .filter(function (r) { return r.some(function (v) { return v.trim() !== ''; }); });
  }

  // 표준 입력 양식 컬럼 순서 (컬럼명 없이 데이터만 붙여넣은 경우 이 순서로 간주)
  var DEFAULT_HEADER = ['Order No.', 'Batch No.', 'OTDR (m)', 'Fiber No.', 'Cable OD Max', 'Cable OD Min', 'Tight Buffer', 'Att1310', 'Att1550', 'Marking'];

  var pasteTimer;
  function usePasted(text) {
    clearTimeout(pasteTimer);
    pasteTimer = setTimeout(function () {
      var table = text.indexOf('\t') !== -1 ? parseTSV(text) : Core.parseCSV(text);
      try { Core.mapRows(table); }
      catch (e) {
        if (table.length && table[0].length >= 10) table = [DEFAULT_HEADER].concat(table);
      }
      useTable(table, '붙여넣은 데이터');
    }, 150);
  }

  function useTable(table, label) {
    try {
      var res = Core.mapRows(table);
      if (!res.rows.length) throw new Error('헤더 아래에 데이터 행이 없습니다.');
      state.rows = res.rows; state.fileName = label;
      $('orderNo').value = $('orderNo').value || res.rows[0].orderNo || '';
      $('quantity').value = ''; $('batchNo').value = '';
      var msg = label + ' 에서 ' + res.rows.length + '행을 읽었습니다.';
      if (res.missing.length) msg += ' 찾지 못한 컬럼: ' + res.missing.join(', ');
      setStatus(msg, res.missing.length ? 'warn' : 'ok');
      refresh();
    } catch (e) { setStatus(e.message, 'err'); }
  }

  // ---------- 판정 및 화면 갱신 ----------
  function refresh() {
    state.ev = state.rows.length ? Core.evaluate(state.rows, $('cableType').value, C) : null;
    var ev = state.ev;
    if (ev) {
      if (!$('quantity').dataset.touched) $('quantity').placeholder = ev.drums + ' D / ' + fmtInt(ev.totalLength) + ' M';
      $('batchNo').placeholder = ev.batches.join(', ');
      $('productName').placeholder = ev.type.name;
    } else {
      $('quantity').placeholder = '데이터를 불러오면 자동 계산';
      $('batchNo').placeholder = '데이터를 불러오면 자동 입력';
      $('productName').placeholder = typeOf().name;
    }
    $('excelBtn').disabled = !ev; $('printBtn').disabled = !ev; $('copyBtn').disabled = !ev;
    renderSpec(); renderSummary(); renderPreview();
  }

  function val(id) {
    var el = $(id);
    if (el.value.trim()) return el.value.trim();
    return el.dataset.auto ? el.placeholder.replace(/^데이터를.*$/, '') : '';
  }
  function fmtInt(n) { return Math.round(n).toLocaleString('en-US'); }
  function fmt(n, d) { return n == null ? '' : Number(n).toFixed(d); }

  function renderSpec() {
    var s = C.specs[typeOf().spec];
    $('specBox').innerHTML =
      '<span>Cable O.D. ' + Core.specText(s.cableOD) + ' mm</span>' +
      '<span>Tight Buffer ' + Core.specText(s.tightBuffer) + ' um</span>' +
      '<span>1310 ' + Core.specText(s.att1310) + '</span>' +
      '<span>1550 ' + Core.specText(s.att1550) + '</span>';
  }

  function renderSummary() {
    var ev = state.ev, box = $('summary');
    if (!ev) { box.hidden = true; return; }
    box.hidden = false;
    $('verdict').textContent = ev.qaResult;
    $('verdict').className = 'verdict ' + (ev.qaResult === 'PASS' ? 'pass' : 'fail');
    var fails = [];
    ev.results.forEach(function (r, i) {
      if (!r.pass) {
        var items = Object.keys(r.checks).filter(function (k) { return !r.checks[k]; }).map(labelOf);
        fails.push('<li>' + esc(state.rows[i].fiberNo || ('행 ' + (i + 1))) + ': ' + items.join(', ') + '</li>');
      }
    });
    $('verdictText').innerHTML = ev.failCount
      ? ev.drums + '건 중 ' + ev.failCount + '건 불합격<ul>' + fails.join('') + '</ul>'
      : ev.drums + '건 전체 합격, 총 ' + fmtInt(ev.totalLength) + ' m';
  }

  function labelOf(k) {
    return { otdr: 'OTDR', odMax: 'O.D. Max', odMin: 'O.D. Min', tightBuffer: 'Tight Buffer', att1310: '1310', att1550: '1550', marking: 'Marking' }[k];
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }

  // ---------- 미리보기 (인쇄/PDF와 동일) ----------
  function renderPreview() {
    var ev = state.ev, t = typeOf();
    var docNo = Core.docNumber($('date').value, val('orderNo'));
    var refRows = C.cableTypes.map(function (x) {
      return '<tr' + (x.id === t.id ? ' class="sel"' : '') + '><td>' + x.id + '</td><td class="l">' + esc(x.name) + '</td><td>' + x.fiber + '</td><td>' + x.color + '</td></tr>';
    }).join('');

    var cover =
      '<section class="paper cover">' +
      '<div class="docno">' + esc(docNo) + '</div>' +
      '<div class="cover-grid"><div>' +
      '<h2 class="cover-title">TEST REPORT</h2>' +
      '<table class="cov">' +
      row2('Date', $('date').value) + row2('Order No.', val('orderNo')) + row2('PO No.', val('poNo')) +
      row2('Customer', val('customer')) + row2('Cable Type', t.id + ' / ' + t.name) + row2('COLOR', val('color')) +
      row2('Quantity', val('quantity')) + '<tr class="tall"><th>Batch No.</th><td>' + esc(val('batchNo')) + '</td></tr>' +
      '</table></div>' +
      '<div class="ref"><p class="ref-title">※ 제품명 참고</p><table class="reftab"><tr><th></th><th>Cable Type</th><th>FIBER</th><th>COLOR</th></tr>' + refRows + '</table></div>' +
      '</div>' + signBox() + '</section>';

    var body = '';
    if (ev) {
      body = state.rows.map(function (r, i) {
        var c = ev.results[i].checks;
        return '<tr>' + td(r.orderNo) + td(r.batchNo) + td(r.otdr, !c.otdr) + td(r.fiberNo) +
          td(fmt(r.odMax, 2), !c.odMax) + td(fmt(r.odMin, 2), !c.odMin) + td(r.tightBuffer, !c.tightBuffer) +
          td(fmt(r.att1310, 3), !c.att1310) + td(fmt(r.att1550, 3), !c.att1550) + td(r.marking, !c.marking) + '</tr>';
      }).join('');
    } else {
      for (var i = 0; i < 8; i++) body += '<tr>' + new Array(11).join('<td>&nbsp;</td>') + '</tr>';
    }
    var s = ev ? ev.spec : C.specs[t.spec];
    var specRow = '<tr class="specrow"><td colspan="2">Spec.</td><td></td><td></td><td colspan="2">' + Core.specText(s.cableOD) +
      '</td><td>' + Core.specText(s.tightBuffer) + '</td><td>' + Core.specText(s.att1310) + '</td><td>' + Core.specText(s.att1550) + '</td><td>OK</td></tr>';

    var report =
      '<section class="paper report">' +
      '<h2 class="rep-title">Test&nbsp;&nbsp;&nbsp;Report</h2>' +
      '<dl class="rep-head">' +
      dl('Product Name', val('productName')) + dl('Color', val('color')) + dl('Inspection Date', $('date').value) +
      dl('Quantity(D/M)', val('quantity')) +
      '<dt>QA RESULTS</dt><dd class="' + (ev ? ev.qaResult.toLowerCase() : '') + '">' + (ev ? ev.qaResult : '') + '</dd></dl>' +
      '<div class="tablewrap"><table class="rep">' +
      '<thead><tr><th rowspan="2">' + esc(C.orderNoLabel) + '</th><th rowspan="2">Batch No.</th><th rowspan="2">OTDR (m)</th><th rowspan="2">FIBER NO.</th>' +
      '<th colspan="2">Cable O.D. (mm)</th><th rowspan="2">Tight Buffer O.D. (um)</th><th colspan="2">Attenuation (dB/Km)</th><th rowspan="2">marking</th></tr>' +
      '<tr><th>Max.</th><th>Min.</th><th>ose1310</th><th>ose1550</th></tr>' + specRow + '</thead>' +
      '<tbody>' + body + '</tbody></table></div></section>';

    $('preview').innerHTML = cover + report;
  }

  function row2(k, v) { return '<tr><th>' + k + '</th><td>' + esc(v) + '</td></tr>'; }
  function dl(k, v) { return '<dt>' + k + '</dt><dd>' + esc(v) + '</dd>'; }
  function td(v, bad) { return '<td' + (bad ? ' class="bad"' : '') + '>' + esc(v == null ? '' : v) + '</td>'; }
  function signBox() {
    return '<table class="sign"><tr><th>작 성</th><th>검 토</th><th>승 인</th></tr><tr><td>' +
      esc($('writer').value) + '</td><td>' + esc($('reviewer').value) + '</td><td>' + esc($('approver').value) + '</td></tr></table>';
  }

  // ---------- 엑셀 출력 ----------
  var thin = { style: 'thin', color: { argb: 'FF000000' } };
  var box = { top: thin, left: thin, bottom: thin, right: thin };
  var YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  var GREY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  var RED = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
  var center = { horizontal: 'center', vertical: 'middle', wrapText: true };

  function noExcel() {
    if (typeof ExcelJS !== 'undefined') return false;
    setStatus('엑셀 라이브러리가 없어 이 기능을 쓸 수 없습니다. 오프라인 패키지(test-report-offline)를 받아서 사용하세요.', 'err');
    return true;
  }

  function exportExcel() {
    var ev = state.ev, t = typeOf();
    if (!ev || noExcel()) return;
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Test Report Generator';

    // Cover
    var cv = wb.addWorksheet('Cover', { pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 } });
    cv.columns = [{ width: 3 }, { width: 16 }, { width: 28 }, { width: 6 }, { width: 9 }, { width: 22 }, { width: 11 }, { width: 8 }];
    cv.mergeCells('B3:C3');
    cv.getCell('B3').value = 'TEST REPORT';
    cv.getCell('B3').font = { name: 'Times New Roman', size: 26, bold: true };
    cv.getCell('B3').alignment = center;
    cv.getRow(3).height = 40;
    cv.getCell('B1').value = Core.docNumber($('date').value, val('orderNo'));
    cv.getCell('B1').font = { size: 9, color: { argb: 'FF666666' } };
    var items = [['Date', $('date').value], ['Order No.', val('orderNo')], ['PO No.', val('poNo')], ['Customer', val('customer')],
      ['Cable Type', t.id + ' / ' + t.name], ['COLOR', val('color')], ['Quantity', val('quantity')], ['Batch No.', val('batchNo')]];
    items.forEach(function (it, i) {
      var r = 6 + i;
      var a = cv.getCell('B' + r), b = cv.getCell('C' + r);
      a.value = it[0]; b.value = it[1];
      a.fill = GREY; a.font = { bold: true }; a.alignment = center; a.border = box;
      b.alignment = { vertical: 'middle', wrapText: true }; b.border = box;
      cv.getRow(r).height = it[0] === 'Batch No.' ? 42 : 20;
    });
    cv.getCell('E5').value = '※ 제품명 참고'; cv.getCell('E5').font = { bold: true, size: 9 };
    ['', 'Cable Type', 'FIBER', 'COLOR'].forEach(function (h, i) {
      var c = cv.getRow(6).getCell(5 + i); c.value = h; c.font = { bold: true, size: 9 }; c.alignment = center; c.border = box;
    });
    C.cableTypes.forEach(function (x, i) {
      var row = cv.getRow(7 + i);
      [x.id, x.name, x.fiber, x.color].forEach(function (v, j) {
        var c = row.getCell(5 + j); c.value = v; c.font = { size: 9, bold: x.id === t.id }; c.border = box;
        c.alignment = j === 1 ? { vertical: 'middle' } : center;
        if (x.id === t.id) c.fill = YELLOW;
      });
    });
    var sr = 22;
    ['작 성', '검 토', '승 인'].forEach(function (h, i) {
      var col = String.fromCharCode(69 + i); // E F G
      cv.getCell(col + sr).value = h; cv.getCell(col + sr).alignment = center; cv.getCell(col + sr).border = box; cv.getCell(col + sr).font = { bold: true };
      var c = cv.getCell(col + (sr + 1));
      c.value = [$('writer').value, $('reviewer').value, $('approver').value][i];
      c.alignment = center; c.border = box;
    });
    cv.getRow(sr + 1).height = 45;

    // Test Report
    var ws = wb.addWorksheet('Test Report', { pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = [16, 14, 11, 9, 10, 10, 15, 12, 12, 12].map(function (w) { return { width: w }; });
    ws.mergeCells('A1:J1');
    ws.getCell('A1').value = 'Test   Report';
    ws.getCell('A1').font = { size: 22, bold: true, underline: true };
    ws.getCell('A1').alignment = center; ws.getRow(1).height = 36;
    var head = [['Product Name', val('productName')], ['Color', val('color')], ['Inspection Date', $('date').value],
      ['Quantity(D/M)', val('quantity')], ['QA RESULTS', ev.qaResult]];
    head.forEach(function (h, i) {
      var r = 3 + i;
      ws.mergeCells('D' + r + ':E' + r); ws.mergeCells('F' + r + ':H' + r);
      ws.getCell('D' + r).value = h[0]; ws.getCell('D' + r).font = { bold: true, italic: true };
      var v = ws.getCell('F' + r); v.value = h[1]; v.border = { bottom: thin };
      if (h[0] === 'QA RESULTS') v.font = { bold: true, color: { argb: ev.qaResult === 'PASS' ? 'FF1F7A3A' : 'FFC00000' } };
    });
    var h1 = 9, h2 = 10;
    var top = [C.orderNoLabel, 'Batch No.', 'OTDR (m)', 'FIBER NO.', 'Cable O.D. (mm)', '', 'Tight Buffer O.D. (um)', 'Attenuation(dB/Km)', '', 'marking'];
    var sub = ['', '', '', '', 'Max.', 'Min.', '', 'ose1310', 'ose1550', ''];
    top.forEach(function (v, i) { ws.getRow(h1).getCell(i + 1).value = v; });
    sub.forEach(function (v, i) { if (v) ws.getRow(h2).getCell(i + 1).value = v; });
    ['A', 'B', 'C', 'D', 'G', 'J'].forEach(function (c) { ws.mergeCells(c + h1 + ':' + c + h2); });
    ws.mergeCells('E' + h1 + ':F' + h1); ws.mergeCells('H' + h1 + ':I' + h1);
    [h1, h2].forEach(function (r) {
      for (var c = 1; c <= 10; c++) { var cell = ws.getRow(r).getCell(c); cell.fill = YELLOW; cell.alignment = center; cell.border = box; }
    });
    var s = ev.spec;
    var spec = ['Spec.', '', '', '', Core.specText(s.cableOD), Core.specText(s.cableOD), Core.specText(s.tightBuffer), Core.specText(s.att1310), Core.specText(s.att1550), 'OK'];
    var specR = ws.getRow(11);
    spec.forEach(function (v, i) { var c = specR.getCell(i + 1); c.value = v; c.font = { size: 9, italic: true }; c.alignment = center; c.border = box; });

    var keys = ['orderNo', 'batchNo', 'otdr', 'fiberNo', 'odMax', 'odMin', 'tightBuffer', 'att1310', 'att1550', 'marking'];
    var numFmt = { odMax: '0.00', odMin: '0.00', att1310: '0.000', att1550: '0.000', otdr: '#,##0', tightBuffer: '0' };
    state.rows.forEach(function (r, i) {
      var row = ws.getRow(12 + i), checks = ev.results[i].checks;
      keys.forEach(function (k, j) {
        var c = row.getCell(j + 1);
        c.value = r[k] == null ? '' : r[k];
        if (numFmt[k]) c.numFmt = numFmt[k];
        c.alignment = center; c.border = box;
        if (checks[k] === false) { c.fill = RED; c.font = { bold: true, color: { argb: 'FFC00000' } }; }
      });
    });

    // 통계 시트
    var st = wb.addWorksheet('Statistics');
    st.columns = [{ width: 22 }, { width: 8 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }];
    st.addRow(['항목', 'n', '평균', '최소', '최대', '표준편차', '규격']).eachCell(function (c) { c.font = { bold: true }; c.fill = GREY; c.border = box; c.alignment = center; });
    [['OTDR (m)', 'otdr', null, 0], ['Cable O.D. Max (mm)', 'odMax', s.cableOD, 3], ['Cable O.D. Min (mm)', 'odMin', s.cableOD, 3],
      ['Tight Buffer O.D. (um)', 'tightBuffer', s.tightBuffer, 1], ['Att. 1310 (dB/km)', 'att1310', s.att1310, 4], ['Att. 1550 (dB/km)', 'att1550', s.att1550, 4]]
      .forEach(function (x) {
        var m = ev.stats[x[1]]; if (!m) return;
        var p = Math.pow(10, x[3]), rd = function (n) { return Math.round(n * p) / p; };
        st.addRow([x[0], m.n, rd(m.avg), rd(m.min), rd(m.max), rd(m.sd), Core.specText(x[2])]).eachCell(function (c) { c.border = box; });
      });
    st.addRow([]);
    st.addRow(['판정', ev.qaResult, '불합격 ' + ev.failCount + '건 / 전체 ' + ev.drums + '건']);
    st.addRow(['원본 데이터', state.fileName]);

    wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveBlob(blob, Core.docNumber($('date').value, val('orderNo')) + '.xlsx');
      setStatus('엑셀 파일을 만들었습니다. 다운로드 폴더를 확인하세요.', 'ok');
    });
  }

  // 결과 표를 탭 구분 텍스트로 복사 → 기존 엑셀 양식 첫 데이터 칸에 붙여넣기
  function copyResult() {
    var ev = state.ev; if (!ev) return;
    var keys = ['orderNo', 'batchNo', 'otdr', 'fiberNo', 'odMax', 'odMin', 'tightBuffer', 'att1310', 'att1550', 'marking'];
    var dig = { odMax: 2, odMin: 2, att1310: 3, att1550: 3 };
    var text = state.rows.map(function (r) {
      return keys.map(function (k) { return r[k] == null ? '' : (dig[k] ? Number(r[k]).toFixed(dig[k]) : r[k]); }).join('\t');
    }).join('\r\n');
    var done = function () { setStatus(state.rows.length + '행을 복사했습니다. 엑셀 양식의 첫 데이터 칸을 선택하고 Ctrl+V 하세요. 판정 결과: ' + ev.qaResult, 'ok'); };
    var ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    if (ok) return done();
    if (navigator.clipboard) navigator.clipboard.writeText(text).then(done, function () { setStatus('복사가 차단되었습니다. 엑셀로 저장 또는 PDF로 인쇄를 사용하세요.', 'err'); });
  }

  function downloadTemplate() {
    if (noExcel()) return;
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet('Data');
    var hdr = ['Order No.', 'Batch No.', 'OTDR (m)', 'Fiber No.', 'Cable OD Max', 'Cable OD Min', 'Tight Buffer', 'Att1310', 'Att1550', 'Marking'];
    ws.addRow(hdr).eachCell(function (c) { c.font = { bold: true }; c.fill = YELLOW; c.border = box; c.alignment = center; });
    ws.columns = hdr.map(function () { return { width: 14 }; });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    wb.xlsx.writeBuffer().then(function (buf) {
      saveBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'input_template.xlsx');
    });
  }

  function saveBlob(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // 수량 칸을 직접 고치면 자동값으로 덮어쓰지 않음
  document.addEventListener('DOMContentLoaded', function () {
    $('quantity').addEventListener('input', function () { this.dataset.touched = this.value ? '1' : ''; });
    init();
  });
})();
