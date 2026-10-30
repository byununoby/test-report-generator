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
    $('date').addEventListener('change', suggestSeq);
    suggestSeq();
    // Customer: 한 번 입력하면 이 PC 브라우저에 고정 (고객사명이 GitHub에 올라가지 않도록 코드에는 넣지 않음)
    $('customer').value = store('customer') || C.defaultCustomer || '';
    $('customer').addEventListener('change', function () { store('customer', this.value.trim()); });
    loadSpecOverrides();
    sel.addEventListener('change', function () { $('color').value = colorName(typeOf().color); renderSpecEditor(); refresh(); });
    $('color').value = colorName(typeOf().color);

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
    $('mapReset').addEventListener('click', function () {
      if (!state.an) return;
      try { localStorage.removeItem(state.sig); } catch (e) {}
      state.map = Core.guessMapping(state.an); state.mapSaved = false; applyMapping();
    });

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
      state = { rows: [], ev: null, fileName: '', an: null };
      ['orderNo', 'poNo', 'quantity', 'batchNo', 'productName'].forEach(function (id) { $(id).value = ''; });
      $('pasteBox').value = '';
      setStatus('', ''); renderMapping(); refresh();
    });
    renderSpecEditor();
    refresh();
  }

  // 드럼 1개당 기준 길이(km). 표지 Quantity의 (km)와 파일명에 사용: 320 D/M → 640km
  var DRUM_KM = 2;

  function colorName(code) { return (C.colorNames && C.colorNames[code]) || code; }

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

  // 엑셀에서 복사한 텍스트 처리 (탭 구분, 줄바꿈 들어간 제목 칸 포함)
  var pasteTimer;
  function usePasted(text) {
    clearTimeout(pasteTimer);
    pasteTimer = setTimeout(function () {
      var table = text.indexOf('\t') !== -1 ? Core.parseTSV(text) : Core.parseCSV(text);
      useTable(table, '붙여넣은 데이터');
    }, 150);
  }

  // ---------- 저장 (이 PC 브라우저에만 보관) ----------
  function store(key, val) {
    try {
      if (val === undefined) { var v = localStorage.getItem(key); return v ? JSON.parse(v) : null; }
      localStorage.setItem(key, JSON.stringify(val));
    } catch (e) { return null; }
  }

  function useTable(table, label) {
    try {
      var an = Core.analyze(table);
      state.an = an; state.fileName = label;
      state.sig = 'colmap2:' + an.labels.join('|');
      var saved = an.hasHeader ? store(state.sig) : null;
      state.map = saved || Core.guessMapping(an);
      state.mapSaved = !!saved;
      ['orderNo', 'poNo', 'quantity', 'batchNo'].forEach(function (id) { $(id).value = ''; });
      applyMapping();
    } catch (e) { setStatus(e.message, 'err'); }
  }

  function applyMapping() {
    var rows = Core.buildRows(state.an, state.map);
    state.rows = rows;
    var missing = Core.FIELDS.filter(function (f) { return !f.optional && !(state.map[f.key] >= 0); }).map(function (f) { return f.label; });
    if (!rows.length) setStatus('연결된 열에서 측정값을 찾지 못했습니다. 아래 열 연결을 확인하세요.', 'err');
    else {
      var msg = label() + ' ' + rows.length + '행을 읽었습니다.' + (state.mapSaved ? ' 저장된 열 연결을 적용했습니다.' : '');
      if (missing.length) msg += ' 연결 안 된 항목: ' + missing.join(', ');
      setStatus(msg, missing.length ? 'warn' : 'ok');
    }
    renderMapping();
    refresh();
  }
  function label() { return state.fileName + '에서'; }

  // 열 연결 화면: 자동 인식 결과를 보여주고 틀리면 고칠 수 있게 함
  function renderMapping() {
    var an = state.an, box = $('mapBox');
    if (!an) { box.hidden = true; return; }
    box.hidden = false;
    var opts = '<option value="-1">(연결 안 함)</option>' + an.labels.map(function (l, i) {
      var name = Core.colName(i) + '열' + (l ? ': ' + l : '');
      if (name.length > 38) name = name.slice(0, 37) + '…';
      return '<option value="' + i + '">' + esc(name) + (an.samples[i] ? '  [' + esc(an.samples[i]) + ']' : '') + '</option>';
    }).join('');
    $('mapGrid').innerHTML = Core.FIELDS.map(function (f) {
      return '<label><span>' + esc(f.label) + '</span><select data-key="' + f.key + '">' + opts + '</select></label>';
    }).join('');
    Array.prototype.forEach.call($('mapGrid').querySelectorAll('select'), function (sel) {
      sel.value = String(state.map[sel.dataset.key]);
      sel.addEventListener('change', function () {
        state.map[this.dataset.key] = parseInt(this.value, 10);
        if (state.an.hasHeader) { store(state.sig, state.map); state.mapSaved = true; }
        applyMapping();
      });
    });
  }

  // ---------- 규격 수정 (이 PC에만 저장, GitHub에 안 올라감) ----------
  var SPEC_FIELDS = [
    ['cableOD', 'min', 'Cable O.D. 최소 (mm)'], ['cableOD', 'max', 'Cable O.D. 최대 (mm)'],
    ['tightBuffer', 'min', 'Tight Buffer 최소 (um)'], ['tightBuffer', 'max', 'Tight Buffer 최대 (um)'],
    ['att1310', 'max', '1310 최대 (dB/km)'], ['att1550', 'max', '1550 최대 (dB/km)'],
    ['marking', 'min', '인쇄간격 최소 (mm)'], ['marking', 'max', '인쇄간격 최대 (mm)']
  ];
  function loadSpecOverrides() {
    var o = store('specOverrides') || {};
    Object.keys(o).forEach(function (g) {
      if (!C.specs[g]) return;
      Object.keys(o[g]).forEach(function (item) {
        C.specs[g][item] = C.specs[g][item] || {};
        Object.keys(o[g][item]).forEach(function (k) { C.specs[g][item][k] = o[g][item][k]; });
      });
    });
    var lbl = store('orderNoLabel'); if (lbl) C.orderNoLabel = lbl;
  }
  function renderSpecEditor() {
    var g = typeOf().spec, sp = C.specs[g];
    $('specGroup').textContent = g + ' 제품 규격';
    $('specGrid').innerHTML = SPEC_FIELDS.map(function (f) {
      var v = sp[f[0]] && sp[f[0]][f[1]] != null ? sp[f[0]][f[1]] : '';
      return '<label>' + f[2] + '<input type="text" inputmode="decimal" data-item="' + f[0] + '" data-k="' + f[1] + '" value="' + v + '"></label>';
    }).join('') + '<label class="wide">보고서 주문번호 열 제목<input type="text" id="orderNoLabelInput" value="' + esc(C.orderNoLabel) + '"></label>';
    Array.prototype.forEach.call($('specGrid').querySelectorAll('input[data-item]'), function (inp) {
      inp.addEventListener('change', function () {
        var n = parseFloat(this.value), item = this.dataset.item, k = this.dataset.k;
        var o = store('specOverrides') || {};
        o[g] = o[g] || {}; o[g][item] = o[g][item] || {};
        sp[item] = sp[item] || {};
        if (isNaN(n)) { delete sp[item][k]; o[g][item][k] = null; } else { sp[item][k] = n; o[g][item][k] = n; }
        store('specOverrides', o);
        refresh();
      });
    });
    $('orderNoLabelInput').addEventListener('change', function () {
      C.orderNoLabel = this.value.trim() || 'Customer Order No.';
      store('orderNoLabel', C.orderNoLabel); renderPreview();
    });
  }

  // ---------- 판정 및 화면 갱신 ----------
  function refresh() {
    state.ev = state.rows.length ? Core.evaluate(state.rows, $('cableType').value, C) : null;
    var ev = state.ev;
    if (ev) {
      var first = state.rows[0] || {};
      $('orderNo').placeholder = Core.salesPrefix(first.salesNo) || '';
      $('poNo').placeholder = first.orderNo || '';
      $('quantity').placeholder = ev.drums + ' D/M (' + ev.drums * DRUM_KM + 'km)';
      var b = ev.batches;
      $('batchNo').placeholder = !b.length ? '' : b.length === 1 ? '#' + b[0] : '#' + b[0] + ' ~ #' + b[b.length - 1];
      $('productName').placeholder = ev.type.name;
    } else {
      ['orderNo', 'poNo', 'quantity', 'batchNo'].forEach(function (id) { $(id).placeholder = '데이터를 불러오면 자동 입력'; });
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
      '<span>1550 ' + Core.specText(s.att1550) + '</span>' +
      (s.marking ? '<span>인쇄간격 ' + Core.specText(s.marking) + ' mm</span>' : '');
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
        fails.push('<li>' + esc(state.rows[i].batchNo || ('행 ' + (i + 1))) + ': ' + items.join(', ') + '</li>');
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

  // ---------- 출력값 (미리보기, 엑셀 공통) ----------
  function out() {
    var t = typeOf(), ev = state.ev;
    var qty = val('quantity');
    return {
      type: t,
      date: $('date').value,
      usDate: Core.usDate($('date').value),
      orderNo: val('orderNo'),
      poNo: val('poNo'),
      customer: val('customer'),
      product: val('productName') || t.name,
      color: val('color'),
      quantity: qty,
      quantityShort: qty.replace(/\s*\(.*\)\s*$/, ''),
      batch: val('batchNo'),
      qa: ev ? ev.qaResult : '',
      signers: [$('writer').value, $('reviewer').value, $('approver').value]
    };
  }

  // ---------- 미리보기 (엑셀 양식과 같은 배치) ----------
  function renderPreview() {
    var ev = state.ev, o = out();
    var refRows = C.cableTypes.map(function (x) {
      return '<tr><td>' + x.id + '</td><td class="l">' + esc(x.name) + '</td><td>' + esc(x.fiber) + '</td><td>' + esc(x.color) + '</td></tr>';
    }).join('');
    var cover =
      '<section class="paper cover">' +
      '<div class="cv-frame">' +
      '<h2 class="cv-title">TEST REPORT</h2>' +
      '<table class="cov">' +
      row2('Date', o.date) + row2('Order No.', o.orderNo) + row2('PO No.', o.poNo) + row2('Customer', o.customer) +
      row2('Cable Type', o.product) + row2('COLOR', o.color) + row2('Quantity', o.quantity) +
      '<tr class="tall"><th>Batch No.</th><td>' + esc(o.batch) + '</td></tr></table>' +
      '<table class="sign"><tr><th>작&nbsp;&nbsp;&nbsp;성</th><th>검&nbsp;&nbsp;&nbsp;토</th><th>승&nbsp;&nbsp;&nbsp;인</th></tr><tr><td>' +
      esc(o.signers[0]) + '</td><td>' + esc(o.signers[1]) + '</td><td>' + esc(o.signers[2]) + '</td></tr></table>' +
      '</div>' +
      '<div class="ref"><p class="ref-title">※ 제품명 참고</p><table class="reftab"><tr><th></th><th>Cable Type</th><th>FIBER</th><th>COLOR</th></tr>' + refRows + '</table></div>' +
      '</section>';

    var body = '';
    if (ev) {
      body = state.rows.map(function (r, i) {
        var c = ev.results[i].checks;
        return '<tr>' + td(r.orderNo, false, 'big') + td(r.batchNo) + td(r.otdr == null ? '' : fmtInt(r.otdr), !c.otdr) + td(r.fiberNo) +
          td(fmt(r.odMax, 2), !c.odMax) + td(fmt(r.odMin, 2), !c.odMin) + td(r.tightBuffer, !c.tightBuffer) +
          td(fmt(r.att1310, 3), !c.att1310) + td(fmt(r.att1550, 3), !c.att1550) + td(r.marking, !c.marking) + '</tr>';
      }).join('');
    } else {
      for (var i = 0; i < 8; i++) body += '<tr>' + new Array(11).join('<td>&nbsp;</td>') + '</tr>';
    }
    var report =
      '<section class="paper report">' +
      '<div class="rule"></div>' +
      '<h2 class="rep-title">Test&nbsp;&nbsp;&nbsp;&nbsp;Report</h2>' +
      '<dl class="rep-head">' +
      dl('Product Name', o.product) + dl('Color', o.color) + dl('Inspection Date', o.usDate) +
      dl('Quantity(D/M)', o.quantityShort) + dl('QA RESULTS', o.qa) + '</dl>' +
      '<div class="tablewrap"><table class="rep">' +
      '<thead><tr><th rowspan="2">' + esc(C.orderNoLabel).replace(' Order', '<br>Order') + '</th><th rowspan="2">Batch No.</th><th rowspan="2">OTDR (m)</th><th rowspan="2">FIBER<br>NO.</th>' +
      '<th colspan="2">Cable O.D. (mm)</th><th rowspan="2">Tight Buffer O.D.<br>(um)</th><th colspan="2">Attenuation(dB/Km)</th><th rowspan="2">marking</th></tr>' +
      '<tr><th>Max.</th><th>Min.</th><th>ose1310</th><th>ose1550</th></tr></thead>' +
      '<tbody>' + body + '</tbody></table></div></section>';

    $('preview').innerHTML = cover + report;
  }

  function row2(k, v) { return '<tr><th>' + k + '</th><td>' + esc(v) + '</td></tr>'; }
  function dl(k, v) { return '<dt>' + k + '</dt><dd>' + esc(v) + '</dd>'; }
  function td(v, bad, cls) {
    var c = (bad ? 'bad ' : '') + (cls || '');
    return '<td' + (c.trim() ? ' class="' + c.trim() + '"' : '') + '>' + esc(v == null ? '' : v) + '</td>';
  }

  // ---------- 엑셀 출력 (사내 양식 배치 그대로) ----------
  var thin = { style: 'thin', color: { argb: 'FF000000' } };
  var medium = { style: 'medium', color: { argb: 'FF000000' } };
  var dotted = { style: 'dotted', color: { argb: 'FF000000' } };
  var box = { top: thin, left: thin, bottom: thin, right: thin };
  var dotBox = { top: dotted, left: dotted, bottom: dotted, right: dotted };
  var YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  var GREY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  var RED = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
  var center = { horizontal: 'center', vertical: 'middle', wrapText: true };
  var KFONT = '맑은 고딕';

  function noExcel() {
    if (typeof ExcelJS !== 'undefined') return false;
    setStatus('엑셀 라이브러리가 없어 이 기능을 쓸 수 없습니다. 오프라인 패키지(test-report-offline)를 받아서 사용하세요.', 'err');
    return true;
  }

  // 범위 전체에 테두리 (병합 셀 포함)
  function boxRange(ws, r1, c1, r2, c2, side) {
    for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) {
      var cell = ws.getRow(r).getCell(c), b = Object.assign({}, cell.border || {});
      if (r === r1) b.top = side; if (r === r2) b.bottom = side;
      if (c === c1) b.left = side; if (c === c2) b.right = side;
      cell.border = b;
    }
  }
  function mergeSet(ws, range, value, style) {
    ws.mergeCells(range);
    var cell = ws.getCell(range.split(':')[0]);
    cell.value = value;
    Object.keys(style || {}).forEach(function (k) { cell[k] = style[k]; });
    return cell;
  }

  function buildCover(wb, o) {
    var ws = wb.addWorksheet('Cover', {
      views: [{ showGridLines: false }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
        horizontalCentered: true, verticalCentered: true,
        margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 } }
    });
    ws.pageSetup.printArea = 'A$1:H$32'; // ExcelJS가 $A$1:$H$32로 저장
    // A4 세로 한 장(A1:H32)을 꽉 채우는 크기
    [7, 7, 11, 11, 11, 13.5, 13.5, 13.5, 6.7, 10, 25, 13, 9].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    for (var r = 1; r <= 32; r++) ws.getRow(r).height = 21;
    ws.getRow(1).height = 10; ws.getRow(32).height = 10;

    mergeSet(ws, 'A5:H8', 'TEST REPORT', { font: { name: 'Times New Roman', size: 36, bold: true }, alignment: center });

    var items = [['Date', o.date], ['Order No.', o.orderNo], ['PO No.', o.poNo], ['Customer', o.customer],
      ['Cable Type', o.product], ['COLOR', o.color], ['Quantity', o.quantity]];
    var bold = { name: KFONT, size: 13, bold: true };
    items.forEach(function (it, i) {
      var r = 14 + i;
      mergeSet(ws, 'C' + r + ':D' + r, it[0], { font: bold, alignment: center, fill: GREY });
      mergeSet(ws, 'E' + r + ':G' + r, it[1], { font: bold, alignment: center });
      ws.getRow(r).height = 30;
      boxRange(ws, r, 3, r, 4, thin); boxRange(ws, r, 5, r, 7, thin);
    });
    mergeSet(ws, 'C21:D22', 'Batch No.', { font: bold, alignment: center, fill: GREY });
    mergeSet(ws, 'E21:G22', o.batch, { font: bold, alignment: center });
    ws.getRow(21).height = 30; ws.getRow(22).height = 30;
    boxRange(ws, 21, 3, 22, 4, thin); boxRange(ws, 21, 5, 22, 7, thin);

    ['작    성', '검    토', '승    인'].forEach(function (h, i) {
      var col = 6 + i, L = String.fromCharCode(64 + col);
      var c = ws.getCell(L + '28'); c.value = h; c.font = { name: KFONT, size: 12, bold: true }; c.alignment = center;
      mergeSet(ws, L + '29:' + L + '31', o.signers[i] || '', { font: { name: KFONT, size: 12 }, alignment: center });
      boxRange(ws, 28, col, 28, col, thin); boxRange(ws, 29, col, 31, col, thin);
    });
    ws.getRow(28).height = 24;
    for (var sr = 29; sr <= 31; sr++) ws.getRow(sr).height = 32;
    boxRange(ws, 2, 1, 31, 8, medium);   // 표지 바깥 테두리

    var small = { name: KFONT, size: 11, bold: true };
    var t = ws.getCell('J13'); t.value = '※ 제품명 참고'; t.font = { name: KFONT, size: 11, bold: true };
    ['', 'Cable Type', 'FIBER', 'COLOR'].forEach(function (h, i) {
      var c = ws.getRow(14).getCell(10 + i); c.value = h; c.font = small; c.alignment = center; c.border = dotBox;
    });
    C.cableTypes.forEach(function (x, i) {
      var row = ws.getRow(15 + i);
      [x.id, x.name, x.fiber, x.color].forEach(function (v, j) {
        var c = row.getCell(10 + j); c.value = v; c.font = small; c.border = dotBox;
        c.alignment = j === 1 ? { horizontal: 'left', vertical: 'middle' } : center;
      });
    });
    return ws;
  }

  function buildType(wb, o) {
    var ev = state.ev;
    var ws = wb.addWorksheet('TYPE', {
      views: [{ showGridLines: false }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } }
    });
    ws.pageSetup.printTitlesRow = '12:13';
    // 데이터 글자(맑은 고딕 11pt)에 맞춰 전체 비율 확대
    [13, 11, 11, 8, 10, 10, 18, 11, 11, 11].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.getRow(1).height = 10; ws.getRow(2).height = 10; ws.getRow(3).height = 6;
    for (var c = 1; c <= 10; c++) ws.getRow(3).getCell(c).border = { bottom: medium };

    ws.getRow(4).height = 38;
    mergeSet(ws, 'A4:J4', 'Test    Report', { font: { name: 'Arial', size: 22, bold: true, underline: true }, alignment: { horizontal: 'center', vertical: 'middle' } });
    ws.getRow(5).height = 6;

    var head = [['Product Name', o.product], ['Color', o.color], ['Inspection Date', o.usDate], ['Quantity(D/M)', o.quantityShort], ['QA RESULTS', o.qa]];
    head.forEach(function (h, i) {
      var r = 6 + i;
      ws.getRow(r).height = 20;
      var l = ws.getCell('D' + r); l.value = h[0]; l.font = { name: 'Arial', size: 12, bold: true, italic: true };
      ws.mergeCells('F' + r + ':G' + r);
      var v = ws.getCell('F' + r); v.value = h[1]; v.font = { name: 'Arial', size: 12 };
      v.alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getCell('F' + r).border = { bottom: thin }; ws.getCell('G' + r).border = { bottom: thin };
    });
    ws.getRow(11).height = 10;

    var hf = { name: KFONT, size: 11 };
    var label = C.orderNoLabel.replace(' Order', '\nOrder');
    [['A12:A13', label], ['B12:B13', 'Batch No.'], ['C12:C13', 'OTDR (m)'], ['D12:D13', 'FIBER\nNO.'],
      ['E12:F12', 'Cable O.D. (mm)'], ['G12:G13', 'Tight Buffer O.D.\n(um)'], ['H12:I12', 'Attenuation(dB/Km)'], ['J12:J13', 'marking']]
      .forEach(function (x) { mergeSet(ws, x[0], x[1], { font: hf, alignment: center }); });
    [['E13', 'Max.'], ['F13', 'Min.'], ['H13', 'ose1310'], ['I13', 'ose1550']].forEach(function (x) {
      var c = ws.getCell(x[0]); c.value = x[1]; c.font = hf; c.alignment = center;
    });
    ws.getRow(12).height = 24; ws.getRow(13).height = 24;
    for (var r0 = 12; r0 <= 13; r0++) for (var c0 = 1; c0 <= 10; c0++) {
      var hc = ws.getRow(r0).getCell(c0); hc.fill = YELLOW; hc.border = box;
    }

    var keys = ['orderNo', 'batchNo', 'otdr', 'fiberNo', 'odMax', 'odMin', 'tightBuffer', 'att1310', 'att1550', 'marking'];
    var numFmt = { otdr: '#,##0', odMax: '0.00', odMin: '0.00', tightBuffer: '0', att1310: '0.000', att1550: '0.000' };
    state.rows.forEach(function (row, i) {
      var r = ws.getRow(14 + i), checks = ev.results[i].checks;
      r.height = 22.2;   // 37픽셀(화면 배율 125% 기준)
      keys.forEach(function (k, j) {
        var c = r.getCell(j + 1), v = row[k] == null ? '' : row[k];
        if ((k === 'marking' || k === 'orderNo' || k === 'fiberNo') && v !== '' && /^\d+(\.\d+)?$/.test(String(v))) v = Number(v);
        c.value = v;
        if (numFmt[k]) c.numFmt = numFmt[k];
        c.font = { name: KFONT, size: 11 };
        c.alignment = center; c.border = box;
        if (checks[k] === false) { c.fill = RED; c.font = { name: KFONT, size: 11, bold: true, color: { argb: 'FFC00000' } }; }
      });
    });
    return ws;
  }

  // 파일명: 261029-1 MDU Drop Cable Test Report(2.9Riser)(TYPE-1_WH_120km).xlsx
  function fileName(o) {
    var d = String(o.date || '').replace(/-/g, '').slice(2) || '000000';
    var seq = parseInt($('fileSeq').value, 10) || 1;
    // 2.9㎜ Riser Cable → 2.9Riser, 2.9㎜ Riser I/O Cable → 2.9IO, 2.9㎜ Plenum Cable → 2.9Plenum
    var size = (/^[\d.]+/.exec(o.type.name) || [''])[0];
    var short = /I\/O/i.test(o.type.name) ? size + 'IO'
      : o.type.name.replace('㎜', '').replace(/mm/i, '').replace(/\s*Cable\s*/i, '').replace(/\s+/g, '');
    var code = o.type.color;
    Object.keys(C.colorNames || {}).forEach(function (k) { if (C.colorNames[k].toLowerCase() === String(o.color).toLowerCase()) code = k; });
    var m = /\(\s*([\d.,]+)\s*km\s*\)/i.exec(o.quantity);
    var km = m ? m[1] : String(state.ev.drums * DRUM_KM);
    var name = d + '-' + seq + ' MDU Drop Cable Test Report(' + short + ')(' +
      o.type.id.toUpperCase().replace(' ', '-') + '_' + code + '_' + km + 'km).xlsx';
    return name.replace(/[\\/:*?"<>|]/g, '');
  }

  // 같은 날짜에 저장할 때마다 순번 자동 증가 (이 PC에 기억, 직접 수정 가능)
  function seqKey() { return 'fileSeq:' + $('date').value; }
  function suggestSeq() { $('fileSeq').value = (store(seqKey()) || 0) + 1; }

  function exportExcel() {
    if (!state.ev || noExcel()) return;
    var o = out(), wb = new ExcelJS.Workbook();
    wb.creator = 'Test Report Generator';
    buildCover(wb, o);
    buildType(wb, o);
    wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveBlob(blob, fileName(o));
      store(seqKey(), parseInt($('fileSeq').value, 10) || 1);
      suggestSeq();
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
