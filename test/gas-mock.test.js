/* Code.js（GAS層）の構文チェックと、グリッド描画・上書き保存の簡易検証（SpreadsheetAppのモック） */
var fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
var src = path.join(__dirname, '..', 'src');

function makeSheet(name) {
  var cells = {}; // 'r,c' -> value
  var sh = {
    name: name, maxRows: 1000, maxCols: 30, charts: [],
    getName: function () { return name; },
    getMaxRows: function () { return this.maxRows; }, getMaxColumns: function () { return this.maxCols; },
    insertRowsAfter: function (a, n) { this.maxRows += n; }, insertColumnsAfter: function (a, n) { this.maxCols += n; },
    getLastRow: function () { var m = 0; Object.keys(cells).forEach(function (k) { if (cells[k] !== '' && cells[k] !== undefined) m = Math.max(m, +k.split(',')[0]); }); return m; },
    getLastColumn: function () { var m = 0; Object.keys(cells).forEach(function (k) { if (cells[k] !== '') m = Math.max(m, +k.split(',')[1]); }); return m; },
    clear: function () { cells = {}; }, clearContents: function () { cells = {}; },
    deleteRow: function (r) { this.deleteRows(r, 1); },
    deleteRows: function (r, n) {
      var next = {};
      Object.keys(cells).forEach(function (k) { var p = k.split(','), rr = +p[0]; if (rr >= r && rr < r + n) return; next[(rr >= r + n ? rr - n : rr) + ',' + p[1]] = cells[k]; });
      cells = next;
      this.maxRows -= n;
      this.groups = []; this.checkboxes = []; this.merges = [];
    },
    groups: [], checkboxes: [], merges: [],
    borders: [],
    getRangeList: function (list) {
      var self = this;
      return {
        insertCheckboxes: function () { self.checkboxes = self.checkboxes.concat(list); },
        setBorder: function (t, l, b, r, v, h, color, style) { self.borders.push({ list: list, top: t, left: l, bottom: b, color: color, style: style }); }
      };
    },
    setRowGroupControlPosition: function () {},
    getRange: function (r, c, nr, nc) {
      if (typeof r === 'string') { r = 2; c = 1; nr = 999; nc = 1; }
      nr = nr || 1; nc = nc || 1;
      var rg = {
        getValues: function () { var out = []; for (var i = 0; i < nr; i++) { var row = []; for (var j = 0; j < nc; j++) { var v = cells[(r + i) + ',' + (c + j)]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
        getDisplayValues: function () { return rg.getValues().map(function (row) { return row.map(String); }); },
        getDisplayValue: function () { return String(rg.getValues()[0][0]); },
        getValue: function () { return rg.getValues()[0][0]; },
        setValues: function (v) { v.forEach(function (row, i) { row.forEach(function (x, j) { cells[(r + i) + ',' + (c + j)] = x; }); }); return rg; },
        setValue: function (v) { cells[r + ',' + c] = v; return rg; },
        getRow: function () { return r; }, getColumn: function () { return c; },
        getNumRows: function () { return nr; }, getNumColumns: function () { return nc; },
        getSheet: function () { return sh; }
      };
      ['setNumberFormat', 'setNumberFormats', 'setBackgrounds', 'setBackground', 'setFontWeights', 'setFontWeight', 'setFontColor', 'setFontColors',
        'setHorizontalAlignment', 'setVerticalAlignment', 'setBorder', 'insertCheckboxes', 'breakApart', 'setVerticalText', 'setFontSize', 'setDataValidation', 'clearDataValidations'].forEach(function (m) { rg[m] = function () { return rg; }; });
      rg.merge = function () { sh.merges.push([r, c, nr, nc]); return rg; };
      rg.getMergedRanges = function () {
        return sh.merges.filter(function (m) { return m[1] === c && m[0] >= r && m[0] < r + nr; })
          .map(function (m) { return { getRow: function () { return m[0]; }, getLastRow: function () { return m[0] + m[2] - 1; } }; });
      };
      rg.setBackgrounds = function (m) { if (r === 1 && c === 1) sh.backgrounds = m; return rg; };
      rg.setFontColors = function (m) { if (r === 1 && c === 1) sh.fontColors = m; return rg; };
      rg.setNumberFormats = function (m) { if (r === 1 && c === 1) sh.numberFormats = m; return rg; };
      rg.shiftRowGroupDepth = function (d) { sh.groups.push([r, nr, d]); return rg; };
      return rg;
    },
    getDataRange: function () { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); },
    setRowHeight: function (r, h) { this.rowHeights = this.rowHeights || {}; this.rowHeights[r] = h; },
    images: [],
    getImages: function () { return this.images.slice(); },
    insertImage: function (blob, c, r) { var self = this, im = { blob: blob, col: c, row: r, alt: '', script: '' };
      im.setAltTextTitle = function (t) { im.alt = t; return im; }; im.getAltTextTitle = function () { return im.alt; };
      im.setAltTextDescription = function () { return im; }; im.setWidth = function (w) { im.w = w; return im; }; im.setHeight = function (h) { im.h = h; return im; };
      im.assignScript = function (f) { im.script = f; return im; };
      im.getAnchorCell = function () { return { getRow: function () { return im.row; }, getColumn: function () { return im.col; } }; };
      im.remove = function () { self.images = self.images.filter(function (x) { return x !== im; }); };
      self.images.push(im); return im; },
    setFrozenRows: function () {}, setFrozenColumns: function () {}, setColumnWidth: function () {}, setColumnWidths: function () {}, hideColumns: function () {}, showColumns: function () {},
    getCharts: function () { return this.charts.slice(); }, removeChart: function (c) { this.charts = this.charts.filter(function (x) { return x !== c; }); }, insertChart: function (c) { this.charts.push(c); },
    newChart: function () { var spec = { options: {} }, b = {}; ['setChartType', 'addRange', 'setNumHeaders', 'setPosition', 'setTransposeRowsAndColumns'].forEach(function (m) { b[m] = function (x) { spec[m] = Array.prototype.slice.call(arguments); return b; }; }); b.setOption = function (k, v) { spec.options[k] = v; return b; }; b.build = function () { return spec; }; return b; },
    hideSheet: function () {},
    getProtections: function () { return []; }, protect: function () { var p = { setDescription: function () { return p; }, setWarningOnly: function () { return p; } }; return p; },
    _cells: function () { return cells; }
  };
  return sh;
}
var sheets = {};
var ss = {
  getSheetByName: function (n) { return sheets[n] || null; },
  insertSheet: function (n) { sheets[n] = makeSheet(n); return sheets[n]; },
  getSpreadsheetTimeZone: function () { return 'Asia/Tokyo'; },
  toast: function (msg) { toasts.push(msg); }, getSheets: function () { return Object.keys(sheets).map(function (k) { return sheets[k]; }); },
  deleteSheet: function (sh) { delete sheets[sh.getName()]; }, setActiveSheet: function () {}
};
var toasts = [], props = {}, triggers = [];
var ctx = {
  console: console,
  SpreadsheetApp: {
    getActive: function () { return ss; }, flush: function () {},
    getUi: function () { throw new Error('no ui'); },
    newDataValidation: function () { var b = {}; ['requireValueInList', 'requireValueInRange', 'setAllowInvalid', 'requireCheckbox'].forEach(function (m) { b[m] = function () { return b; }; }); b.build = function () { return {}; }; return b; },
    ProtectionType: { SHEET: 'SHEET' },
    BorderStyle: { SOLID: 'SOLID', SOLID_THICK: 'SOLID_THICK' },
    GroupControlTogglePosition: { BEFORE: 'BEFORE' }
  },
  Utilities: {
    newBlob: function (bytes, type, name) { return { bytes: bytes, type: type, name: name }; },
    base64Decode: function (s) { return Buffer.from(s, 'base64'); },
    formatDate: function (d, tz, f) { var y = d.getFullYear(), m = ('0' + (d.getMonth() + 1)).slice(-2); return f === 'yyyy-MM' ? y + '-' + m : y + '' + m + '01000000'; },
    getUuid: function () { return Math.random().toString(16).slice(2) + 'abcd'; }
  },
  LockService: { getDocumentLock: function () { return { tryLock: function () { return true; }, releaseLock: function () {} }; } },
  PropertiesService: { getDocumentProperties: function () { return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; }, setProperties: function (o) { Object.keys(o).forEach(function (k) { props[k] = o[k]; }); }, deleteProperty: function (k) { delete props[k]; } }; } },
  ScriptApp: { getProjectTriggers: function () { return []; }, newTrigger: function (h) { var t = { forSpreadsheet: function () { return t; }, onEdit: function () { return t; }, timeBased: function () { return t; }, everyMinutes: function () { return t; }, create: function () { triggers.push(h); } }; return t; } },
  Session: { getActiveUser: function () { return { getEmail: function () { return 'tester@example.com'; } }; } },
  Charts: { ChartType: { LINE: 'LINE', COLUMN: 'COLUMN', COMBO: 'COMBO' } },
  HtmlService: { createHtmlOutput: function (h) { var o = { html: h }; o.setWidth = function () { return o; }; o.setHeight = function () { return o; }; return o; } }
};
vm.createContext(ctx);
['Engine.js', 'Schema.js', 'Sample.js', 'Logo.js', 'Code.js'].forEach(function (f) {
  vm.runInContext(fs.readFileSync(path.join(src, f), 'utf8'), ctx, { filename: f });
});
console.log('  ok  構文チェック（4ファイル）');

// 初期設定 → 再計算
ctx.setup();
function cell(r, c) { var v = sheets['シミュレーター']._cells()[r + ',' + c]; return v === undefined ? '' : v; }
var KEY = 12, CK = 11, MC = 13; // キー列・計上列・最初の月の列（M）
function findKey(key) { for (var r = 1; r < 400; r++) if (cell(r, KEY) === key) return r; return -1; }
function compute() { return vm.runInContext('CFEngine.compute(loadData_(SpreadsheetApp.getActive()))', ctx); }
var gs = sheets['シミュレーター'];

var R = compute();
assert.strictEqual(R.warnings.length, 0, R.warnings.join('\n'));
console.log('  ok  初期設定後、シートから読み込んだデータで警告なし');

// 見出し：2行目は年月、3行目は相対ラベル、決算月の印
assert.strictEqual(cell(2, MC), ymLabel(R.months[0]));
assert.strictEqual(cell(3, MC), '先月（n-1）');
assert.strictEqual(cell(3, MC + 1), '今月');
assert.strictEqual(cell(3, MC + 2), 'n+1月');
var fiscalCols = []; for (var j = 0; j < 19; j++) if (cell(1, MC + j) === '▼決算月▼') fiscalCols.push(R.months[j].slice(5));
assert.ok(fiscalCols.length >= 1 && fiscalCols.every(function (m) { return m === '03'; }));
var fiscalCol = MC + R.months.map(function (m) { return m.slice(5); }).indexOf('03');
assert.strictEqual(sheets['シミュレーター'].backgrounds[0][fiscalCol - 1], null);   // 塗りつぶしなし
assert.strictEqual(sheets['シミュレーター'].fontColors[0][fiscalCol - 1], '#ff0000'); // 赤の文字
function ymLabel(ym) { return ym.slice(0, 4) + '年' + Number(ym.slice(5, 7)) + '月'; }
console.log('  ok  見出し：年月・相対ラベル（先月/今月/n+k月）・決算月の印');

// 集計ブロックの並び（模範例と同じ順）
var order = ['#fc', '#fcEq', '#fcEqDebt', '#out', '#burn', '#runway', '#in', '#inEq', '#inEqDebt'].map(findKey);
assert.deepStrictEqual(order, [4, 5, 6, 7, 8, 9, 10, 11, 12]);
console.log('  ok  集計ブロック：正味FC→出資除く→出資・融資除く→出金合計→burn rate→ランウェイ→入金合計…');

// 入金合計の数式を「大分類→中分類→明細（計上✔のみ）」とたどって合計すると、エンジンの値と一致する
function refs(formula, col) { var re = new RegExp(col + '(\\d+)', 'g'), out = [], m; while ((m = re.exec(formula))) out.push(+m[1]); return out; }
function evalHierarchy(totalKey, j) {
  var colL = colLetter(MC + j), total = 0;
  refs(cell(findKey(totalKey), MC + j), colL).forEach(function (mr) {
    assert.ok(/^major:/.test(cell(mr, KEY)), '大分類行でない: ' + mr);
    refs(cell(mr, MC + j), colL).forEach(function (nr) {
      assert.ok(/^cat:/.test(cell(nr, KEY)), '中分類行でない: ' + nr);
      var f = String(cell(nr, MC + j)), m = /SUMIF\(\$K\$(\d+):\$K\$(\d+),TRUE/.exec(f);
      if (!m) return;
      for (var r = +m[1]; r <= +m[2]; r++) if (cell(r, CK) === true) total += cell(r, MC + j);
    });
  });
  return total;
}
function colLetter(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
[1, 5, 8].forEach(function (j) {
  assert.strictEqual(Math.round(evalHierarchy('#in', j) * 1000), R.totals.inflow[j], '入金 j=' + j);
  assert.strictEqual(Math.round(evalHierarchy('#out', j) * 1000), R.totals.outflow[j], '出金 j=' + j);
});
console.log('  ok  入金合計・出金合計：大分類→中分類（SUMIFで計上✔のみ）→明細の数式構造がエンジンの合計と一致（千円表示）');

var fcF = cell(findKey('#fc'), MC + 2);
assert.ok(/SUM\(\$N\d+:O\d+\)/.test(fcF), '正味FCが今月からの累計式でない: ' + fcF);
assert.ok(!/SUM/.test(cell(findKey('#fc'), MC)), '先月の正味FCは月初残高＋現金同等物のみ: ' + cell(findKey('#fc'), MC));
assert.ok(/^=\(\$M\$\d+(\+\$M\$\d+)*\)/.test(fcF), '正味FCが月初残高を参照していない: ' + fcF);
assert.strictEqual(cell(findKey('#burn'), MC), '');
assert.strictEqual(cell(findKey('#burn'), MC + 3), '=(N7*0.8+O7*0.9+P7*1)/2.7');
assert.ok(/^=O4-\(SUM\(\$N\d+:O\d+\)(\+SUM\(\$N\d+:O\d+\))*\)$/.test(cell(findKey('#fcEq'), MC + 2)), cell(findKey('#fcEq'), MC + 2));
console.log('  ok  正味FC（月初残高＋入金累計−出金累計＋現金同等物）・burn rate（加重平均）・出資除く（累計控除）の数式');

assert.ok(gs.checkboxes.length > 20 && gs.groups.length > 10 && gs.merges.length === 5);
var groupsBefore = gs.groups.length;
console.log('  ok  計上チェックボックス ' + gs.checkboxes.length + ' 個・行グループ ' + gs.groups.length + ' 個・セクション帯 5 個');

// ---- 見た目：フォント色・数値書式・罫線
var fcM = gs.fontColors, nfM = gs.numberFormats;
function fontAt(key, col) { return fcM[findKey(key) - 1][col - 1]; }
assert.strictEqual(fontAt('item:I040', 5), '#0000ff');            // 入力行（明細項目）の項目名＝青
assert.strictEqual(fontAt('item:I040', MC + 2), '#0000ff');       // 入力行の数値＝青
assert.strictEqual(fontAt('item:I013', MC + 2), '#0000ff');       // 計上✗でも入力行は青
assert.strictEqual(fontAt('sys:SYS:社会保険料', 5), '#000000');   // 自動計算行＝黒
assert.strictEqual(fontAt('gen:L:R01:I022', MC + 6), '#000000');  // 連動行＝黒
assert.strictEqual(fontAt('#fcEq', 4), '#000000');                // 集計の小見出し＝黒
assert.strictEqual(fontAt('cat:A01', MC + 1), '#000000');         // 小計行＝黒
assert.strictEqual(nfM[findKey('#fc') - 1][MC], '#,##0;[Red]-#,##0;0');   // マイナスは「-」付き赤字
console.log('  ok  フォント色：入力行は青・自動計算行は黒、マイナスは「-」付きの赤字');

assert.ok(gs.borders.every(function (b) { return b.style === 'SOLID'; }), '太線が残っている');
var thinBlack = gs.borders.filter(function (b) { return b.color === '#000000' && b.left === null; })[0];
var bandBox = gs.borders.filter(function (b) { return b.color === '#000000' && b.left === true; })[0];
var thinGray = gs.borders.filter(function (b) { return b.style === 'SOLID' && b.color === '#b7b7b7'; })[0];
var lastColL = colLetter(MC + 18);
assert.ok(thinBlack.list.indexOf('C' + findKey('major:売上') + ':' + lastColL + findKey('major:売上')) >= 0);
assert.ok(thinBlack.list.indexOf('C' + findKey('#fc') + ':' + lastColL + findKey('#fc')) >= 0);
assert.ok(thinGray.list.indexOf('C' + findKey('cat:A01') + ':' + lastColL + findKey('cat:A01')) >= 0);
assert.ok(thinGray.list.indexOf('C' + findKey('item:I001') + ':' + lastColL + findKey('item:I001')) >= 0);
assert.ok(thinGray.list.indexOf('C' + findKey('#fcEq') + ':' + lastColL + findKey('#fcEq')) >= 0);
assert.ok(thinBlack.list.every(function (a) { return thinGray.list.indexOf(a) < 0; }));
assert.strictEqual(bandBox.list.length, 5);
assert.ok(bandBox.list.indexOf('B4:' + lastColL + '12') >= 0);      // 集計帯：B4〜最終列12行目のブロック
assert.ok(bandBox.top === true && bandBox.bottom === true && bandBox.left === true);  // 上辺・下辺・左辺
assert.ok(gs.borders.indexOf(thinGray) < gs.borders.indexOf(thinBlack) && gs.borders.indexOf(thinBlack) < gs.borders.indexOf(bandBox));
var bgM = gs.backgrounds;
function bgAt(key, col) { return bgM[findKey(key) - 1][col - 1]; }
assert.strictEqual(bgAt('item:I040', MC), '#d9d9d9');              // 先月：色の無いセルはグレー
assert.strictEqual(bgAt('#in', MC), '#d9d9d9');
assert.strictEqual(bgAt('#burn', MC), '#ff9900');                  // 既に色があるセルはそのまま
assert.strictEqual(bgAt('major:売上', MC), '#e3ebfa');
assert.strictEqual(bgM[1][MC - 1] === '#d9d9d9', false);           // 見出し行（B列が結合されていない）は対象外
assert.strictEqual(bgAt('item:I040', MC + 5) === '#d9d9d9', false); // 今月以降は対象外
console.log('  ok  M列（先月）：セクション帯の範囲で色の無いセルはグレー');
console.log('  ok  罫線：すべて細線。セクション帯はブロックの上辺・下辺（B〜最終列）と左辺を黒、C列の行は黒、D/E列の行は灰');

// ---- 挙動：編集時は記録だけ（再計算しない）→ メニュー「再計算」または無操作時に反映
assert.ok(triggers.indexOf('onEditInstalled') >= 0 && triggers.indexOf('autoRecalc') >= 0);
var rentRow = findKey('item:I040');
var fcBefore = cell(findKey('#fc'), MC + 3);
gs.getRange(rentRow, MC + 3).setValue(999);
ctx.onEditInstalled({ range: gs.getRange(rentRow, MC + 3), source: ss, oldValue: 450, user: null });
assert.strictEqual(props.dirty, '1');
assert.ok(/未反映/.test(cell(2, 5)));
assert.strictEqual(cell(rentRow, MC + 3), 999);                     // 入力した値はそのまま残る
assert.strictEqual(compute().rows.filter(function (x) { return x.key === 'I040'; })[0].values[3], 450000); // まだDBには書かない
// 素早い連続編集で編集トリガーが発火しなかったケース（onEditInstalled を呼ばない）
var saasRow = findKey('item:I050');
gs.getRange(saasRow, MC + 4).setValue(200);
gs.getRange(saasRow, MC + 5).setValue(210);
console.log('  ok  編集トリガーは「未反映」の印を付けるだけ（DB書き込み・再計算をしないので連続入力を妨げない）');

ctx.autoRecalc();                                                    // 最後の編集から20秒たっていないので何もしない
assert.strictEqual(props.dirty, '1');
props.lastEdit = String(Date.now() - 30000);
ctx.autoRecalc();                                                    // 無操作20秒以上 → 自動再計算
assert.strictEqual(props.dirty, undefined);
assert.ok(/最終計算/.test(cell(2, 5)));
var R3 = compute();
assert.strictEqual(R3.rows.filter(function (x) { return x.key === 'I040'; })[0].values[3], 999000);
assert.strictEqual(R3.rows.filter(function (x) { return x.key === 'I050'; })[0].values[4], 200000);
assert.strictEqual(R3.rows.filter(function (x) { return x.key === 'I050'; })[0].values[5], 210000);
assert.strictEqual(cell(findKey('item:I050'), MC + 5), 210);
assert.strictEqual(gs.groups.length, groupsBefore, '値だけの変更で行グループが作り直された');
console.log('  ok  自動再計算：無操作20秒以上で1回だけ再計算。トリガーが発火しなかった連続編集も差分検出ですべて取り込む');

ctx.menuRecalc();                                                    // 何も変えずに再計算しても上書きは増えない
assert.strictEqual(sheets['明細'].getDataRange().getValues().filter(function (r) { return r[2] === '上書き'; }).length, 3);
console.log('  ok  変更のない再計算では明細に何も追加しない');

// ---- ダッシュボード（1画面に1つ：B1で切り替え、A列のチェックでグラフの系列を選択）
var ds = sheets['ダッシュボード'], hp = sheets['_グラフ用'];
assert.ok(ds && hp && !sheets['グラフ'], 'ダッシュボード／_グラフ用 シートが無い、または旧グラフシートが残っている');
function dcell(r, c) { var v = ds._cells()[r + ',' + c]; return v === undefined ? '' : v; }
function hcell(r, c) { var v = hp._cells()[r + ',' + c]; return v === undefined ? '' : v; }
function dfind(label) { for (var r = 1; r < 100; r++) if (String(dcell(r, 2)).replace(/^　/, '') === label) return r; return -1; }
function dashEdit(r, c, v) { ds.getRange(r, c).setValue(v); ctx.onEditInstalled({ range: ds.getRange(r, c), source: ss, user: null }); }
function series() { var s = ds.getCharts()[0].options.series; return Object.keys(s).map(function (k) { return s[k]; }); }
assert.strictEqual(dcell(1, 2), '全体');
assert.ok(/^全体Dashboard/.test(dcell(1, 3)));
assert.strictEqual(dcell(2, 3), R.months[0].slice(0, 4) + '/' + Number(R.months[0].slice(5, 7)));
assert.strictEqual(dcell(2, 3 + 19), 'K JPY');
assert.strictEqual(dcell(3, 2), '正味フリーキャッシュ');
assert.strictEqual(dcell(3, 3), "='シミュレーター'!M" + findKey('#fc'));
assert.strictEqual(dcell(dfind('burn rate'), 5), "='シミュレーター'!O" + findKey('#burn'));
assert.ok([2, 3, 4, 5, 6].every(function (r) { return dcell(r, 1) === true; }));       // 初期状態は全系列を表示
assert.strictEqual(ds.getCharts().length, 1);
var c0 = ds.getCharts()[0];
assert.ok(c0.setChartType[0] === 'COMBO' && c0.setTransposeRowsAndColumns[0] === true);
var s0 = series();
assert.ok(s0.length === 4 && s0[0].type === 'bars' && s0[0].targetAxisIndex === 1 && s0[1].type === 'line');
assert.strictEqual(c0.setPosition[0], 3 + 4 + 1);                             // 表のすぐ下にグラフ
assert.strictEqual(hcell(2, 1), '正味フリーキャッシュ');
console.log('  ok  ダッシュボード：初期表示は全体（表＋表の直下に複合グラフ、全系列を表示）');

dashEdit(1, 2, '出金');
assert.strictEqual(dcell(3, 2), '固定費用');
assert.strictEqual(dcell(3, 3), "='シミュレーター'!M" + findKey('major:固定費用'));
assert.strictEqual(dcell(dfind('人件費（雇用保険・社会保険対象）'), 3), "='シミュレーター'!M" + findKey('cat:B01'));
var nOut = 0; for (var r1 = 3; r1 < 60; r1++) if (dcell(r1, 2) !== '') nOut++;
assert.strictEqual(series().length, nOut);
assert.strictEqual(series().filter(function (x) { return x.type === 'bars'; }).length, 2);
assert.strictEqual(ds.getCharts().length, 1);
console.log('  ok  B1のプルダウンで出金に切り替え（大分類＝棒、中分類＝折れ線）');

dashEdit(3, 1, false);                                                       // 固定費用をグラフから外す
assert.strictEqual(series().length, nOut - 1);
assert.strictEqual(hcell(2, 1), '人件費（雇用保険・社会保険対象）');
assert.strictEqual(dcell(2, 1), false);                                      // 全選択のチェックも外れる
dashEdit(2, 1, false);                                                       // 全解除 → グラフなし
assert.strictEqual(ds.getCharts().length, 0);
dashEdit(2, 1, true);                                                        // 全選択
assert.strictEqual(series().length, nOut);
dashEdit(dfind('地代家賃'), 1, false);
console.log('  ok  A列のチェックでグラフの系列を選択（全選択／全解除も可）');

dashEdit(1, 2, '入金');
assert.strictEqual(dcell(3, 2), '売上');
assert.strictEqual(dcell(dfind('受託・コンサル'), 3), "='シミュレーター'!M" + findKey('cat:A01'));
dashEdit(1, 2, '出金');
assert.strictEqual(dcell(dfind('地代家賃'), 1), false);                       // 切り替えても選択は保持
assert.strictEqual(series().length, nOut - 1);
var chartObj = ds.getCharts()[0];
ctx.menuRecalc();
assert.strictEqual(ds.getCharts()[0], chartObj, '行構成が変わらないのにグラフが作り直された');
assert.strictEqual(dcell(1, 2), '出金');
console.log('  ok  ダッシュボードごとに選択を保持。行構成が変わらない再計算では作り直さない');

// 手作業で行が挿入され、結合（セクション帯）がずれた場合 → 次の再計算でレイアウトを作り直す
gs.merges = gs.merges.map(function (m) { return m[0] > 20 ? [m[0] + 1, m[1], m[2], m[3]] : m; });
ctx.menuRecalc();
assert.strictEqual(gs.merges.length, 5);
assert.ok(gs.groups.length === groupsBefore);
var outBand = gs.merges.filter(function (m) { return m[0] === findKey('major:固定費用'); })[0];
assert.ok(outBand, '出金の帯が出金の先頭行から結合されていない');
console.log('  ok  結合がずれていたら（手作業での行挿入など）レイアウトを自動で作り直す');

gs.getRange(findKey('item:I040'), MC + 3).setValue('');
ctx.menuRecalc();
assert.strictEqual(cell(findKey('item:I040'), MC + 3), 450);
console.log('  ok  セルを空にしてメニュー「再計算」→ ルール値（450千円）に戻る');

gs.getRange(findKey('item:I040'), MC).setValue(1);
ctx.onEditInstalled({ range: gs.getRange(findKey('item:I040'), MC), source: ss, oldValue: '450', user: null });
assert.strictEqual(cell(findKey('item:I040'), MC), 450);             // 再計算を待たずにその場で元に戻す
assert.ok(toasts.some(function (t) { return /先月/.test(t); }));
gs.getRange(findKey('item:I040'), MC).setValue(5);                  // トリガーが発火しなかった場合も、再計算で元に戻る
ctx.menuRecalc();
assert.strictEqual(cell(findKey('item:I040'), MC), 450);
assert.ok(!toasts.some(function (t) { return /混雑/.test(t); }));
console.log('  ok  先月（実績月）への直接入力は元に戻る（「混雑中」の警告は出ない）');

// 計上チェック：見込み案件（✗）を✔にすると合計に入る
var dRow = findKey('item:I013');
assert.strictEqual(cell(dRow, CK), false);
var inBefore = compute().totals.inflow[5];
gs.getRange(dRow, CK).setValue(true);
ctx.menuRecalc();
assert.strictEqual(cell(findKey('item:I013'), CK), true);
assert.strictEqual(compute().totals.inflow[5] - inBefore, 440000);
console.log('  ok  K列の計上チェックを切り替えると項目シートの「有効」が更新され、合計に反映');

// 月初残高：先月の列に入力すると口座項目の金額が更新される
var bankRow = findKey('open:K01');
gs.getRange(bankRow, MC).setValue(59000);
ctx.menuRecalc();
assert.strictEqual(compute().totals.opening[0], 59500000);
assert.strictEqual(cell(findKey('open:K01'), MC), 59000);
console.log('  ok  月初残高の口座行に入力すると項目の金額（円）が更新される');
var log = sheets['変更ログ'].getLastRow();
assert.ok(log >= 3);
console.log('  ok  変更ログに記録');

// ---- ロゴ（C1、クリックで再計算）・アプリ名
var logos = gs.getImages().filter(function (im) { return im.alt === 'cashboard_logo'; });
assert.strictEqual(logos.length, 1);
assert.ok(logos[0].row === 1 && logos[0].col === 3 && logos[0].script === 'menuRecalc');
assert.ok(logos[0].blob.type === 'image/jpeg' && logos[0].blob.bytes.length > 5000);
assert.strictEqual(cell(1, 3), '');                                   // C1 に文字は入れない
ctx.menuRebuild();                                                    // レイアウトを作り直しても1つだけ
assert.strictEqual(gs.getImages().filter(function (im) { return im.alt === 'cashboard_logo'; }).length, 1);
assert.strictEqual(vm.runInContext('APP_NAME', ctx), 'cashboard');
console.log('  ok  C1 にロゴ画像（クリックで再計算）。作り直しても重複しない。メニュー名は cashboard');

// ---- 基準月：空欄なら最初の再計算で今日の月に固定される
function paramValue(key) { var v = sheets['パラメータ'].getDataRange().getValues(); for (var i = 1; i < v.length; i++) if (v[i][0] === key) return v[i][1]; return null; }
var base0 = compute().base;
assert.strictEqual(paramValue('基準月'), base0);
console.log('  ok  基準月が空欄なら今日の月で固定（カレンダーが変わっても勝手に進まない）');

// ---- 月初処理
var info = vm.runInContext('monthStartInfo_()', ctx);
var next = vm.runInContext('CFEngine.addMonths', ctx)(base0, 1);
assert.strictEqual(info.from, base0);
assert.strictEqual(info.to, next);
assert.ok(info.openings.length === 2 && info.openings[0].key === 'K01');
assert.ok(info.last.some(function (x) { return x.key === 'I040' && x.editable; }));
assert.ok(info.last.some(function (x) { return x.kind === 'sys' && !x.editable; }));
var before = compute();
var nowVals = {};
before.rows.forEach(function (r) { nowVals[r.key] = r.values[1]; });
var rentLast = info.last.filter(function (x) { return x.key === 'I040'; })[0];
var msg = ctx.applyMonthStart({ from: base0, openings: [{ key: 'K01', amount: 61234567 }, { key: 'K02', amount: 400000 }],
  last: [{ key: 'I040', value: rentLast.value + 50 }] });
assert.ok(/完了/.test(msg), msg);
var after = compute();
assert.strictEqual(after.base, next);
assert.strictEqual(paramValue('基準月'), next);
assert.strictEqual(after.months[0], before.months[1]);                  // 表が1か月左にずれる
assert.strictEqual(cell(2, MC), ymLabel(after.months[0]));             // M2〜の年月が＋1か月
assert.strictEqual(cell(2, MC + 18), ymLabel(after.months[18]));
['I040', 'I050', 'I020', 'L:R02:I020'].forEach(function (k) {         // 新しい先月（M列）＝処理前の今月（N列）
  var r = after.rows.filter(function (x) { return x.key === k; })[0];
  assert.strictEqual(r.values[0], nowVals[k], k);
  assert.strictEqual(r.sources[0], '実績', k);
});
assert.strictEqual(cell(findKey('item:I040'), MC), nowVals['I040'] / 1000);
assert.strictEqual(after.totals.opening[0], 61634567);                 // 入力した月初残高
assert.strictEqual(cell(findKey('open:K01'), MC), 61234.567);
var rentAfter = after.rows.filter(function (x) { return x.key === 'I040'; })[0];
assert.strictEqual(rentAfter.values[18], (rentLast.value + 50) * 1000); // 最終月を直した値＝上書き
assert.strictEqual(rentAfter.sources[18], '上書き');
var saas = after.rows.filter(function (x) { return x.key === 'I050'; })[0];
assert.ok(saas.values[18] > 0 && saas.sources[18] !== '上書き');       // 直していない行はルールで推計
assert.throws(function () { ctx.applyMonthStart({ from: base0, openings: [], last: [] }); }, /すでに月初処理/);
console.log('  ok  月初処理：表が1か月進み、旧今月の値が先月（実績）に、月初残高は入力値、最終月はルール推計＋直した行は上書き');
var compute2 = compute();
ctx.menuRecalc();
assert.strictEqual(compute().totals.fc[18], compute2.totals.fc[18]);
console.log('  ok  月初処理のあとの再計算で値が変わらない（差分取り込みが誤検出しない）');

// ---- メニュー「月初処理」：確認・入力はスクリプト側の画面（ui.alert / ui.prompt）。HTMLダイアログ（google.script.run）は使わない
function fakeUi(script) {
  var log = [];
  var ui = {
    ButtonSet: { YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL', OK: 'OK' }, Button: { YES: 'YES', NO: 'NO', OK: 'OK', CANCEL: 'CANCEL' },
    alert: function (a, b) { var m = b || a; log.push(m); return /続けますか|実行しますか/.test(m) ? script.shift() : 'OK'; },
    prompt: function (t, m) { log.push(m); var x = script.shift(); return { getSelectedButton: function () { return x === null ? 'CANCEL' : 'OK'; }, getResponseText: function () { return x || ''; } }; }
  };
  return { ui: ui, log: log };
}
var realGetUi = ctx.SpreadsheetApp.getUi;
var baseB = compute().base;
var f1 = fakeUi(['NO']);
ctx.SpreadsheetApp.getUi = function () { return f1.ui; };
ctx.menuMonthStart();
assert.strictEqual(compute().base, baseB);                                  // 「いいえ」なら何も変えない
var f2 = fakeUi(['YES', 'abc', '１２３，０００円', '', 'YES']);
ctx.SpreadsheetApp.getUi = function () { return f2.ui; };
ctx.menuMonthStart();
ctx.SpreadsheetApp.getUi = realGetUi;
var R5 = compute();
assert.strictEqual(R5.base, vm.runInContext('CFEngine.addMonths', ctx)(baseB, 1));
assert.ok(f2.log.some(function (m) { return /数値として読めませんでした/.test(m); }));   // 数値以外は聞き直す
assert.strictEqual(R5.rows.filter(function (r) { return r.key === 'K01'; })[0].values[0], 123000);
assert.strictEqual(R5.rows.filter(function (r) { return r.key === 'K02'; })[0].values[0], 400000); // 空欄＝処理前の値のまま
assert.ok(/月初処理が完了しました/.test(f2.log[f2.log.length - 1]));
assert.ok(!/google\.script\.run\(/.test(require('fs').readFileSync(require('path').join(src, 'Code.js'), 'utf8')));
console.log('  ok  メニュー「月初処理」：確認→口座ごとの残高入力（数値以外は聞き直し、空欄は処理前の値）→最終確認→実行。いいえ／キャンセルなら何も変えない');
console.log('\nall passed');
