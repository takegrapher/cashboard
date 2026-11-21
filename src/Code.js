/**
 * Code.js — スプレッドシートとのやり取り（読み込み・描画・編集の取り込み・メニュー）
 * 計算そのものは Engine.js（CFEngine.compute）が行う。
 *
 * 「シミュレーター」シートの列構成（模範例「キャッシュフロー」シートに準拠）
 *   A: 余白  B: セクション帯  C: 大分類  D: 中分類  E: 明細項目
 *   F: 説明  G: 費目対象先  H: タグ（F〜Hは非表示。必要なら再表示）
 *   I: 構成比  J: 合計（今月〜）  K: 計上（✔/✗）  L: キー（非表示）  M〜: 月（先月・今月・n+1〜n+17月）
 */

var APP_NAME = 'cashboard';
var SHEET = {
  GRID: 'シミュレーター',
  DASH: 'ダッシュボード',
  CALC: 'CALC_展開結果',
  LOG: '変更ログ',
  SNAP: 'スナップショット',
  DETAIL: '明細',
  ITEM: '項目'
};
// 編集されたら自動で再計算する入力シート（DB）
var INPUT_SHEETS = ['科目', '取引先', '項目', '人件費', '連動ルール', '明細', 'パラメータ'];

var GRID = { CHECK_COL: 11, KEY_COL: 12, FIRST_MONTH_COL: 13, HEADER_ROWS: 3 };
var BAND_COLOR = { '集計': '#ff9900', '月初残高': '#999999', '現金同等物': '#b6d7a8', '入金': '#a4c2f4', '出金': '#ea9999' };
var MAJOR_COLOR = { '月初残高': '#efefef', '現金同等物': '#e6f2e0', '入金': '#e3ebfa', '出金': '#f9e4e4' };
var COLOR = {
  header: '#1f3864', headerFont: '#ffffff', past: '#f3f3f3', nowHead: '#a4c2f4', now: '#eaf1fb', burn: '#ff9900',
  actual: '#d9d9d9', actualPlan: '#ededed', override: '#fff2cc', fiscal: '#ff0000', excluded: '#b7b7b7',
  input: '#0000ff', gen: '#3c78d8', sys: '#38761d', label: '#efefef', pastCol: '#d9d9d9'
};
var NUM_FMT = '#,##0;[Red]-#,##0;0';          // マイナスは「-」付きの赤字
var BORDER = { black: '#000000', gray: '#b7b7b7' };
var FONT = { auto: '#000000', input: '#0000ff' }; // 自動計算の行＝黒、ユーザーが入力する行＝青

// ================================================================ メニュー・トリガー
function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_NAME)
    .addItem('再計算', 'menuRecalc')
    .addItem('月初処理（1か月進める）…', 'menuMonthStart')
    .addItem('選択セルの内訳を表示', 'showBreakdown')
    .addItem('スナップショットを保存', 'saveSnapshot')
    .addItem('警告を表示', 'showWarnings')
    .addSeparator()
    .addItem('レイアウトを作り直す', 'menuRebuild')
    .addItem('初期設定（シート作成・トリガー登録）', 'setup')
    .addToUi();
}

/*
 * 再計算のタイミング
 *   ・セルを編集したとき：「未反映あり」の印を付けるだけ（すぐ終わるので連続入力を妨げない）。
 *     編集内容は、再計算の直前に「前回表示した値」との差分から取り込んでDB（明細・項目シート）に記録する
 *   ・再計算するのは次のどちらか
 *       1) メニュー「再計算」を実行したとき
 *       2) 自動再計算：1分ごとの時間トリガーが、未反映の変更があり、かつ最後の編集から AUTO_IDLE_SEC 秒以上
 *          操作がないときだけ再計算する（パラメータ「自動再計算」を「しない」にすると止まる）
 */
var AUTO_IDLE_SEC = 20;

/** インストール型の編集トリガー（setup で登録。オーナー権限で動くので保護シートにも書ける） */
function onEditInstalled(e) {
  if (!e || !e.range) return;
  var name = e.range.getSheet().getName();
  try {
    if (name === SHEET.GRID) { handleGridEdit_(e); return; }
    if (name === SHEET.DASH) { handleDashEdit_(e); return; }
    if (INPUT_SHEETS.indexOf(name) < 0) return;
    if (name === SHEET.DETAIL) stampDetailRows_(e);
    markDirty_(e.source);
  } catch (err) {
    console.error(err && err.stack || err);
    e.source.toast(String(err && err.message || err), 'エラー', 10);
  }
}

/** 未反映の変更があることを記録し、シミュレーターの2行目に表示する */
function markDirty_(ss) {
  var p = PropertiesService.getDocumentProperties();
  p.setProperties({ dirty: '1', lastEdit: String(Date.now()), editSeq: String(Number(p.getProperty('editSeq') || 0) + 1) });
  var sh = ss.getSheetByName(SHEET.GRID);
  if (sh) sh.getRange(2, 5).setValue('⏳ 未反映の変更があります（メニュー「再計算」、または操作をやめて約1分後に自動反映）').setFontColor('#cc0000');
}

function menuRecalc() {
  var ss = SpreadsheetApp.getActive();
  var ok = runRecalc_(60000);
  ss.toast(ok ? '再計算しました' : '自動再計算が実行中です。終わり次第反映されます', APP_NAME, 4);
}

/** 時間トリガー（1分ごと）から呼ばれる自動再計算 */
function autoRecalc() {
  var p = PropertiesService.getDocumentProperties();
  if (p.getProperty('dirty') !== '1') return;
  if (Date.now() - Number(p.getProperty('lastEdit') || 0) < AUTO_IDLE_SEC * 1000) return;
  var ss = SpreadsheetApp.getActive();
  var param = ss.getSheetByName('パラメータ');
  if (param) {
    var rows = param.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === '自動再計算' && /しない|オフ|OFF|FALSE|無効/i.test(String(rows[i][1]))) return;
    }
  }
  runRecalc_(5000);
}

/** ロックを取って再計算する。再計算中に新しい編集が入った場合は「未反映」のまま残し、次回に回す */
function runRecalc_(waitMs) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(waitMs)) return false;
  try {
    var p = PropertiesService.getDocumentProperties();
    var seq = p.getProperty('editSeq');
    recalc_();
    if (p.getProperty('editSeq') === seq) p.deleteProperty('dirty');
    else markDirty_(SpreadsheetApp.getActive());
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** 行のグループ化や結合が崩れたときに、シミュレーターシートを作り直す */
function menuRebuild() {
  PropertiesService.getDocumentProperties().deleteProperty('gridSig');
  menuRecalc();
}

/*
 * シミュレーター上の編集の取り込み（差分検出方式）
 *   素早く続けて編集すると、Googleスプレッドシートは2件目以降の編集トリガーを発火しないことがある。
 *   そのため編集トリガーには頼らず、再計算の直前に「前回表示した値（_表示状態シート）」と「今のセルの値」を
 *   比べて、変わったセルをまとめてDBへ記録する。
 */
var STATE_SHEET = '_表示状態';

function captureGridEdits_(ss) {
  var grid = ss.getSheetByName(SHEET.GRID), state = ss.getSheetByName(STATE_SHEET);
  if (!grid || !state || state.getLastRow() < 2) return { overrides: 0, updates: 0, past: 0 };
  var sv = state.getDataRange().getValues();
  var stMonths = sv[0].slice(2).map(function (x) { return CFEngine.normalizeYm(x); }); // 日付に自動変換されていても年月に戻す
  var old = {};
  for (var i = 1; i < sv.length; i++) old[String(sv[i][0])] = sv[i];
  var F = GRID.FIRST_MONTH_COL, M = stMonths.length;
  var n = grid.getLastRow();
  if (n <= GRID.HEADER_ROWS) return { overrides: 0, updates: 0, past: 0 };
  var curMonths = grid.getRange(2, F, 1, M).getDisplayValues()[0].map(function (x) { return CFEngine.normalizeYm(x); });
  if (curMonths.join() !== stMonths.join()) return { overrides: 0, updates: 0, past: 0 }; // 月が切り替わった直後は比較しない
  var base = curMonths[1];
  var S = Number(PropertiesService.getDocumentProperties().getProperty('displayUnit')) || 1;
  var keys = grid.getRange(1, GRID.KEY_COL, n, 1).getDisplayValues();
  var checks = grid.getRange(1, GRID.CHECK_COL, n, 1).getValues();
  var vals = grid.getRange(1, F, n, M).getValues();
  var overrides = [], updates = [], past = 0;
  function same(a, b) {
    if ((a === '' || a === null) && (b === '' || b === null)) return true;
    if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-6;
    return String(a) === String(b);
  }
  for (var r = GRID.HEADER_ROWS; r < n; r++) {
    var key = keys[r][0], o = old[key];
    if (!o) continue;
    var k = parseKey_(key);
    if (k.type !== 'item' && k.type !== 'open' && k.type !== 'gen' && k.type !== 'sys') continue;
    if ((k.type === 'item' || k.type === 'open') && checks[r][0] !== o[1] && typeof checks[r][0] === 'boolean') {
      updates.push({ id: k.id, header: '有効', value: checks[r][0] });
    }
    for (var j = 0; j < M; j++) {
      var cur = vals[r][j], prev = o[2 + j];
      if (same(cur, prev)) continue;
      if (k.type === 'open') {
        if (j !== 0) continue;
        var nv = cur === '' ? 0 : CFEngine.num(cur, null);
        if (nv !== null) updates.push({ id: k.id, header: '金額（税抜）', value: Math.round(nv * S) });
        continue;
      }
      if (curMonths[j] < base) { past++; continue; }
      if (cur === '' || cur === null) { overrides.push({ key: k.id, ym: curMonths[j], value: null }); continue; }
      var num = CFEngine.num(cur, null);
      if (num !== null) overrides.push({ key: k.id, ym: curMonths[j], value: Math.round(num * S) });
    }
  }
  var who = '（シミュレーター上の編集）';
  if (overrides.length) { applyOverrides_(ss, overrides, who); logChanges_(ss, overrides, who, ''); }
  if (updates.length) {
    updateItems_(ss, updates);
    logChanges_(ss, updates.map(function (u) { return { key: u.id, ym: u.header, value: u.value }; }), who, '');
  }
  if (past) ss.toast('先月（実績月）のセルは直接入力できないため、元に戻しました。実績は「明細」シートに登録してください', '入力できません', 8);
  return { overrides: overrides.length, updates: updates.length, past: past };
}

/** 今回表示した値を保存する（次回の差分検出の基準） */
function saveGridState_(ss, months, rows) {
  var sh = ss.getSheetByName(STATE_SHEET);
  if (!sh) { sh = ss.insertSheet(STATE_SHEET); sh.hideSheet(); }
  var data = [['キー', '計上'].concat(months)].concat(rows);
  sh.clearContents();
  ensureSize_(sh, data.length + 1, data[0].length);
  sh.getRange(1, 1, 1, data[0].length).setNumberFormat('@'); // 年月の見出しが日付に変換されないように
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function userEmail_(e) {
  try { if (e && e.user && e.user.getEmail && e.user.getEmail()) return e.user.getEmail(); } catch (x) { /* ignore */ }
  try { return Session.getActiveUser().getEmail() || ''; } catch (x) { return ''; }
}

// ================================================================ 再計算
function todayYm_(ss) {
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM');
}

function loadData_(ss) {
  function t(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 1) return [];
    return cfRowsToObjects(name, sh.getDataRange().getValues());
  }
  var params = {};
  t('パラメータ').forEach(function (p) { var k = String(p.key || '').trim(); if (k) params[k] = p.value; });
  return {
    todayYm: todayYm_(ss), params: params,
    accounts: t('科目'), partners: t('取引先'), items: t('項目'), payroll: t('人件費'),
    links: t('連動ルール'), details: t('明細')
  };
}

function recalc_() {
  var ss = SpreadsheetApp.getActive();
  pinBaseMonth_(ss);
  captureGridEdits_(ss);
  var R = CFEngine.compute(loadData_(ss));
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('displayUnit', String(R.params.displayUnit));
  writeCalc_(ss, R);
  renderDashboard_(ss, R, renderGrid_(ss, R));
  props.setProperty('warnings', JSON.stringify(R.warnings.slice(0, 100)).slice(0, 8500));
  if (R.warnings.length) ss.toast('警告が ' + R.warnings.length + ' 件あります（メニュー → 警告を表示）', '再計算', 6);
  SpreadsheetApp.flush();
  return R;
}

function getOrCreate_(ss, name, index) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = index === undefined ? ss.insertSheet(name) : ss.insertSheet(name, index);
  return sh;
}
function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
function colLetter_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
function ymLabel_(ym) { return ym.slice(0, 4) + '年' + Number(ym.slice(5, 7)) + '月'; }

// ---------------------------------------------------------------- CALC_展開結果（計算の根拠を縦持ちで保存）
function writeCalc_(ss, R) {
  var sh = getOrCreate_(ss, SHEET.CALC);
  var header = ['項目キー', '科目', '項目名', '年月', '金額', 'ソース', '計算根拠', '発生月', '税区分', '明細ID', '摘要'];
  var data = [header].concat(R.long.map(function (x) {
    return [x.key, x.accountName, x.itemName, x.ym, x.amount, x.source, x.ref, x.accrual, x.taxCode, x.detailId, x.memo];
  }));
  sh.clearContents();
  ensureSize_(sh, data.length + 1, header.length);
  sh.getRange(1, 1, data.length, header.length).setNumberFormat('@');
  sh.getRange(2, 5, Math.max(1, data.length - 1), 1).setNumberFormat('#,##0');
  sh.getRange(1, 1, data.length, header.length).setValues(data);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------- シミュレーター（模範例レイアウト）
function rowKey_(r) {
  if (r.section === '月初残高') return 'open:' + r.key;
  return (r.kind === 'item' ? 'item:' : r.kind === 'gen' ? 'gen:' : 'sys:') + r.key;
}
function parseKey_(s) {
  s = String(s || '');
  var i = s.indexOf(':');
  if (i < 0) return { type: s.charAt(0) === '#' ? 'sum' : '', id: s };
  return { type: s.slice(0, i), id: s.slice(i + 1) };
}

function renderGrid_(ss, R) {
  var sh = getOrCreate_(ss, SHEET.GRID, 0);
  var tz = ss.getSpreadsheetTimeZone();
  var M = R.months.length, F = GRID.FIRST_MONTH_COL, W = F - 1 + M, S = R.params.displayUnit || 1;
  function L(j) { return colLetter_(F + j); }
  var firstL = L(0), nowL = L(1), lastL = L(M - 1);

  // ---- 1) 行の並びを決める
  var rows = [];
  function add(d) { d.r = rows.length + 1; rows.push(d); return d; }
  add({ t: 'h1' }); add({ t: 'h2' }); add({ t: 'h3' });
  var SUMS = [
    ['fc', 3, '正味フリーキャッシュ', '月末時点の現金・預金＋現金同等物（先月の列＝月初残高＝今月初の実際の残高）'],
    ['fcEq', 4, '出資除く正味フリーキャッシュ', '表示期間内のEquity調達額（累計）を除いた場合'],
    ['fcEqDebt', 4, '出資・融資除く正味フリーキャッシュ', '表示期間内のEquity・Debt調達額（累計）を除いた場合'],
    ['out', 3, '出金合計', '計上✔の出金の単月合計'],
    ['burn', 3, 'burn rate', '今月以降の出金の加重平均（当月1.0、1ヶ月遡るごとに0.1減、最大10ヶ月）'],
    ['runway', 3, 'ランウェイ（月）', '正味フリーキャッシュ ÷ burn rate'],
    ['in', 3, '入金合計', '計上✔の入金の単月合計'],
    ['inEq', 4, '出資除く入金合計', 'Equityによる調達を除いた入金合計'],
    ['inEqDebt', 4, '出資・融資除く入金合計', 'Equity・Debtによる調達を除いた入金合計']
  ];
  var sumRow = {}, bands = [], groups = [], secInfo = {};
  var first = rows.length + 1;
  SUMS.forEach(function (s) { sumRow[s[0]] = add({ t: 'sum', key: s[0], col: s[1], label: s[2], desc: s[3], sec: '集計' }); });
  bands.push({ name: '集計', sec: '集計', first: first, last: rows.length });
  add({ t: 'blank' });
  ['月初残高', '現金同等物', '入金', '出金'].forEach(function (sec) {
    var accs = R.accounts.filter(function (a) { return a.section === sec; });
    if (!accs.length) return;
    var info = secInfo[sec] = { majors: [], minors: [] };
    var start = rows.length + 1, majors = [];
    accs.forEach(function (a) { if (majors.indexOf(a.major) < 0) majors.push(a.major); });
    majors.forEach(function (mj) {
      var mRow = add({ t: 'major', sec: sec, label: mj, children: [] });
      info.majors.push(mRow);
      accs.filter(function (a) { return a.major === mj; }).forEach(function (a) {
        var nRow = add({ t: 'minor', sec: sec, acc: a, items: [] });
        mRow.children.push(nRow);
        info.minors.push(nRow);
        R.rows.filter(function (r) { return r.accountId === a.id; }).forEach(function (r) {
          nRow.items.push(add({ t: 'value', sec: sec, row: r }));
        });
        if (nRow.items.length) groups.push([nRow.r + 1, nRow.items.length]);
      });
      if (rows.length > mRow.r) groups.push([mRow.r + 1, rows.length - mRow.r]);
    });
    bands.push({ name: sec === '現金同等物' ? '現金・現金同等物' : sec, sec: sec, first: start, last: rows.length });
    add({ t: 'blank' });
  });

  // ---- 2) 数式の部品
  function plus(list, j) { return list.length ? list.map(function (d) { return L(j) + d.r; }).join('+') : '0'; }
  function majorsOf(sec) { return secInfo[sec] ? secInfo[sec].majors : []; }
  function inMinors(fin) { return secInfo['入金'] ? secInfo['入金'].minors.filter(function (n) { return n.acc.finance === fin; }) : []; }
  // 累計は今月から（先月の入出金は月初残高に反映済み）
  function cum(list, j) { return list.length && j > 0 ? list.map(function (d) { return 'SUM($' + nowL + d.r + ':' + L(j) + d.r + ')'; }).join('+') : '0'; }
  var openMajors = majorsOf('月初残高');
  var openRef = openMajors.length ? '(' + openMajors.map(function (d) { return '$' + firstL + '$' + d.r; }).join('+') + ')' : String(R.totals.opening[0] / S);
  var rIn = sumRow['in'].r, rOut = sumRow['out'].r, rFc = sumRow['fc'].r, rBurn = sumRow['burn'].r;
  function sumFormula(key, j) {
    var c = L(j);
    switch (key) {
      case 'in': return '=' + plus(majorsOf('入金'), j);
      case 'out': return '=' + plus(majorsOf('出金'), j);
      case 'inEq': return '=' + c + rIn + '-(' + plus(inMinors('出資'), j) + ')';
      case 'inEqDebt': return '=' + c + rIn + '-(' + plus(inMinors('出資'), j) + ')-(' + plus(inMinors('融資'), j) + ')';
      case 'fc': return '=' + openRef + (j > 0 ? '+SUM($' + nowL + rIn + ':' + c + rIn + ')-SUM($' + nowL + rOut + ':' + c + rOut + ')' : '') + '+(' + plus(majorsOf('現金同等物'), j) + ')';
      case 'fcEq': return '=' + c + rFc + '-(' + cum(inMinors('出資'), j) + ')';
      case 'fcEqDebt': return '=' + c + rFc + '-(' + cum(inMinors('出資'), j) + ')-(' + cum(inMinors('融資'), j) + ')';
      case 'burn':
        if (j === 0) return '';
        var terms = [], sw = 0;
        for (var q = Math.max(1, j - 9); q <= j; q++) { var w = Math.round((1 - 0.1 * (j - q)) * 10) / 10; sw += w; terms.push(L(q) + rOut + '*' + w); }
        return '=(' + terms.join('+') + ')/' + (Math.round(sw * 10) / 10);
      case 'runway':
        if (j === 0) return '';
        return '=IFERROR(IF(' + c + rBurn + '>0,' + c + rFc + '/' + c + rBurn + ',""),"")';
    }
    return '';
  }
  function totalRowOf(sec) { return sec === '入金' ? rIn : sec === '出金' ? rOut : 0; }

  // ---- 3) セルの値・書式を組み立てる
  var data = [], bg = [], weight = [], fcolor = [], fmt = [];
  function fill(v) { var a = []; for (var i = 0; i < W; i++) a.push(v); return a; }
  var nowTime = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm');
  rows.forEach(function (d) {
    var v = fill(''), b = fill(null), w = fill('normal'), fc = fill(null), f = fill(NUM_FMT);
    var j;
    v[GRID.KEY_COL - 1] = '';
    if (d.t === 'h1') {
      // C1 にはロゴ画像（クリックで再計算）を重ねる。文字は入れない
      for (j = 0; j < M; j++) {
        if (Number(R.months[j].slice(5, 7)) === R.params.fiscalEndMonth) {
          v[F - 1 + j] = '▼決算月▼'; fc[F - 1 + j] = COLOR.fiscal; w[F - 1 + j] = 'bold'; // 塗りつぶしなし・赤の太字
        }
      }
      f = fill('@');
    } else if (d.t === 'h2' || d.t === 'h3') {
      if (d.t === 'h2') {
        var cal = todayYm_(ss);
        v[4] = '最終計算 ' + nowTime + (R.warnings.length ? '　⚠ 警告 ' + R.warnings.length + ' 件' : '') +
          (R.base < cal ? '　⚠ 月初処理が未実施です（カレンダー上の今月：' + ymLabel_(cal) + '）' : '');
        fc[4] = R.base < cal ? '#cc0000' : '#666666';
      }
      if (d.t === 'h3') {
        v[4] = S === 1000 ? 'K JPY' : 'JPY';
        v[5] = '説明'; v[6] = '費目対象先'; v[7] = 'タグ'; v[8] = '構成比'; v[9] = '合計（今月〜）'; v[10] = '計上';
        for (j = 5; j <= 10; j++) { b[j] = COLOR.label; w[j] = 'bold'; }
      }
      for (j = 0; j < M; j++) {
        var cc = F - 1 + j;
        v[cc] = d.t === 'h2' ? ymLabel_(R.months[j]) : (j === 0 ? '先月（n-1）' : j === 1 ? '今月' : 'n+' + (j - 1) + '月');
        b[cc] = j === 0 ? COLOR.past : j === 1 ? COLOR.nowHead : null;
        w[cc] = d.t === 'h2' ? 'bold' : 'normal';
        if (d.t === 'h3') fc[cc] = '#666666';
      }
      f = fill('@');
    } else if (d.t === 'sum') {
      v[d.col - 1] = d.label; v[5] = d.desc; v[GRID.KEY_COL - 1] = '#' + d.key;
      w[d.col - 1] = d.col === 3 ? 'bold' : 'normal';
      for (j = 2; j < W; j++) fc[j] = FONT.auto;
      for (j = 0; j < M; j++) {
        v[F - 1 + j] = sumFormula(d.key, j);
        if (j === 1) b[F - 1 + j] = COLOR.now;
        if (d.col === 3) w[F - 1 + j] = 'bold';
        if (d.key === 'runway') f[F - 1 + j] = '0.0;[Red]-0.0;0.0';
      }
      if (d.key === 'in' || d.key === 'out') v[9] = '=SUM(' + nowL + d.r + ':' + lastL + d.r + ')';
      if (d.key === 'burn') for (j = 2; j < W; j++) b[j] = COLOR.burn;
    } else if (d.t === 'major' || d.t === 'minor' || d.t === 'value') {
      var sec = d.sec, isOpen = sec === '月初残高', tot = totalRowOf(sec), r = d.row;
      if (d.t === 'major') {
        v[2] = d.label; w[2] = 'bold';
        for (j = 2; j < W; j++) b[j] = MAJOR_COLOR[sec];
        v[GRID.KEY_COL - 1] = 'major:' + d.label;
      } else if (d.t === 'minor') {
        v[3] = d.acc.name; v[5] = d.acc.desc; v[GRID.KEY_COL - 1] = 'cat:' + d.acc.id;
      } else {
        v[4] = r.itemName + (r.kind === 'gen' ? '（連動）' : r.kind === 'sys' ? '（自動）' : '');
        v[5] = r.desc; v[6] = r.partnerName; v[7] = r.tags; v[10] = r.included;
        v[GRID.KEY_COL - 1] = rowKey_(r);
      }
      // 明細項目（ユーザーが入力する行）は青、連動・自動計算の行と小計行は黒
      var rowFont = d.t === 'value' && r.kind === 'item' ? FONT.input : FONT.auto;
      for (j = 2; j < W; j++) fc[j] = rowFont;
      for (j = 0; j < M; j++) {
        var c = F - 1 + j;
        if (isOpen && j > 0) continue;
        if (d.t === 'major') v[c] = '=' + plus(d.children, j);
        else if (d.t === 'minor') {
          v[c] = d.items.length
            ? '=SUMIF($K$' + d.items[0].r + ':$K$' + d.items[d.items.length - 1].r + ',TRUE,' + L(j) + d.items[0].r + ':' + L(j) + d.items[d.items.length - 1].r + ')'
            : 0;
        } else {
          v[c] = r.values[j] / S;
          var s = r.sources[j];
          b[c] = s === '上書き' ? COLOR.override : s === '実績' ? COLOR.actual : s === '実績+予定' ? COLOR.actualPlan : (j === 1 ? COLOR.now : null);
        }
        if (d.t !== 'value' && j === 1 && d.t !== 'major') b[c] = COLOR.now;
        if (d.t === 'major') w[c] = 'bold';
      }
      if (tot) {
        v[9] = '=SUM(' + nowL + d.r + ':' + lastL + d.r + ')';
        v[8] = d.t === 'value' ? '=IF(K' + d.r + ',IFERROR(J' + d.r + '/J$' + tot + ',""),"")' : '=IFERROR(J' + d.r + '/J$' + tot + ',"")';
        f[8] = '0.0%;[Red]-0.0%;0.0%';
      }
    }
    data.push(v); bg.push(b); weight.push(w); fcolor.push(fc); fmt.push(f);
  });
  bands.forEach(function (bd) {
    data[bd.first - 1][1] = bd.name;
    for (var r = bd.first; r <= bd.last; r++) {
      bg[r - 1][1] = BAND_COLOR[bd.sec]; weight[r - 1][1] = 'bold';
      if (!bg[r - 1][F - 1]) bg[r - 1][F - 1] = COLOR.pastCol; // 先月（M列）：他の色が無いセルはグレー
    }
  });

  // ---- 4) 書き込み（行構成が変わったときだけ、行・結合・グループ・チェックボックスを作り直す）
  var props = PropertiesService.getDocumentProperties();
  var sig = JSON.stringify([R.months, rows.map(function (d) {
    return d.t + '|' + (d.row ? rowKey_(d.row) : d.acc ? d.acc.id : (d.label || d.key || ''));
  })]);
  var full = props.getProperty('gridSig') !== sig || sh.getMaxRows() < rows.length || !bandsIntact_(sh, bands);
  if (full) {
    sh.setFrozenRows(0);
    sh.setFrozenColumns(0);
    var old = sh.getMaxRows();
    sh.insertRowsAfter(old, rows.length + 20);
    sh.deleteRows(1, old); // 古い行ごと削除して、結合・グループ・入力規則を一掃する
    ensureSize_(sh, rows.length + 20, W);
    var cb = rows.filter(function (d) { return d.t === 'value'; }).map(function (d) { return 'K' + d.r; });
    if (cb.length) sh.getRangeList(cb).insertCheckboxes();
  }
  var rg = sh.getRange(1, 1, rows.length, W);
  rg.setNumberFormats(fmt);
  rg.setValues(data);
  // 次回の差分検出の基準として、明細項目の行に表示した値を保存
  saveGridState_(ss, R.months, rows.filter(function (d) { return d.t === 'value'; }).map(function (d) {
    var v = data[d.r - 1];
    return [v[GRID.KEY_COL - 1], v[GRID.CHECK_COL - 1]].concat(v.slice(F - 1, F - 1 + M));
  }));
  rg.setBackgrounds(bg);
  rg.setFontWeights(weight);
  rg.setFontColors(fcolor);
  if (full) {
    bands.forEach(function (bd) {
      sh.getRange(bd.first, 2, bd.last - bd.first + 1, 1).merge()
        .setVerticalText(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
    });
    sh.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
    groups.forEach(function (g) { sh.getRange(g[0], 1, g[1], 1).shiftRowGroupDepth(1); });
    drawBorders_(sh, rows, bands, W);
    sh.setRowHeight(1, 46);
    sh.getRange(2, F, 2, M).setHorizontalAlignment('center');
    sh.showColumns(1, sh.getMaxColumns()); // 以前のレイアウトで非表示にした列をいったん戻す
    var widths = [16, 26, 18, 18, 250, 300, 140, 90, 60, 95, 40, 60];
    widths.forEach(function (px, i) { sh.setColumnWidth(i + 1, px); });
    sh.setColumnWidths(F, M, 82);
    sh.hideColumns(6, 3);              // 説明・費目対象先・タグ（模範例と同じく通常は非表示）
    sh.hideColumns(GRID.KEY_COL);
    sh.setFrozenRows(GRID.HEADER_ROWS);
    sh.setFrozenColumns(GRID.CHECK_COL);
    props.setProperty('gridSig', sig);
  }
  sh.getRange(1, F, 1, M).setHorizontalAlignment('center').setVerticalAlignment('bottom'); // 決算月の印：横中央・下揃え
  ensureLogo_(sh);
  return { sumRow: sumRow, secInfo: secInfo };
}

/**
 * C1 にロゴ画像を置く（クリックするとメニュー「再計算」と同じ処理が動く）。
 * レイアウトを作り直すと行ごと消えることがあるので、描画のたびに有無を確かめて足りなければ置き直す。
 */
var LOGO = { ALT: 'cashboard_logo', W: 210, H: 45 };
function ensureLogo_(sh) {
  if (typeof CF_LOGO_JPEG_B64 === 'undefined') return;
  var found = (sh.getImages ? sh.getImages() : []).filter(function (im) {
    try { return im.getAltTextTitle() === LOGO.ALT; } catch (x) { return false; }
  });
  found.slice(1).forEach(function (im) { im.remove(); }); // 重複していたら1つにする
  if (found.length) {
    var a = found[0].getAnchorCell();
    if (a.getRow() === 1 && a.getColumn() === 3) return;
    found[0].remove();
  }
  var blob = Utilities.newBlob(Utilities.base64Decode(CF_LOGO_JPEG_B64), 'image/jpeg', 'cashboard_logo.jpg');
  var img = sh.insertImage(blob, 3, 1, 2, 1);
  img.setAltTextTitle(LOGO.ALT).setAltTextDescription('クリックすると再計算します');
  img.setWidth(LOGO.W).setHeight(LOGO.H);
  img.assignScript('menuRecalc');
}

/**
 * B列の結合範囲が想定どおりか（手作業で行が挿入・削除されると結合やチェックボックスがずれるので、
 * ずれていたらレイアウトを作り直す）
 */
function bandsIntact_(sh, bands) {
  var want = bands.map(function (bd) { return bd.first + ':' + bd.last; }).sort().join(',');
  var got = sh.getRange(1, 2, sh.getMaxRows(), 1).getMergedRanges().map(function (m) {
    return m.getRow() + ':' + m.getLastRow();
  }).sort().join(',');
  return want === got;
}

/**
 * 罫線（見出し行より下が対象。後に引いたものが優先されるので、灰→黒→セクション帯の外枠の順に引く）
 *   ・D/E列に文字がある行：上辺・下辺を C列〜最終列 灰色の細線
 *   ・C列に文字がある行　：上辺・下辺を C列〜最終列 黒の細線
 *   ・B列の結合範囲（セクション帯）：上辺・下辺を B列〜最終列、左辺を B列 黒の細線
 */
function drawBorders_(sh, rows, bands, W) {
  var lastL = colLetter_(W), gray = [], black = [];
  rows.forEach(function (d) {
    var hasC = d.t === 'major' || (d.t === 'sum' && d.col === 3);
    var hasDE = d.t === 'minor' || d.t === 'value' || (d.t === 'sum' && d.col === 4);
    if (hasC) black.push('C' + d.r + ':' + lastL + d.r);
    else if (hasDE) gray.push('C' + d.r + ':' + lastL + d.r);
  });
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  if (gray.length) sh.getRangeList(gray).setBorder(true, null, true, null, null, null, BORDER.gray, SOLID);
  if (black.length) sh.getRangeList(black).setBorder(true, null, true, null, null, null, BORDER.black, SOLID);
  // セクション帯：B列〜最終列のブロック全体の上辺・下辺・左辺を黒の細線で囲む
  // （結合セルの一部だけに罫線を引くと下辺が表示されないことがあるため、ブロック単位で外枠を引く）
  var blocks = bands.map(function (bd) { return 'B' + bd.first + ':' + lastL + bd.last; });
  if (blocks.length) sh.getRangeList(blocks).setBorder(true, true, true, null, null, null, BORDER.black, SOLID);
}

// ---------------------------------------------------------------- ダッシュボード（模範例「Dashboard全体／入金／出金」に準拠）
/**
 * 「ダッシュボード」シートは1画面に1つのダッシュボードだけを表示する。
 *   B1 のプルダウン（全体／入金／出金）で切り替え
 *   A列のチェックボックスで、グラフに描く系列を選ぶ（A2 は全選択／全解除）
 * 表は「シミュレーター」のセルを参照する数式なので値は常に最新。
 * グラフは選択した行だけを非表示シート「_グラフ用」に並べて、そこから描く。
 */
var DASH = {
  VIEWS: ['全体', '入金', '出金'],
  IN_BAR: ['#a4c2f4', '#b6d7a8', '#6d9eeb', '#8e7cc3', '#f6b26b', '#76a5af', '#d5a6bd'],
  OUT_BAR: ['#6d9eeb', '#f6b26b', '#93c47d', '#e06666', '#8e7cc3', '#76a5af', '#d5a6bd'],
  LINE: ['#cc4125', '#e69138', '#3c78d8', '#6aa84f', '#a64d79', '#45818e', '#bf9000', '#674ea7', '#999999', '#134f5c', '#990000', '#38761d', '#c27ba0', '#0b5394'],
  HELPER: '_グラフ用', FIRST: 3, F: 3, MONTH_W: 64, LABEL_W: 210, CHART_H: 380
};

/** 再計算のたびに呼ばれる。ダッシュボードの定義（行・参照先）を保存し、変わったときだけ描き直す */
function renderDashboard_(ss, R, G) {
  var old = ss.getSheetByName('グラフ'); // 旧形式のグラフシート（自動生成のみ）は置き換える
  if (old && ss.getSheets().length > 1) ss.deleteSheet(old);
  function secRows(sec, barColors) {
    var info = G.secInfo[sec], list = [], li = 0;
    if (!info) return list;
    info.majors.forEach(function (mj, i) {
      list.push({ l: mj.label, gr: mj.r, k: 'bar', c: barColors[i % barColors.length], a: 0, m: 1 });
      mj.children.forEach(function (mn) {
        list.push({ l: mn.acc.name, gr: mn.r, k: 'line', c: DASH.LINE[li++ % DASH.LINE.length], a: 1, m: 0 });
      });
    });
    return list;
  }
  var sr = G.sumRow;
  var def = {
    months: R.months.map(function (ym) { return ym.slice(0, 4) + '/' + Number(ym.slice(5, 7)); }), unit: (R.params.displayUnit || 1) === 1000 ? 'K JPY' : 'JPY',
    views: {
      '全体': { color: BAND_COLOR['集計'], rows: [
        { l: '正味フリーキャッシュ', gr: sr.fc.r, k: 'bar', c: '#c9daf8', a: 1, m: 1 },
        { l: '出金合計', gr: sr.out.r, k: 'line', c: '#cc0000', a: 0, m: 1 },
        { l: 'burn rate', gr: sr.burn.r, k: 'line', c: '#f1c232', a: 0, m: 0 },
        { l: '入金合計', gr: sr['in'].r, k: 'line', c: '#38761d', a: 0, m: 1 }
      ] },
      '入金': { color: BAND_COLOR['入金'], rows: secRows('入金', DASH.IN_BAR) },
      '出金': { color: BAND_COLOR['出金'], rows: secRows('出金', DASH.OUT_BAR) }
    }
  };
  var sig = JSON.stringify(def);
  var props = PropertiesService.getDocumentProperties();
  if (ss.getSheetByName(SHEET.DASH) && ss.getSheetByName(DASH.HELPER) && props.getProperty('dashDef') === sig) return;
  props.setProperty('dashDef', sig);
  drawDashboard_(ss);
}

function dashState_() {
  var p = PropertiesService.getDocumentProperties();
  var def = JSON.parse(p.getProperty('dashDef') || 'null');
  var view = p.getProperty('dashView') || '全体';
  if (DASH.VIEWS.indexOf(view) < 0) view = '全体';
  var off = JSON.parse(p.getProperty('dashOff') || '{}'); // グラフから外した行 { view: { label: 1 } }
  return { def: def, view: view, off: off, props: p };
}

/** 選択中のダッシュボードの表を描き、グラフも描き直す */
function drawDashboard_(ss) {
  var st = dashState_();
  if (!st.def) return;
  var sh = getOrCreate_(ss, SHEET.DASH, 1);
  var V = st.def.views[st.view], rows = V.rows, off = st.off[st.view] || {};
  var M = st.def.months.length, F = DASH.F, W = F + M; // A:チェック B:項目 C〜:月 最後:単位
  var first = DASH.FIRST, last = first + rows.length - 1, n = Math.max(last, first);
  var gname = "'" + SHEET.GRID + "'!";
  var isAll = st.view === '全体';

  var data = [], bg = [], wt = [], fmt = [], fc = [];
  var r, c, j;
  for (r = 0; r < n; r++) {
    var v = [], b0 = [], w0 = [], f0 = [], c0 = [];
    for (c = 0; c < W; c++) { v.push(''); b0.push(null); w0.push('normal'); f0.push(NUM_FMT); c0.push('#000000'); }
    data.push(v); bg.push(b0); wt.push(w0); fmt.push(f0); fc.push(c0);
  }
  // 1行目：切り替えプルダウン（B1）と案内
  for (c = 0; c < W; c++) { bg[0][c] = V.color; fmt[0][c] = '@'; fmt[1][c] = '@'; }
  data[0][1] = st.view; wt[0][1] = 'bold'; bg[0][1] = '#ffffff';
  data[0][2] = st.view + 'Dashboard　◀ B1 のプルダウンで 全体／入金／出金 を切り替え。A列のチェックでグラフに描く系列を選択';
  wt[0][2] = 'bold';
  // 2行目：見出し
  data[1][1] = '項目（A列：グラフに表示）'; wt[1][1] = 'bold';
  for (j = 0; j < M; j++) { data[1][F - 1 + j] = st.def.months[j]; wt[1][F - 1 + j] = 'bold'; }
  data[1][W - 1] = st.def.unit;
  data[1][0] = rows.every(function (x) { return !off[x.l]; });
  fmt[1][0] = 'General'; // チェックボックス（全選択）は文字列書式にしない
  // 3行目〜：表
  rows.forEach(function (x, i) {
    var rr = first - 1 + i;
    data[rr][0] = !off[x.l];
    data[rr][1] = x.m ? x.l : '　' + x.l;
    for (j = 0; j < M; j++) data[rr][F - 1 + j] = '=' + gname + colLetter_(GRID.FIRST_MONTH_COL + j) + x.gr;
    if (x.m) for (c = 1; c < W - 1; c++) wt[rr][c] = 'bold';
    if (x.k === 'bar') {
      if (isAll) bg[rr][1] = x.c; else for (c = 1; c < W - 1; c++) bg[rr][c] = x.c;
    } else {
      fc[rr][1] = x.c; // 折れ線の系列は項目名をその線の色で表示
    }
  });

  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  var all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.clearDataValidations();
  sh.clear();
  ensureSize_(sh, n + 30, W);
  var rg = sh.getRange(1, 1, n, W);
  rg.setNumberFormats(fmt);
  rg.setValues(data);
  rg.setBackgrounds(bg);
  rg.setFontWeights(wt);
  rg.setFontColors(fc);
  sh.getRange(1, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(DASH.VIEWS, true).setAllowInvalid(false).build())
    .setHorizontalAlignment('center');
  sh.getRange(2, 1).insertCheckboxes();
  if (rows.length) sh.getRange(first, 1, rows.length, 1).insertCheckboxes();
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  var box = sh.getRange(2, 1, n - 1, W);
  box.setBorder(true, true, true, true, true, true, BORDER.gray, SOLID);
  box.setBorder(true, true, true, true, null, null, BORDER.black, SOLID);
  sh.getRange(2, F, 1, M).setHorizontalAlignment('center');
  sh.setColumnWidth(1, 28);
  sh.setColumnWidth(2, DASH.LABEL_W);
  sh.setColumnWidths(F, M, DASH.MONTH_W);
  sh.setColumnWidth(W, 50);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  drawDashChart_(ss, sh, st);
}

/** 選択された行だけでグラフを描き直す */
function drawDashChart_(ss, sh, st) {
  var V = st.def.views[st.view], off = st.off[st.view] || {}, M = st.def.months.length;
  var sel = V.rows.filter(function (x) { return !off[x.l]; });
  var hp = ss.getSheetByName(DASH.HELPER);
  if (!hp) { hp = ss.insertSheet(DASH.HELPER); hp.hideSheet(); }
  hp.clear();
  ensureSize_(hp, sel.length + 2, M + 1);
  var gname = "'" + SHEET.GRID + "'!";
  var t = [[''].concat(st.def.months)];
  sel.forEach(function (x) {
    var row = [x.l];
    for (var j = 0; j < M; j++) row.push('=' + gname + colLetter_(GRID.FIRST_MONTH_COL + j) + x.gr);
    t.push(row);
  });
  hp.getRange(1, 1, 1, M + 1).setNumberFormat('@');
  hp.getRange(1, 1, t.length, M + 1).setValues(t);

  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  if (!sel.length) return;
  var series = {};
  sel.forEach(function (x, i) {
    series[i] = x.k === 'bar'
      ? { type: 'bars', targetAxisIndex: x.a, color: x.c }
      : { type: 'line', targetAxisIndex: x.a, color: x.c, pointSize: st.view === '全体' ? 0 : 6, lineWidth: 2 };
  });
  var width = 28 + DASH.LABEL_W + DASH.MONTH_W * M + 50;
  sh.insertChart(sh.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(hp.getRange(1, 1, sel.length + 1, M + 1))
    .setTransposeRowsAndColumns(true)
    .setNumHeaders(1)
    .setOption('series', series)
    .setOption('legend', { position: 'top', maxLines: 3 })
    .setOption('vAxes', { 0: { format: '#,##0' }, 1: { format: '#,##0' } })
    .setOption('width', width).setOption('height', DASH.CHART_H)
    .setPosition(DASH.FIRST + V.rows.length + 1, 1, 0, 0)
    .build());
}

/** ダッシュボードシートの編集：B1（切り替え）・A列（系列の選択） */
function handleDashEdit_(e) {
  var r0 = e.range.getRow(), c0 = e.range.getColumn();
  var touchesSel = r0 === 1 && c0 <= 2 && c0 + e.range.getNumColumns() - 1 >= 2;
  var touchesChk = c0 === 1 && r0 + e.range.getNumRows() - 1 >= 2;
  if (!touchesSel && !touchesChk) return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;
  try {
    var ss = e.source, sh = e.range.getSheet();
    var st = dashState_();
    if (!st.def) return;
    if (touchesSel) {
      var v = String(sh.getRange(1, 2).getValue());
      if (DASH.VIEWS.indexOf(v) < 0) v = st.view;
      st.props.setProperty('dashView', v);
      drawDashboard_(ss);
      return;
    }
    var rows = st.def.views[st.view].rows, n = rows.length;
    if (!n) return;
    if (r0 === 2 && e.range.getNumRows() === 1) { // 全選択／全解除
      var on = sh.getRange(2, 1).getValue() === true, vals = [];
      for (var i = 0; i < n; i++) vals.push([on]);
      sh.getRange(DASH.FIRST, 1, n, 1).setValues(vals);
    }
    // 素早く続けてクリックされた場合に備え、チェックの状態が落ち着くまで描き直す（最大3回）
    var prev = null;
    for (var loop = 0; loop < 3; loop++) {
      var chk = sh.getRange(DASH.FIRST, 1, n, 1).getValues(), off = {};
      rows.forEach(function (x, k) { if (chk[k][0] !== true) off[x.l] = 1; });
      var key = JSON.stringify(off);
      if (key === prev) break;
      prev = key;
      st.off[st.view] = off;
      st.props.setProperty('dashOff', JSON.stringify(st.off));
      drawDashChart_(ss, sh, st);
      sh.getRange(2, 1).setValue(Object.keys(off).length === 0);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}

// ================================================================ 月初処理（表を1か月進める）
/*
 * 「今月」はパラメータ「基準月」で決まる。カレンダーの月が変わっても勝手には進まず、
 * メニュー「月初処理」を実行したときだけ1か月進める。
 *   1) 今月（N列）の値を、その月の実績として確定する（明細シートに「実績」を追加）
 *        → 進めたあとの先月（M列）には、処理前のN列の値がそのまま入る
 *   2) 月初残高（今月初の実際の残高）をユーザーが入力する
 *   3) 基準月を1か月進める（年月の見出しがすべて＋1か月）
 *   4) 新しく増える最終月（AE列）は各項目のルール（固定額・人件費・連動・比率など）で推計する。
 *      処理のあとシミュレーター上で直接直せる（直した値は「上書き」として保存される）
 */

/** パラメータ「基準月」が空欄なら、今日の月で固定する（カレンダーが変わっても勝手に進まないように） */
function pinBaseMonth_(ss) {
  var sh = ss.getSheetByName('パラメータ');
  if (!sh || sh.getLastRow() < 1) return;
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() !== '基準月') continue;
    if (CFEngine.normalizeYm(values[i][1])) return;
    sh.getRange(i + 1, 2).setNumberFormat('@').setValue(todayYm_(ss));
    return;
  }
  var last = sh.getLastRow() + 1;
  sh.getRange(last, 1, 1, 3).setValues([['基準月', '', '「今月」とする月（YYYY-MM）。メニュー「月初処理」で1か月ずつ進める']]);
  sh.getRange(last, 2).setNumberFormat('@').setValue(todayYm_(ss));
}

function setBaseMonth_(ss, ym) {
  var sh = ss.getSheetByName('パラメータ');
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === '基準月') { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(ym); return; }
  }
  throw new Error('パラメータシートに「基準月」の行がありません');
}

/** 今月の値を実績として確定する対象（ユーザーが入力する明細項目と連動行。比率・自動計算・月初残高は除く） */
function freezableRow_(r) {
  if (r.section === '月初残高') return false;
  return (r.kind === 'item' && r.calcType !== '比率') || r.kind === 'gen';
}

/** ダイアログに出す内容を作る（処理前に一度再計算して、未反映の編集を取り込んでおく） */
function monthStartInfo_() {
  var ss = SpreadsheetApp.getActive();
  if (!runRecalc_(60000)) throw new Error('再計算が実行中です。少し待ってからもう一度実行してください');
  var data = loadData_(ss);
  var R = CFEngine.compute(data);
  var next = CFEngine.addMonths(R.base, 1);
  var S = R.params.displayUnit || 1;
  data.params['基準月'] = next;
  var N = CFEngine.compute(data); // 進めたあとの姿（新しい最終月の推計値を出すため）
  var lastYm = N.months[N.months.length - 1];
  var T = R.totals;
  return {
    from: R.base, to: next, fromLabel: ymLabel_(R.base), toLabel: ymLabel_(next), lastLabel: ymLabel_(lastYm),
    calendar: todayYm_(ss), unit: S === 1000 ? '千円' : '円', S: S,
    estimate: Math.round(T.closing[1]), // シミュレーター上の今月末（＝来月初）の現金・預金の見込み
    openings: R.rows.filter(function (r) { return r.section === '月初残高'; }).map(function (r) {
      return { key: r.key, name: r.itemName, account: r.accountName, amount: Math.round(r.values[0]), included: r.included };
    }),
    last: N.rows.filter(function (r) { return freezableRow_(r) || r.kind === 'sys'; }).map(function (r) {
      return { key: r.key, name: r.itemName, account: r.accountName, section: r.section, kind: r.kind,
        value: Math.round(r.values[r.values.length - 1] / S * 1000) / 1000, editable: r.kind !== 'sys' };
    })
  };
}

/**
 * メニュー「月初処理」。確認・入力はすべてスクリプト側の画面（ui.alert / ui.prompt）で行う。
 * ※ HTMLダイアログのボタンからスクリプトを呼ぶ方式（google.script.run）は、ブラウザで複数の
 *   Googleアカウントにログインしていると PERMISSION_DENIED で失敗するため使わない。
 */
function menuMonthStart() {
  var ui = SpreadsheetApp.getUi();
  var info;
  try { info = monthStartInfo_(); } catch (err) { ui.alert(String(err && err.message || err)); return; }
  var fmt = function (n) { return Number(n).toLocaleString('ja-JP'); };
  var title = '月初処理：' + info.fromLabel + ' → ' + info.toLabel;

  // 1) 確認
  var msg = '今月を ' + info.fromLabel + ' から ' + info.toLabel + ' に進めます。\n\n' +
    '・今月（N列）の値を「実績」として確定し、先月（M列）に移します\n' +
    '・年月の見出しがすべて1か月進みます\n' +
    '・新しい最終月（' + info.lastLabel + '）は各項目のルール（固定額・人件費・連動・比率など）で推計します\n' +
    '　（推計値は、処理のあとシミュレーター上で直接直せます。直した値は「上書き」になります）\n' +
    '・次の画面で、口座ごとの月初残高（' + info.toLabel + '初時点の実際の残高）を入力します\n';
  if (info.to > info.calendar) msg += '\n⚠ カレンダー上の今月（' + ymLabel_(info.calendar) + '）より先の月に進めようとしています。月が替わってから実行するのが通常です。\n';
  msg += '\n続けますか？';
  if (ui.alert(title, msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // 2) 月初残高の入力（口座ごと）
  var openings = [];
  for (var i = 0; i < info.openings.length; i++) {
    var o = info.openings[i], value = null, note = '';
    while (value === null) {
      var res = ui.prompt(title + '（月初残高 ' + (i + 1) + '/' + info.openings.length + '）',
        note + o.account + '：' + o.name + (o.included ? '' : '（計上✗）') + '\n\n' +
        info.toLabel + '初時点の実際の残高を「円」で入力してください。\n' +
        '処理前の値：' + fmt(o.amount) + ' 円\n' +
        '（参考：シミュレーター上の見込み 現金・預金合計 ' + fmt(info.estimate) + ' 円）\n\n' +
        '空欄のまま OK を押すと、処理前の値をそのまま使います。', ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() !== ui.Button.OK) { ui.alert('月初処理を中止しました（何も変更していません）'); return; }
      var t = String(res.getResponseText()).replace(/[,，\s円]/g, '').replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
      if (t === '') value = o.amount;
      else if (/^-?\d+(\.\d+)?$/.test(t)) value = Math.round(Number(t));
      else note = '⚠ 数値として読めませんでした：「' + res.getResponseText() + '」\n\n';
    }
    openings.push({ key: o.key, amount: value, label: o.name });
  }

  // 3) 最終確認
  var sum = '次の内容で実行します。\n\n今月：' + info.fromLabel + ' → ' + info.toLabel + '\n\n月初残高：\n' +
    openings.map(function (x) { return '　' + x.label + '：' + fmt(x.amount) + ' 円'; }).join('\n') +
    '\n\nこの処理は元に戻せません（明細シートに実績行が追加され、基準月が進みます）。実行しますか？';
  if (ui.alert(title, sum, ui.ButtonSet.YES_NO) !== ui.Button.YES) { ui.alert('月初処理を中止しました（何も変更していません）'); return; }

  try {
    var done = applyMonthStart({ from: info.from, openings: openings, last: [] });
    ui.alert(title, done + '\n\n新しい最終月（' + info.lastLabel + '）の値は、シミュレーター上で直接直せます。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('月初処理に失敗しました：' + String(err && err.message || err));
  }
}

/** 月初処理の本体（menuMonthStart から呼ぶ。p.last で最終月の上書きも渡せる） */
function applyMonthStart(p) {
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(60000)) throw new Error('再計算が実行中です。少し待ってからもう一度実行してください');
  try {
    captureGridEdits_(ss);
    var R = CFEngine.compute(loadData_(ss));
    if (R.base !== p.from) throw new Error('すでに月初処理が行われています（今月：' + ymLabel_(R.base) + '）');
    var now = R.months[1], next = CFEngine.addMonths(R.base, 1), S = R.params.displayUnit || 1;
    var user = userEmail_(null);

    // 1) 今月の値を実績として確定（既にある実績との差額を追加）
    var actual = {};
    R.long.forEach(function (x) { if (x.ym === now && x.source === '実績') actual[x.key] = (actual[x.key] || 0) + x.amount; });
    var freeze = [];
    R.rows.forEach(function (r) {
      if (!freezableRow_(r)) return;
      var delta = Math.round(r.values[1] - (actual[r.key] || 0));
      if (delta !== 0) freeze.push({ key: r.key, kind: '実績', ym: now, amount: delta, memo: '月初処理：' + ymLabel_(now) + 'の値を実績として確定' });
    });
    appendDetails_(ss, freeze, user);

    // 2) 月初残高
    updateItems_(ss, (p.openings || []).map(function (o) { return { id: String(o.key), header: '金額（税抜）', value: Math.round(Number(o.amount)) }; }));

    // 3) 基準月を進める
    setBaseMonth_(ss, next);

    // 4) 新しい最終月の値を直した行は「上書き」として保存
    var lastYm = CFEngine.addMonths(next, CFEngine.WINDOW - 2);
    var changes = (p.last || []).filter(function (x) { return isFinite(Number(x.value)); })
      .map(function (x) { return { key: String(x.key), ym: lastYm, value: Math.round(Number(x.value) * S) }; });
    if (changes.length) applyOverrides_(ss, changes, user);

    logMonthStart_(ss, user, p.from, next, freeze.length, changes.length);
    recalc_();
    PropertiesService.getDocumentProperties().deleteProperty('dirty');
    return '月初処理が完了しました。\n\n今月：' + ymLabel_(next) + '\n実績として確定した行：' + freeze.length + ' 行\n最終月を直した行：' + changes.length + ' 行';
  } finally {
    lock.releaseLock();
  }
}

/** 明細シートに行を追加する（list: {key, kind, ym, amount, memo}） */
function appendDetails_(ss, list, user) {
  if (!list.length) return;
  var sh = ss.getSheetByName(SHEET.DETAIL);
  var values = sh.getDataRange().getValues();
  var width = values[0].length, H = headerIndex_(values[0]), now = new Date();
  var rows = list.map(function (x) {
    var row = [];
    for (var i = 0; i < width; i++) row.push('');
    row[H['明細ID']] = newDetailId_();
    row[H['項目ID']] = x.key;
    row[H['種別']] = x.kind;
    row[H['入出金月']] = x.ym;
    row[H['金額（税抜）']] = x.amount;
    row[H['摘要']] = x.memo;
    row[H['登録者']] = user;
    row[H['登録日時']] = now;
    return row;
  });
  var start = sh.getLastRow() + 1;
  ensureSize_(sh, start + rows.length, width);
  sh.getRange(start, H['入出金月'] + 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, width).setValues(rows);
}

function logMonthStart_(ss, user, from, to, nFreeze, nOver) {
  var sh = ss.getSheetByName(SHEET.LOG);
  if (!sh) return;
  var r = sh.getLastRow() + 1;
  ensureSize_(sh, r, 6);
  sh.getRange(r, 1, 1, 6).setValues([[new Date(), user, '月初処理', from + ' → ' + to, '実績確定 ' + nFreeze + ' 行', '最終月の上書き ' + nOver + ' 行']]);
}

// ================================================================ シミュレーター上の直接入力
/**
 * 編集トリガーは「未反映あり」の印を付けるだけにする（すぐ終わるので連続入力を妨げない）。
 * 編集内容そのものは、次の再計算の直前に captureGridEdits_ が差分から取り込む。
 * ただし1セルの明らかな誤入力（先月の列・小計行・数値以外）は、その場で元に戻す。
 */
function handleGridEdit_(e) {
  var ss = e.source, sh = e.range.getSheet();
  var r0 = e.range.getRow(), c0 = e.range.getColumn();
  markDirty_(ss);
  if (e.range.getNumRows() * e.range.getNumColumns() !== 1 || r0 <= GRID.HEADER_ROWS) return;
  var F = GRID.FIRST_MONTH_COL, CK = GRID.CHECK_COL;
  var k = parseKey_(sh.getRange(r0, GRID.KEY_COL).getDisplayValue());
  var editable = false, msg = '';
  if (c0 === CK) editable = k.type === 'item' || k.type === 'open';
  else if (c0 >= F && c0 < F + CFEngine.WINDOW) {
    var mi = c0 - F, v = e.range.getValue();
    if (k.type === 'open') editable = mi === 0;
    else if (k.type === 'item' || k.type === 'gen' || k.type === 'sys') {
      if (mi === 0) msg = '先月（実績月）は直接入力できません。「明細」シートに実績として登録してください';
      else if (v !== '' && CFEngine.num(v, null) === null) msg = '数値以外は入力できません';
      else editable = true;
    }
  }
  if (editable) return;
  var ov = e.oldValue === undefined ? '' : e.oldValue;
  if (c0 === CK) ov = String(ov).toUpperCase() === 'TRUE';
  else if (ov !== '' && CFEngine.num(ov, null) !== null) ov = CFEngine.num(ov, 0);
  e.range.setValue(ov);
  if (msg) ss.toast(msg, '入力できません', 8);
}

function headerIndex_(header) {
  var H = {};
  header.forEach(function (h, i) { H[String(h).trim()] = i; });
  return H;
}

/** 項目シートの指定列（有効・金額）を更新する */
function updateItems_(ss, updates) {
  var sh = ss.getSheetByName(SHEET.ITEM);
  var values = sh.getDataRange().getValues();
  var H = headerIndex_(values[0]);
  var rowOf = {};
  for (var r = 1; r < values.length; r++) rowOf[String(values[r][H['項目ID']]).trim()] = r + 1;
  updates.forEach(function (u) {
    if (!rowOf[u.id] || H[u.header] === undefined) return;
    sh.getRange(rowOf[u.id], H[u.header] + 1).setValue(u.value);
  });
}

function applyOverrides_(ss, changes, user) {
  var sh = ss.getSheetByName(SHEET.DETAIL);
  if (!sh) throw new Error('「明細」シートがありません。メニューの「初期設定」を実行してください');
  var values = sh.getDataRange().getValues();
  var width = values[0].length;
  var H = headerIndex_(values[0]);
  ['明細ID', '項目ID', '種別', '入出金月', '金額（税抜）', '摘要', '登録者', '登録日時'].forEach(function (h) {
    if (H[h] === undefined) throw new Error('「明細」シートに見出し「' + h + '」がありません');
  });
  var idx = {};
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][H['種別']]).indexOf('上書') < 0) continue;
    idx[String(values[r][H['項目ID']]).trim() + '|' + CFEngine.normalizeYm(values[r][H['入出金月']])] = r + 1;
  }
  var now = new Date(), toDelete = [], toAppend = [];
  changes.forEach(function (c) {
    var rowNo = idx[c.key + '|' + c.ym];
    if (c.value === null) { if (rowNo > 0) toDelete.push(rowNo); return; }
    if (rowNo > 0) {
      sh.getRange(rowNo, H['金額（税抜）'] + 1).setValue(c.value);
      sh.getRange(rowNo, H['登録者'] + 1).setValue(user);
      sh.getRange(rowNo, H['登録日時'] + 1).setValue(now);
      return;
    }
    var row = [];
    for (var i = 0; i < width; i++) row.push('');
    row[H['明細ID']] = newDetailId_();
    row[H['項目ID']] = c.key;
    row[H['種別']] = '上書き';
    row[H['入出金月']] = c.ym;
    row[H['金額（税抜）']] = c.value;
    row[H['摘要']] = 'シミュレーター上で上書き';
    row[H['登録者']] = user;
    row[H['登録日時']] = now;
    toAppend.push(row);
    idx[c.key + '|' + c.ym] = -1; // 同じ編集内での重複追加を防ぐ
  });
  if (toAppend.length) {
    var start = sh.getLastRow() + 1;
    ensureSize_(sh, start + toAppend.length, width);
    sh.getRange(start, H['入出金月'] + 1, toAppend.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, toAppend.length, width).setValues(toAppend);
  }
  toDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
}

function newDetailId_() {
  return 'D' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMddHHmmss') + Utilities.getUuid().slice(0, 4);
}

/** 明細シートの新しい行に明細ID・登録者・登録日時を自動で入れる */
function stampDetailRows_(e) {
  var sh = e.range.getSheet();
  var r0 = Math.max(2, e.range.getRow()), r1 = e.range.getRow() + e.range.getNumRows() - 1;
  if (r1 < r0) return;
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var H = headerIndex_(header);
  if (H['明細ID'] === undefined || H['項目ID'] === undefined) return;
  var rows = sh.getRange(r0, 1, r1 - r0 + 1, header.length).getValues();
  var user = userEmail_(e), now = new Date();
  rows.forEach(function (row, i) {
    if (!String(row[H['項目ID']]).trim() || String(row[H['明細ID']]).trim()) return;
    var r = r0 + i;
    sh.getRange(r, H['明細ID'] + 1).setValue(newDetailId_());
    if (H['登録者'] !== undefined) sh.getRange(r, H['登録者'] + 1).setValue(user);
    if (H['登録日時'] !== undefined) sh.getRange(r, H['登録日時'] + 1).setValue(now);
  });
}

function logChanges_(ss, changes, user, oldValue) {
  var sh = ss.getSheetByName(SHEET.LOG) || createTable_(ss, SHEET.LOG, ['日時', 'ユーザー', '項目キー', '年月', '変更前', '変更後']);
  var now = new Date();
  var rows = changes.map(function (c) {
    return [now, user, c.key, c.ym, oldValue === undefined ? '' : oldValue, c.value === null ? '（上書き解除）' : c.value];
  });
  var start = sh.getLastRow() + 1;
  ensureSize_(sh, start + rows.length, 6);
  sh.getRange(start, 4, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, 6).setValues(rows);
}

function createTable_(ss, name, header) {
  var sh = getOrCreate_(ss, name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// ================================================================ 内訳・警告・スナップショット
function showBreakdown() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sh.getName() !== SHEET.GRID) { ui.alert('「シミュレーター」シートで、項目行の月のセルを選択してから実行してください'); return; }
  var cell = sh.getActiveCell();
  var mi = cell.getColumn() - GRID.FIRST_MONTH_COL;
  var k = parseKey_(sh.getRange(cell.getRow(), GRID.KEY_COL).getDisplayValue());
  if (mi < 0 || mi >= CFEngine.WINDOW || ['item', 'gen', 'sys'].indexOf(k.type) < 0 || cell.getRow() <= GRID.HEADER_ROWS) {
    ui.alert('明細項目の行（E列に項目名がある行）の月のセルを選択してください'); return;
  }
  var ym = CFEngine.normalizeYm(sh.getRange(2, cell.getColumn()).getDisplayValue());
  var name = sh.getRange(cell.getRow(), 5).getDisplayValue();
  var calc = ss.getSheetByName(SHEET.CALC).getDataRange().getValues();
  var H = headerIndex_(calc[0]);
  var list = calc.slice(1).filter(function (r) { return String(r[H['項目キー']]) === k.id && String(r[H['年月']]) === ym; });
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var total = 0;
  var trs = list.map(function (r) {
    total += Number(r[H['金額']]) || 0;
    return '<tr><td>' + esc(r[H['ソース']]) + '</td><td class="n">' + Number(r[H['金額']]).toLocaleString() + '</td><td>' +
      esc(r[H['計算根拠']]) + (r[H['摘要']] ? '<br><small>' + esc(r[H['摘要']]) + '</small>' : '') + '</td><td>' + esc(r[H['発生月']]) + '</td></tr>';
  }).join('');
  var html = '<style>body{font:13px sans-serif}table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid #ddd;padding:4px;vertical-align:top}' +
    '.n{text-align:right;white-space:nowrap}small{color:#666}</style>' +
    '<p><b>' + esc(name) + '</b><br>' + esc(ym) + '　合計 <b>' + total.toLocaleString() + '</b> 円（税抜）</p>' +
    (list.length ? '<table><tr><th>ソース</th><th>金額（円）</th><th>根拠</th><th>発生月</th></tr>' + trs + '</table>' : '<p>このセルの構成要素はありません。</p>') +
    '<p><small>黄色のセル＝上書き。セルを空にするとルールの値に戻ります。</small></p>';
  ui.showSidebar(HtmlService.createHtmlOutput(html).setTitle('内訳：' + name + ' ' + ym));
}

function showWarnings() {
  var w = JSON.parse(PropertiesService.getDocumentProperties().getProperty('warnings') || '[]');
  SpreadsheetApp.getUi().alert(w.length ? '警告（' + w.length + ' 件）\n\n・' + w.join('\n・') : '警告はありません');
}

function saveSnapshot() {
  var ss = SpreadsheetApp.getActive(), R = null;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(60000)) { ss.toast('再計算が実行中です。少し待ってからもう一度実行してください', 'スナップショット', 6); return; }
  try {
    var seq = PropertiesService.getDocumentProperties().getProperty('editSeq');
    R = recalc_();
    if (PropertiesService.getDocumentProperties().getProperty('editSeq') === seq) PropertiesService.getDocumentProperties().deleteProperty('dirty');
  } finally { lock.releaseLock(); }
  var now = new Date(), user = userEmail_(null);
  var id = 'S' + Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmmss');
  var header = ['スナップショットID', '保存日時', '保存者', '基準月', '項目キー', '科目', '項目名', '年月', '金額'];
  var sh = ss.getSheetByName(SHEET.SNAP) || createTable_(ss, SHEET.SNAP, header);
  var rows = [];
  R.rows.forEach(function (r) {
    if (!r.included) return;
    R.months.forEach(function (m, j) { if (r.values[j]) rows.push([id, now, user, R.base, r.key, r.accountName, r.itemName, m, r.values[j]]); });
  });
  [['#in', '入金合計', 'inflow'], ['#out', '出金合計', 'outflow'], ['#fc', '正味フリーキャッシュ', 'fc']].forEach(function (t) {
    R.months.forEach(function (m, j) { rows.push([id, now, user, R.base, t[0], '', t[1], m, R.totals[t[2]][j]]); });
  });
  var start = sh.getLastRow() + 1;
  ensureSize_(sh, start + rows.length, header.length);
  sh.getRange(start, 4, rows.length, 5).setNumberFormat('@');
  sh.getRange(start, 9, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(start, 1, rows.length, header.length).setValues(rows);
  ss.toast('スナップショット ' + id + ' を保存しました（' + rows.length + ' 行）', 'スナップショット', 5);
}

// ================================================================ 初期設定
function setup() {
  var ss = SpreadsheetApp.getActive(), ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (x) { ui = null; } // スクリプトエディタから実行した場合はUIなし
  var hasData = ss.getSheetByName('項目') && ss.getSheetByName('項目').getLastRow() > 1;
  var withSample = !hasData;
  if (hasData && ui) {
    var ans = ui.alert('初期設定',
      '入力シートにデータがあります。\n\n「はい」：サンプルデータで作り直す（科目・項目・明細などの既存データは消えます）\n「いいえ」：既存データは残し、シートの体裁・入力規則・トリガーだけ整える',
      ui.ButtonSet.YES_NO_CANCEL);
    if (ans === ui.Button.CANCEL || ans === ui.Button.CLOSE) return;
    withSample = ans === ui.Button.YES;
  }
  var sample = cfBuildSample(todayYm_(ss));
  Object.keys(CF_SCHEMA).forEach(function (name) { setupTable_(ss, name, withSample ? sample[name] : null); });
  applyValidations_(ss);
  [SHEET.LOG, SHEET.SNAP].forEach(function (n) {
    if (!ss.getSheetByName(n)) createTable_(ss, n, n === SHEET.LOG
      ? ['日時', 'ユーザー', '項目キー', '年月', '変更前', '変更後']
      : ['スナップショットID', '保存日時', '保存者', '基準月', '項目キー', '科目', '項目名', '年月', '金額']);
  });
  getOrCreate_(ss, SHEET.GRID, 0);
  getOrCreate_(ss, SHEET.DASH, 1);
  var calc = getOrCreate_(ss, SHEET.CALC);
  protectWarningOnly_(calc, '計算結果（自動生成）。手で編集しないでください');
  protectWarningOnly_(ss.getSheetByName(SHEET.LOG), '変更ログ（自動記録）');
  installTriggers_(ss);
  var defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  PropertiesService.getDocumentProperties().deleteProperty('gridSig');
  runRecalc_(60000);
  ss.setActiveSheet(ss.getSheetByName(SHEET.GRID));
  if (ui) ui.alert('初期設定が完了しました。\n\n・「シミュレーター」シートの月のセルに直接入力すると「上書き」として保存されます\n・K列（計上）のチェックを外すと、その項目は合計から外れます\n・編集はすぐ記録され、メニュー「再計算」を実行するか、操作をやめて約1分たつと集計に反映されます');
}

function setupTable_(ss, name, rows) {
  var cols = CF_SCHEMA[name];
  var sh = getOrCreate_(ss, name);
  var header = cols.map(function (c) { return c[0]; });
  if (rows) {
    sh.clear();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  }
  ensureSize_(sh, Math.max(200, (rows ? rows.length : 0) + 100), header.length);
  var maxRows = sh.getMaxRows();
  cols.forEach(function (c, i) {
    if (c[2]) sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat(c[2]);
    sh.setColumnWidth(i + 1, c[3] || 100);
  });
  if (rows || sh.getLastRow() === 0) sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground(COLOR.header).setFontColor(COLOR.headerFont);
  sh.setFrozenRows(1);
  if (rows && rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function applyValidations_(ss) {
  function list(values) { return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build(); }
  function ref(sheet, allowInvalid) {
    return SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(sheet).getRange('A2:A'), true)
      .setAllowInvalid(!!allowInvalid).build();
  }
  var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  function col(sheet, header, rule) {
    var sh = ss.getSheetByName(sheet);
    var idx = CF_SCHEMA[sheet].map(function (c) { return c[0]; }).indexOf(header);
    if (!sh || idx < 0) return;
    sh.getRange(2, idx + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  }
  col('科目', '区分', list(CF_CHOICES.section));
  col('科目', '財務区分', list(CF_CHOICES.finance));
  col('科目', '既定税区分', list(CF_CHOICES.taxCodes));
  col('項目', '科目ID', ref('科目'));
  col('項目', '計算種別', list(CF_CHOICES.calcType));
  col('項目', '周期', list(CF_CHOICES.cycle));
  col('項目', '取引先ID', ref('取引先'));
  col('項目', '税区分', list(CF_CHOICES.taxCodes));
  col('項目', '有効', checkbox);
  col('人件費', '項目ID', ref('項目'));
  col('人件費', '支給', list(CF_CHOICES.payTiming));
  col('連動ルール', '条件科目ID', ref('科目'));
  col('連動ルール', '生成科目ID', ref('科目'));
  col('連動ルール', '税区分', list(CF_CHOICES.taxCodes));
  col('連動ルール', 'タイミング', list(CF_CHOICES.timing));
  col('連動ルール', '有効', checkbox);
  col('明細', '項目ID', ref('項目', true)); // 連動行（L:…）・自動計算行（SYS:…）のキーも入れられるよう警告のみ
  col('明細', '種別', list(CF_CHOICES.kind));
  col('明細', '税区分', list(CF_CHOICES.taxCodes));
}

function protectWarningOnly_(sh, desc) {
  if (!sh) return;
  var existing = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (existing.length) return;
  sh.protect().setDescription(desc).setWarningOnly(true);
}

function installTriggers_(ss) {
  var handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (handlers.indexOf('onEditInstalled') < 0) ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  if (handlers.indexOf('autoRecalc') < 0) ScriptApp.newTrigger('autoRecalc').timeBased().everyMinutes(1).create();
}
