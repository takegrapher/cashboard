/**
 * Schema.js — DBシート（テーブル）の定義
 *
 * 各シートの1行目が見出し。列の順番は入れ替えても、見出し名で読み取るので動作する。
 * [見出し, エンジン側のフィールド名, 列の書式('@'=テキスト), 幅]
 */
var CF_SCHEMA = {
  '科目': [
    ['科目ID', 'id', '@', 70], ['区分', 'section', '@', 90], ['大分類', 'major', '@', 130], ['中分類', 'name', '@', 200],
    ['財務区分', 'finance', '@', 70], ['並び順', 'order', '0', 60], ['既定税区分', 'defaultTax', '@', 110],
    ['説明', 'desc', '@', 320], ['備考', 'note', '@', 200]
  ],
  '取引先': [
    ['取引先ID', 'id', '@', 80], ['取引先名', 'name', '@', 180], ['締め日', 'closing', '@', 70],
    ['サイト（月）', 'site', '0', 80], ['備考', 'note', '@', 260]
  ],
  '項目': [
    ['項目ID', 'id', '@', 70], ['科目ID', 'accountId', '@', 70], ['項目名', 'name', '@', 180],
    ['計算種別', 'calcType', '@', 80], ['金額（税抜）', 'amount', '#,##0', 110], ['周期', 'cycle', '@', 70],
    ['指定月', 'cycleMonths', '@', 80], ['開始月', 'start', '@', 80], ['終了月', 'end', '@', 80],
    ['取引先ID', 'partnerId', '@', 80], ['税区分', 'taxCode', '@', 110], ['タグ', 'tags', '@', 110],
    ['比率対象', 'ratioTargets', '@', 100], ['比率', 'ratio', '0.0%', 70], ['有効', 'active', '', 50],
    ['説明', 'desc', '@', 240]
  ],
  '人件費': [
    ['項目ID', 'itemId', '@', 70], ['月給（額面）', 'gross', '#,##0', 110], ['支給', 'payTiming', '@', 60],
    ['賞与月', 'bonusMonths', '@', 70], ['賞与額（1回）', 'bonus', '#,##0', 110],
    ['社保会社負担率', 'co', '0.0%', 100], ['社保本人負担率', 'ee', '0.0%', 100], ['源泉税率', 'wh', '0.0%', 80],
    ['住民税（月額）', 'resident', '#,##0', 100], ['備考', 'note', '@', 240]
  ],
  '連動ルール': [
    ['ルールID', 'id', '@', 70], ['条件科目ID', 'condAccountId', '@', 80], ['条件タグ', 'condTag', '@', 90],
    ['生成科目ID', 'genAccountId', '@', 80], ['生成項目名', 'nameTpl', '@', 200], ['金額（税抜）', 'amount', '#,##0', 100],
    ['税区分', 'taxCode', '@', 110], ['タイミング', 'timing', '@', 80], ['ずらし（月）', 'offset', '0', 80],
    ['有効', 'active', '', 50], ['備考', 'note', '@', 260]
  ],
  '明細': [
    ['明細ID', 'id', '@', 110], ['項目ID', 'itemId', '@', 120], ['種別', 'kind', '@', 60],
    ['発生日', 'date', 'yyyy-mm-dd', 95], ['入出金月', 'cashYm', '@', 80], ['金額（税抜）', 'amount', '#,##0', 110],
    ['税区分', 'taxCode', '@', 110], ['摘要', 'memo', '@', 240], ['登録者', 'user', '@', 160],
    ['登録日時', 'createdAt', 'yyyy-mm-dd hh:mm', 130]
  ],
  'パラメータ': [
    ['キー', 'key', '@', 170], ['値', 'value', '', 130], ['説明', 'note', '@', 520]
  ]
};

// 入力規則（プルダウン）の選択肢
var CF_CHOICES = {
  taxCodes: ['課税10%', '課税8%（軽減）', '非課税', '対象外'],
  section: ['月初残高', '現金同等物', '入金', '出金'],
  finance: ['営業', '出資', '融資'],
  calcType: ['明細', '固定額', '人件費', '比率'],
  cycle: ['毎月', '指定月', '一括'],
  payTiming: ['当月', '翌月'],
  timing: ['開始月', '毎月', '終了月'],
  kind: ['実績', '予定', '上書き']
};

/** シートの2次元配列（1行目=見出し）を、エンジン用のオブジェクト配列に変換する */
function cfRowsToObjects(table, values) {
  var cols = CF_SCHEMA[table];
  if (!cols || !values || values.length < 2) return [];
  var header = values[0].map(function (h) { return String(h).trim(); });
  var map = [];
  cols.forEach(function (c) {
    var idx = header.indexOf(c[0]);
    if (idx >= 0) map.push([idx, c[1]]);
  });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var empty = row.every(function (v) { return v === '' || v === null || v === undefined; });
    if (empty) continue;
    var o = { _row: r + 1 };
    map.forEach(function (m) { o[m[1]] = row[m[0]]; });
    out.push(o);
  }
  return out;
}

if (typeof module !== 'undefined') module.exports = { CF_SCHEMA: CF_SCHEMA, CF_CHOICES: CF_CHOICES, cfRowsToObjects: cfRowsToObjects };
