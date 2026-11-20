/* エンジンの検証（Node.js）: node test/engine.test.js */
var assert = require('assert');
var path = require('path');
var src = path.join(__dirname, '..', 'src');
global.CFEngine = require(path.join(src, 'Engine.js')).CFEngine;
var S = require(path.join(src, 'Schema.js'));
var cfBuildSample = require(path.join(src, 'Sample.js')).cfBuildSample;
var E = global.CFEngine;

var BASE = '2026-11';

function toInput(sample, paramOverrides, extraDetails) {
  function objs(table) {
    var header = S.CF_SCHEMA[table].map(function (c) { return c[0]; });
    return S.cfRowsToObjects(table, [header].concat(sample[table]));
  }
  var params = {};
  objs('パラメータ').forEach(function (p) { params[p.key] = p.value; });
  params['基準月'] = BASE;
  Object.keys(paramOverrides || {}).forEach(function (k) { params[k] = paramOverrides[k]; });
  return {
    todayYm: BASE, params: params,
    accounts: objs('科目'), partners: objs('取引先'), items: objs('項目'), payroll: objs('人件費'),
    links: objs('連動ルール'), details: objs('明細').concat(extraDetails || [])
  };
}
function run(paramOverrides, extraDetails, mutate) {
  var sample = cfBuildSample(BASE);
  if (mutate) mutate(sample);
  return E.compute(toInput(sample, paramOverrides, extraDetails));
}
function row(R, key) { return R.rows.filter(function (r) { return r.key === key; })[0]; }
function at(R, key, ym) { var r = row(R, key); return r ? r.values[R.months.indexOf(ym)] : undefined; }
function src_(R, key, ym) { var r = row(R, key); return r.sources[R.months.indexOf(ym)]; }

var tests = [];
function test(name, fn) { tests.push([name, fn]); }

var R = run();

test('表示期間は先月〜17ヶ月後の19ヶ月', function () {
  assert.strictEqual(R.months.length, 19);
  assert.strictEqual(R.months[0], '2026-10');
  assert.strictEqual(R.months[1], '2026-11');
  assert.strictEqual(R.months[18], '2028-04');
});

test('警告なし', function () { assert.deepStrictEqual(R.warnings, []); });

test('連動：エンジニア_2の入社月（2027-04）に貸与PC 30万円が1回だけ', function () {
  var r = row(R, 'L:R01:I022');
  assert.ok(r, '貸与PC行がない');
  assert.strictEqual(at(R, 'L:R01:I022', '2027-04'), 300000);
  assert.strictEqual(r.values.reduce(function (a, b) { return a + b; }, 0), 300000);
});

test('連動：開始月が空欄（既存）のエンジニア_1にはPC行が出ない', function () {
  assert.strictEqual(row(R, 'L:R01:I021'), undefined);
});

test('連動：SaaSライセンスは人件費の項目ごと・毎月（先月は実績のみ）', function () {
  assert.strictEqual(at(R, 'L:R02:I020', '2026-10'), 5000);
  assert.strictEqual(src_(R, 'L:R02:I020', '2026-10'), '実績');
  assert.strictEqual(at(R, 'L:R02:I020', '2027-01'), 5000);
  assert.strictEqual(at(R, 'L:R02:I022', '2027-03'), 0);
  assert.strictEqual(at(R, 'L:R02:I022', '2027-04'), 5000);
});

test('連動：元の項目を無効にすると連動行も消える', function () {
  var R2 = run(null, null, function (s) { s['項目'].forEach(function (r) { if (r[0] === 'I022') r[14] = false; }); });
  assert.strictEqual(row(R2, 'L:R01:I022'), undefined);
  assert.strictEqual(row(R2, 'L:R02:I022'), undefined);
});

test('入金サイト：顧客B（翌々月末）2026-10請求は2026-12入金', function () {
  assert.strictEqual(at(R, 'I002', '2026-12'), 800000);
});

test('入金サイト：顧客A（翌月末）先月は実績、今月は先月請求分の予定', function () {
  assert.strictEqual(at(R, 'I001', '2026-10'), 1500000);
  assert.strictEqual(src_(R, 'I001', '2026-10'), '実績');
  assert.strictEqual(at(R, 'I001', '2026-11'), 1600000);
  assert.strictEqual(src_(R, 'I001', '2026-11'), '予定');
});

test('取引先の締め日：20日締めで21日以降の請求は1ヶ月後ろへずれる', function () {
  var R2 = run(null, [{ id: 'T1', itemId: 'I003', kind: '予定', date: '2027-01-25', cashYm: '', amount: 100000 }]);
  // 1/25 は 2/20 締め → 3月末入金
  assert.strictEqual(at(R2, 'I003', '2027-03') - at(R, 'I003', '2027-03'), 100000);
});

test('人件費：エンジニア_2は入社月から手取り、社保は翌月から', function () {
  assert.strictEqual(at(R, 'I022', '2027-03'), 0);
  // 650,000 − 本人社保97,500 − 源泉(552,500×5%=27,625→27625) − 住民税0
  assert.strictEqual(at(R, 'I022', '2027-04'), 650000 - 97500 - 27625);
});

test('人件費：賞与月（6月）は賞与込みで計算', function () {
  var g = 1200000, ee = 180000, wh = Math.round((g - ee) * 0.05);
  assert.strictEqual(at(R, 'I021', '2027-06'), g - ee - wh - 30000);
});

test('社保・源泉・住民税は翌月納付（特例なし）', function () {
  // 2026-11分：代表(150k+150k)＋エンジニア_1(90k+90k) → 2026-12に480,000
  assert.strictEqual(at(R, 'SYS:社会保険料', '2026-12'), 480000);
  assert.strictEqual(at(R, 'SYS:住民税', '2026-12'), 30000);
});

test('源泉の納期特例：納付は1月と7月だけ', function () {
  var R2 = run({ '源泉納期特例': true });
  var r = row(R2, 'SYS:源泉所得税');
  R2.months.forEach(function (m, i) {
    if (i === 0) return; // 先月は実績
    var mm = m.slice(5);
    if (mm !== '01' && mm !== '07') assert.strictEqual(r.values[i], 0, m);
  });
  assert.ok(at(R2, 'SYS:源泉所得税', '2027-01') > 0);
  assert.ok(at(R2, 'SYS:源泉所得税', '2027-07') > 0);
});

test('今月に実績があるとルール値は置き換わる（家賃が二重計上されない）', function () {
  assert.strictEqual(at(R, 'I040', '2026-11'), 450000);
  assert.strictEqual(src_(R, 'I040', '2026-11'), '実績');
});

test('上書き：今月以降は上書きが優先、先月の上書きは無視', function () {
  var R2 = run(null, [
    { id: 'O1', itemId: 'I040', kind: '上書き', cashYm: '2027-01', amount: 999 },
    { id: 'O2', itemId: 'I050', kind: '上書き', cashYm: '2026-10', amount: 1 }
  ]);
  assert.strictEqual(at(R2, 'I040', '2027-01'), 999);
  assert.strictEqual(src_(R2, 'I040', '2027-01'), '上書き');
  assert.strictEqual(at(R2, 'I050', '2026-10'), 120000);
});

test('比率：決済手数料＝売上タグの入金×3.6%', function () {
  var m = '2027-02';
  var sales = at(R, 'I001', m) + at(R, 'I002', m) + at(R, 'I003', m);
  assert.strictEqual(at(R, 'I060', m), Math.round(sales * 0.036));
});

test('仮受消費税＝課税入金×10%', function () {
  var m = '2027-02';
  var sales = at(R, 'I001', m) + at(R, 'I002', m) + at(R, 'I003', m);
  assert.strictEqual(at(R, 'SYS:仮受消費税', m), Math.round(sales * 0.1));
});

test('消費税：前期年税額240万円→中間1回、2026-11に120万円', function () {
  assert.strictEqual(R.fiscal[0].interimCount, 1);
  assert.strictEqual(at(R, 'SYS:消費税納付', '2026-11'), 1200000);
});

test('消費税：当期（〜2027-03）の確定申告は2027-05、年税額−中間', function () {
  var f = R.fiscal[0];
  assert.strictEqual(f.fyEnd, '2027-03');
  assert.ok(f.complete);
  assert.strictEqual(at(R, 'SYS:消費税納付', '2027-05'), Math.floor((f.annualTax - 1200000) / 100) * 100);
});

test('消費税：翌期の中間納付は当期の年税額が基準', function () {
  var f0 = R.fiscal[0], f1 = R.fiscal[1];
  var expectN = f0.annualTax * 0.78 > 4000000 ? 3 : f0.annualTax * 0.78 > 480000 ? 1 : 0;
  assert.strictEqual(f1.interimCount, expectN);
});

test('中間納付回数のパラメータ指定（3回）', function () {
  var R2 = run({ '中間納付回数': 3 });
  // 期首2026-04 → 2026-08, 2026-11, 2027-02 が納付月。表示期間内は 2026-11 と 2027-02
  assert.strictEqual(at(R2, 'SYS:消費税納付', '2026-11'), 600000);
  assert.strictEqual(at(R2, 'SYS:消費税納付', '2027-02'), 600000);
});

test('残高の通算：月末＝月初＋入金−出金、翌月初＝前月末（月初残高＝口座項目の合計）', function () {
  var T = R.totals;
  for (var i = 0; i < 19; i++) {
    assert.strictEqual(T.closing[i], T.opening[i] + T.inflow[i] - T.outflow[i]);
    assert.strictEqual(T.fc[i], T.closing[i] + T.cashEq[i]);
    if (i) assert.strictEqual(T.opening[i], T.closing[i - 1]);
  }
  assert.strictEqual(T.opening[0], 60000000);
});

test('burn rate＝今月以降の出金の加重平均（1, 0.9, 0.8…）、ランウェイ＝正味FC÷burn', function () {
  var T = R.totals;
  assert.strictEqual(T.burn[0], null);
  assert.strictEqual(T.burn[1], T.outflow[1]);
  var b3 = (T.outflow[1] * 0.8 + T.outflow[2] * 0.9 + T.outflow[3]) / 2.7;
  assert.ok(Math.abs(T.burn[3] - b3) < 1e-6);
  var b15 = 0, w = 0;
  for (var q = 6; q <= 15; q++) { var ww = 1 - 0.1 * (15 - q); b15 += T.outflow[q] * ww; w += ww; }
  assert.ok(Math.abs(T.burn[15] - b15 / w) < 1e-6);
  assert.ok(Math.abs(T.runway[5] - T.fc[5] / T.burn[5]) < 1e-9);
});

test('出資除く／出資・融資除く：調達額の累計を差し引く', function () {
  var T = R.totals, j8 = R.months.indexOf('2027-06'), j4 = R.months.indexOf('2027-03');
  assert.strictEqual(T.equityIn[j8], 200000000);
  assert.strictEqual(T.debtIn[j4], 30000000);
  assert.strictEqual(T.fc[18] - T.fcExEquity[18], 200000000);
  assert.strictEqual(T.fc[18] - T.fcExEquityDebt[18], 230000000);
  assert.strictEqual(T.inflowExEquity[j8], T.inflow[j8] - 200000000);
});

test('計上✗の項目：値は表示されるが合計・消費税には含まれない', function () {
  var r = row(R, 'I013');
  assert.strictEqual(r.included, false);
  assert.strictEqual(at(R, 'I013', '2027-03'), 400000); // 2027-02発生→翌月末入金
  var s = 0, tax = 0;
  R.rows.forEach(function (x) { if (x.section === '入金' && x.included) s += x.values[5]; });
  assert.strictEqual(s, R.totals.inflow[5]);
  var R2 = run(null, null, function (smp) { smp['項目'].forEach(function (x) { if (x[0] === 'I013') x[14] = true; }); });
  assert.strictEqual(R2.totals.inflow[5] - R.totals.inflow[5], 400000 + 40000 + 0); // 本体＋仮受消費税
  assert.ok(at(R2, 'I060', '2027-03') > at(R, 'I060', '2027-03'));       // 決済手数料（比率）にも反映
});

test('入金合計＝入金区分の計上✔の行の合計', function () {
  var s = 0;
  R.rows.forEach(function (r) { if (r.section === '入金' && r.included) s += r.values[7]; });
  assert.strictEqual(s, R.totals.inflow[7]);
});

test('科目一覧は区分→並び順で、項目のない科目も含む', function () {
  var secs = R.accounts.map(function (a) { return a.section; });
  assert.strictEqual(secs[0], '月初残高');
  assert.ok(R.accounts.some(function (a) { return a.name === '新株予約権'; }));
});

var failed = 0;
tests.forEach(function (t) {
  try { t[1](); console.log('  ok  ' + t[0]); }
  catch (e) { failed++; console.log('  NG  ' + t[0] + '\n      ' + e.message); }
});

// グリッドの見た目確認用ダンプ
if (process.argv.indexOf('--dump') >= 0) {
  var pad = function (s, n) { s = String(s); while (s.length < n) s = ' ' + s; return s; };
  console.log('\n' + pad('', 34) + R.months.map(function (m) { return pad(m.slice(2), 10); }).join(''));
  R.rows.forEach(function (r) {
    console.log((r.section + ' ' + r.itemName).slice(0, 30).padEnd(32) + '  ' + r.values.map(function (v) { return pad(v ? Math.round(v / 1000) + 'k' : '-', 10); }).join(''));
  });
  ['inflow', 'outflow', 'closing', 'runway'].forEach(function (k) {
    console.log(k.padEnd(34) + R.totals[k].map(function (v) { return pad(v === null ? '黒字' : k === 'runway' ? v.toFixed(1) : Math.round(v / 1000) + 'k', 10); }).join(''));
  });
  console.log(JSON.stringify(R.fiscal, null, 1));
}

console.log('\n' + (tests.length - failed) + '/' + tests.length + ' passed');
process.exit(failed ? 1 : 0);
