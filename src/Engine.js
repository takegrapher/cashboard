/**
 * Engine.js — フリーキャッシュ・シミュレーターの計算エンジン
 *
 * Google Apps Script（V8）と Node.js の両方で動く純粋な計算ロジック。
 * スプレッドシートAPIには一切依存しない（入出力はプレーンなオブジェクト）。
 *
 * 金額の考え方
 *   - 明細・項目の金額はすべて「税抜」。消費税は「仮受消費税」「仮払消費税」行として別に出す。
 *   - 各セルの値は「入出金月（キャッシュが動く月）」ベース。消費税の年税額は「発生月」ベースで集計する。
 *
 * セル値の優先順位（項目キー × 入出金月）
 *   先月（基準月より前）: 実績明細のみ
 *   今月以降:            上書き ＞ 実績＋予定明細 ＞ ルール＋予定明細
 *                        （その月に実績が1件でもあれば、ルール値はその実績で置き換わる）
 */
var CFEngine = (function () {
  'use strict';

  var WINDOW = 19;    // 先月 + 今月 + 17ヶ月
  var LOOKBACK = 14;  // 表示期間より前の発生分（サイト・納期特例でずれ込む分）を生成する月数
  var EXTEND = 14;    // 表示期間の後ろへ入出金を追う月数（消費税の発生ベース集計用）

  var SYS = {
    SOCIAL: 'SYS:社会保険料',
    WITHHOLD: 'SYS:源泉所得税',
    RESIDENT: 'SYS:住民税',
    TAX_IN: 'SYS:仮受消費税',
    TAX_OUT: 'SYS:仮払消費税',
    TAX_PAY: 'SYS:消費税納付'
  };
  var SYS_NAMES = {};
  SYS_NAMES[SYS.SOCIAL] = '社会保険料（納付）';
  SYS_NAMES[SYS.WITHHOLD] = '源泉所得税（納付）';
  SYS_NAMES[SYS.RESIDENT] = '住民税・特別徴収（納付）';
  SYS_NAMES[SYS.TAX_IN] = '仮受消費税（入金に含まれる消費税）';
  SYS_NAMES[SYS.TAX_OUT] = '仮払消費税（出金に含まれる消費税）';
  SYS_NAMES[SYS.TAX_PAY] = '消費税納付（中間・確定）';

  // ---------------------------------------------------------------- 汎用ユーティリティ
  function pad(n) { return n < 10 ? '0' + n : String(n); }
  function str(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function isDate(v) { return Object.prototype.toString.call(v) === '[object Date]'; }

  function num(v, d) {
    if (v === null || v === undefined || v === '') return d;
    if (typeof v === 'number') return isFinite(v) ? v : d;
    if (typeof v === 'boolean') return d;
    var s = String(v).replace(/[,，\s円¥￥]/g, '');
    var pct = /[%％]$/.test(s);
    if (pct) s = s.slice(0, -1);
    if (s === '') return d;
    var n = Number(s);
    if (!isFinite(n)) return d;
    return pct ? n / 100 : n;
  }
  // 率：0.15 / "15%" / 15 のいずれも 0.15 として扱う
  function rate(v, d) {
    var n = num(v, null);
    if (n === null) return d;
    return Math.abs(n) > 1 ? n / 100 : n;
  }
  function bool(v) {
    if (v === true) return true;
    if (v === false || v === null || v === undefined) return false;
    var s = String(v).trim().toUpperCase();
    return ['TRUE', '1', 'YES', 'Y', 'あり', '有', 'する', '○', '◯', 'ON'].indexOf(s) >= 0;
  }
  function parseIntList(v) {
    if (typeof v === 'number') return [Math.round(v)];
    return str(v).split(/[,、，\s・\/]+/).map(function (x) { return parseInt(x, 10); })
      .filter(function (x) { return x >= 1 && x <= 12; });
  }
  function parseTags(v) {
    return str(v).split(/[,、，\s#＃]+/).filter(function (x) { return x !== ''; });
  }
  function parseTokens(v) {
    return str(v).split(/[,、，\s]+/).filter(function (x) { return x !== ''; });
  }
  function fmt(n) {
    var s = String(Math.round(n));
    return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  function floor100(n) { return n >= 0 ? Math.floor(n / 100) * 100 : -Math.floor(-n / 100) * 100; }

  // ---------------------------------------------------------------- 年月ユーティリティ（'YYYY-MM' 文字列）
  function ymToIndex(ym) { return Number(ym.slice(0, 4)) * 12 + Number(ym.slice(5, 7)) - 1; }
  function indexToYm(i) { var y = Math.floor(i / 12); return y + '-' + pad(i - y * 12 + 1); }
  function addMonths(ym, n) { return indexToYm(ymToIndex(ym) + n); }
  function monthOf(ym) { return Number(ym.slice(5, 7)); }
  function yearOf(ym) { return Number(ym.slice(0, 4)); }
  function maxYm(a, b) { return a > b ? a : b; }
  function minYm(a, b) { return a < b ? a : b; }
  function eachMonth(from, to, fn) {
    if (!from || !to || from > to) return;
    for (var i = ymToIndex(from), e = ymToIndex(to); i <= e; i++) fn(indexToYm(i));
  }
  function normalizeYm(v) {
    if (v === null || v === undefined || v === '') return '';
    if (isDate(v)) return isNaN(v.getTime()) ? '' : v.getFullYear() + '-' + pad(v.getMonth() + 1);
    var s = String(v).trim().replace(/[年\/\.／]/g, '-').replace(/月.*$/, '');
    var m = s.match(/^(\d{4})-(\d{1,2})/);
    if (!m) return '';
    var mm = Number(m[2]);
    return mm >= 1 && mm <= 12 ? m[1] + '-' + pad(mm) : '';
  }
  function normalizeDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (isDate(v)) {
      if (isNaN(v.getTime())) return null;
      return { ym: v.getFullYear() + '-' + pad(v.getMonth() + 1), day: v.getDate() };
    }
    var m = String(v).trim().replace(/[年\/\.／]/g, '-').replace(/月/g, '-').replace(/日/g, '')
      .match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) return null;
    return { ym: m[1] + '-' + pad(Number(m[2])), day: Number(m[3]) };
  }

  // ---------------------------------------------------------------- 税区分
  function taxRateOf(code, P) {
    var s = str(code);
    if (!s || /非課税|対象外|不課税|免税/.test(s)) return 0;
    if (/軽減|8/.test(s)) return P.redRate;
    if (/課税|標準|10/.test(s)) return P.stdRate;
    return 0;
  }

  // ---------------------------------------------------------------- パラメータ
  function normalizeParams(raw, todayYm) {
    raw = raw || {};
    var g = function (k) { return raw[k]; };
    var interim = g('中間納付回数');
    var fem = Math.round(num(g('決算月'), 3));
    return {
      baseYm: normalizeYm(g('基準月')) || todayYm,
      fiscalEndMonth: fem >= 1 && fem <= 12 ? fem : 3,
      openingBalance: num(g('先月初残高'), 0),
      runwayN: Math.max(1, Math.round(num(g('ランウェイ算出月数'), 3))),
      stdRate: rate(g('標準税率'), 0.10),
      redRate: rate(g('軽減税率'), 0.08),
      socialCo: rate(g('社保会社負担率'), 0.15),
      socialEe: rate(g('社保本人負担率'), 0.15),
      withholdRate: rate(g('源泉税率'), 0.05),
      withholdSpecial: bool(g('源泉納期特例')),
      residentSpecial: bool(g('住民税納期特例')),
      prevAnnualTax: num(g('前期確定消費税年税額'), 0),
      prevFinalPayment: num(g('前期確定申告納付額'), 0),
      carryTax: num(g('当期窓前消費税累計'), 0),
      interimOverride: str(interim) === '' ? null : Math.round(num(interim, 0)),
      displayUnit: (/円/.test(str(g('表示単位'))) && !/千/.test(str(g('表示単位')))) ? 1 : 1000,
      acc: {
        social: str(g('科目_社会保険料')),
        withhold: str(g('科目_源泉税住民税')),
        taxIn: str(g('科目_消費税受取')),
        taxOut: str(g('科目_消費税支払')),
        taxPay: str(g('科目_消費税納付'))
      }
    };
  }

  var SECTION_ORDER = { '月初残高': 0, '現金同等物': 1, '入金': 2, '出金': 3 };
  function normSection(v) {
    var s = str(v);
    if (/月初|残高/.test(s)) return '月初残高';
    if (/同等/.test(s)) return '現金同等物';
    if (/出/.test(s)) return '出金';
    return '入金';
  }
  function normFinance(v) {
    var s = str(v);
    if (/出資|増資|株式|equity/i.test(s)) return '出資';
    if (/融資|借入|debt/i.test(s)) return '融資';
    return '営業';
  }

  function normCalcType(v) {
    var s = str(v);
    if (/固定/.test(s)) return '固定額';
    if (/人件/.test(s)) return '人件費';
    if (/比率|率/.test(s)) return '比率';
    return '明細';
  }
  function normCycle(v) {
    var s = str(v);
    if (/一括|単発|1回|一回/.test(s)) return '一括';
    if (/指定|年/.test(s)) return '指定月';
    return '毎月';
  }
  function normTiming(v) {
    var s = str(v);
    if (/終了/.test(s)) return '終了月';
    if (/毎月/.test(s)) return '毎月';
    return '開始月';
  }

  // 中間納付の回数（前期の国税年税額で判定。年税額は国税＋地方税なので 78/100 を国税分とみなす）
  function autoInterimCount(prevAnnual) {
    var national = prevAnnual * 78 / 100;
    if (national > 48000000) return 11;
    if (national > 4000000) return 3;
    if (national > 480000) return 1;
    return 0;
  }

  // ================================================================ メイン計算
  /**
   * @param {Object} data
   *   todayYm:  'YYYY-MM'（パラメータ「基準月」が空欄の場合に使う）
   *   params:   { キー: 値 }
   *   accounts, partners, items, payroll, links, details: 行オブジェクトの配列（Schema.js のフィールド名）
   */
  function compute(data) {
    var warnings = [];
    var P = normalizeParams(data.params, data.todayYm);
    var base = P.baseYm;
    if (!base) throw new Error('基準月を決められません（パラメータ「基準月」か todayYm が必要です）');
    var W0 = addMonths(base, -1);
    var months = [];
    for (var i = 0; i < WINDOW; i++) months.push(addMonths(W0, i));
    var WL = months[WINDOW - 1];
    var genStart = addMonths(W0, -LOOKBACK);
    var genEnd = WL;
    var cashEnd = addMonths(WL, EXTEND);
    var cashMonths = [];
    eachMonth(W0, cashEnd, function (m) { cashMonths.push(m); });

    // ------------------------------------------------ 科目
    var accounts = {};
    (data.accounts || []).forEach(function (a, idx) {
      var id = str(a.id);
      if (!id) return;
      if (accounts[id]) { warnings.push('科目IDが重複しています: ' + id); return; }
      accounts[id] = {
        id: id, section: normSection(a.section), major: str(a.major) || 'その他',
        name: str(a.name) || id, finance: normFinance(a.finance), desc: str(a.desc),
        order: num(a.order, 1000 + idx), idx: idx, defaultTax: str(a.defaultTax)
      };
    });
    function ensureAccount(id, section, name, label) {
      if (id && accounts[id]) {
        if (accounts[id].section !== section) {
          warnings.push('パラメータ「' + label + '」の科目 ' + id + ' は「' + section + '」区分である必要があります');
        }
        return id;
      }
      if (id) warnings.push('パラメータ「' + label + '」の科目ID ' + id + ' が科目シートにありません（仮の科目で表示します）');
      var vid = '__' + name;
      if (!accounts[vid]) {
        accounts[vid] = { id: vid, section: section, major: '自動計算', name: name, finance: '営業', desc: '',
          order: 90000 + Object.keys(accounts).length, idx: 99999, defaultTax: '対象外' };
      }
      return vid;
    }

    // ------------------------------------------------ 取引先
    var partners = {};
    (data.partners || []).forEach(function (p) {
      var id = str(p.id);
      if (!id) return;
      var c = str(p.closing);
      partners[id] = {
        id: id, name: str(p.name) || id,
        closing: (c === '' || /末/.test(c)) ? 31 : Math.max(1, Math.min(31, Math.round(num(c, 31)))),
        site: Math.round(num(p.site, 0))
      };
    });

    // ------------------------------------------------ 行（項目キー）レジストリ
    var rows = {}, rowList = [];
    function addRow(r) {
      rows[r.key] = r;
      rowList.push(r);
      return r;
    }
    var ruleIdx = {};   // key -> cash -> [comp]
    var detIdx = {};    // key -> cash -> [comp]
    function comp(accrual, cash, amount, taxCode, ref) {
      return { accrual: accrual, cash: cash, amount: Math.round(amount), taxCode: taxCode, ref: ref, kind: 'ルール' };
    }
    function pushRule(key, c) {
      if (c.cash < W0 || c.cash > cashEnd) return;
      var m = ruleIdx[key] || (ruleIdx[key] = {});
      (m[c.cash] || (m[c.cash] = [])).push(c);
    }

    // ------------------------------------------------ 項目
    var items = [];
    (data.items || []).forEach(function (it, idx) {
      var id = str(it.id);
      if (!id) return;
      if (rows[id] || id.indexOf('SYS:') === 0 || id.indexOf('L:') === 0) {
        warnings.push('項目IDが重複しているか、予約済みの形式です: ' + id);
        return;
      }
      var accId = str(it.accountId);
      if (!accounts[accId]) { warnings.push('項目「' + (str(it.name) || id) + '」の科目ID ' + accId + ' が科目シートにありません'); return; }
      var partnerId = str(it.partnerId);
      if (partnerId && !partners[partnerId]) warnings.push('項目「' + (str(it.name) || id) + '」の取引先ID ' + partnerId + ' が取引先シートにありません');
      var x = addRow({
        key: id, kind: 'item', accountId: accId, name: str(it.name) || id,
        calcType: normCalcType(it.calcType), amount: num(it.amount, 0),
        cycle: normCycle(it.cycle), cycleMonths: parseIntList(it.cycleMonths),
        start: normalizeYm(it.start), end: normalizeYm(it.end),
        partnerId: partners[partnerId] ? partnerId : '',
        taxCode: str(it.taxCode) || accounts[accId].defaultTax || '対象外',
        tags: parseTags(it.tags), ratioTargets: parseTokens(it.ratioTargets), ratio: rate(it.ratio, 0),
        active: (it.active === '' || it.active === null || it.active === undefined) ? true : bool(it.active),
        desc: str(it.desc), section: accounts[accId].section,
        order: idx
      });
      if (x.start && x.end && x.start > x.end) warnings.push('項目「' + x.name + '」の開始月が終了月より後です');
      if (x.cycle === '指定月' && x.calcType === '固定額' && !x.cycleMonths.length) {
        warnings.push('項目「' + x.name + '」は周期が「指定月」ですが指定月が空です');
      }
      items.push(x);
    });
    function siteOf(r) { return r && r.partnerId && partners[r.partnerId] ? partners[r.partnerId].site : 0; }

    // ------------------------------------------------ システム行（社保・源泉・住民税・消費税）
    var sysAccMap = {};
    sysAccMap[SYS.SOCIAL] = ensureAccount(P.acc.social, '出金', '法定福利費（社会保険料）', '科目_社会保険料');
    sysAccMap[SYS.WITHHOLD] = ensureAccount(P.acc.withhold, '出金', '源泉税・住民税', '科目_源泉税住民税');
    sysAccMap[SYS.RESIDENT] = sysAccMap[SYS.WITHHOLD];
    sysAccMap[SYS.TAX_IN] = ensureAccount(P.acc.taxIn, '入金', '消費税（受取）', '科目_消費税受取');
    sysAccMap[SYS.TAX_OUT] = ensureAccount(P.acc.taxOut, '出金', '消費税（支払・納付）', '科目_消費税支払');
    sysAccMap[SYS.TAX_PAY] = ensureAccount(P.acc.taxPay || P.acc.taxOut, '出金', '消費税（支払・納付）', '科目_消費税納付');
    var sysOrder = 0;
    [SYS.SOCIAL, SYS.WITHHOLD, SYS.RESIDENT, SYS.TAX_IN, SYS.TAX_OUT, SYS.TAX_PAY].forEach(function (k) {
      addRow({ key: k, kind: 'sys', accountId: sysAccMap[k], name: SYS_NAMES[k], calcType: '自動計算', taxCode: '対象外', partnerId: '', tags: [], active: true,
        section: accounts[sysAccMap[k]].section, desc: '', order: sysOrder++ });
    });

    // ------------------------------------------------ 固定額
    function genFixed(x) {
      var site = siteOf(x);
      if (x.cycle === '一括') {
        if (!x.start) { warnings.push('項目「' + x.name + '」は周期が「一括」なので開始月が必要です'); return; }
        if (x.start >= genStart && x.start <= genEnd) pushRule(x.key, comp(x.start, addMonths(x.start, site), x.amount, x.taxCode, '固定額（一括）'));
        return;
      }
      eachMonth(maxYm(x.start || genStart, genStart), minYm(x.end || genEnd, genEnd), function (a) {
        if (x.cycle === '指定月' && x.cycleMonths.indexOf(monthOf(a)) < 0) return;
        pushRule(x.key, comp(a, addMonths(a, site), x.amount, x.taxCode,
          x.cycle === '指定月' ? '固定額（指定月）' : '固定額（毎月）' + (site ? ' 発生' + a + '→サイト' + site + 'ヶ月' : '')));
      });
    }

    // ------------------------------------------------ 人件費
    var payroll = {};
    (data.payroll || []).forEach(function (p) {
      var id = str(p.itemId);
      if (!id) return;
      if (!rows[id]) { warnings.push('人件費シートの項目ID ' + id + ' が項目シートにありません'); return; }
      payroll[id] = {
        gross: num(p.gross, null),
        lag: /翌/.test(str(p.payTiming)) ? 1 : 0,
        bonusMonths: parseIntList(p.bonusMonths),
        bonus: num(p.bonus, 0),
        co: rate(p.co, null), ee: rate(p.ee, null), wh: rate(p.wh, null),
        resident: num(p.resident, 0)
      };
    });
    function withholdDue(pay) {
      if (!P.withholdSpecial) return addMonths(pay, 1);
      var y = yearOf(pay), m = monthOf(pay);
      return m <= 6 ? y + '-07' : (y + 1) + '-01';
    }
    function residentDue(pay) {
      if (!P.residentSpecial) return addMonths(pay, 1);
      var y = yearOf(pay), m = monthOf(pay);
      if (m >= 6 && m <= 11) return y + '-12';
      return m === 12 ? (y + 1) + '-06' : y + '-06';
    }
    function genPayroll(x) {
      var pr = payroll[x.key] || { gross: null, lag: 0, bonusMonths: [], bonus: 0, co: null, ee: null, wh: null, resident: 0 };
      var gross = pr.gross !== null ? pr.gross : x.amount;
      if (!gross) warnings.push('人件費「' + x.name + '」の月給が未設定です（人件費シートか項目シートの金額）');
      var coR = pr.co !== null ? pr.co : P.socialCo;
      var eeR = pr.ee !== null ? pr.ee : P.socialEe;
      var whR = pr.wh !== null ? pr.wh : P.withholdRate;
      eachMonth(maxYm(x.start || genStart, genStart), minYm(x.end || genEnd, genEnd), function (a) {
        var bonus = (pr.bonus && pr.bonusMonths.indexOf(monthOf(a)) >= 0) ? pr.bonus : 0;
        var g = gross + bonus;
        var ee = Math.round(g * eeR), co = Math.round(g * coR);
        var wh = Math.round((g - ee) * whR);
        var res = pr.resident;
        var pay = addMonths(a, pr.lag);
        pushRule(x.key, comp(a, pay, g - ee - wh - res, '対象外',
          '手取り＝額面' + fmt(g) + (bonus ? '（うち賞与' + fmt(bonus) + '）' : '') + '−社保本人' + fmt(ee) + '−源泉' + fmt(wh) + '−住民税' + fmt(res)));
        if (!x.active) return; // 計上✗の人件費は手取り行だけ表示し、社保・源泉・住民税には含めない
        pushRule(SYS.SOCIAL, comp(a, addMonths(a, 1), ee + co, '対象外', x.name + ' ' + a + '分（会社' + fmt(co) + '＋本人' + fmt(ee) + '）'));
        if (wh) pushRule(SYS.WITHHOLD, comp(pay, withholdDue(pay), wh, '対象外', x.name + ' ' + pay + '支給分'));
        if (res) pushRule(SYS.RESIDENT, comp(pay, residentDue(pay), res, '対象外', x.name + ' ' + pay + '支給分'));
      });
    }

    // 計上✗の項目も値は計算して表示する（合計・消費税・連動・比率の対象からは外す）
    items.forEach(function (x) {
      if (x.section === '月初残高') return;
      if (x.calcType === '固定額') genFixed(x);
      else if (x.calcType === '人件費') genPayroll(x);
    });

    // ------------------------------------------------ 連動ルール（項目の追加・削除に連動して項目を自動生成）
    (data.links || []).forEach(function (L, li) {
      var id = str(L.id);
      if (!id) return;
      if (!(L.active === '' || L.active === null || L.active === undefined || bool(L.active))) return;
      var condAcc = str(L.condAccountId), condTag = parseTags(L.condTag)[0] || '';
      var genAcc = str(L.genAccountId);
      if (!condAcc && !condTag) { warnings.push('連動ルール ' + id + ' は条件科目IDか条件タグのどちらかが必要です'); return; }
      if (!accounts[genAcc]) { warnings.push('連動ルール ' + id + ' の生成科目ID ' + genAcc + ' が科目シートにありません'); return; }
      var timing = normTiming(L.timing), offset = Math.round(num(L.offset, 0)), amount = num(L.amount, 0);
      var taxCode = str(L.taxCode) || accounts[genAcc].defaultTax || '対象外';
      var tpl = str(L.nameTpl) || '{項目名} 連動';
      items.forEach(function (s) {
        if (!s.active || s.section === '月初残高') return;
        if (condAcc && s.accountId !== condAcc) return;
        if (condTag && s.tags.indexOf(condTag) < 0) return;
        var key = 'L:' + id + ':' + s.key;
        addRow({ key: key, kind: 'gen', accountId: genAcc, name: tpl.replace(/\{項目名\}/g, s.name), calcType: '連動',
          taxCode: taxCode, partnerId: '', tags: [], active: true, section: accounts[genAcc].section,
          desc: '連動ルール ' + id + '（' + s.name + '）', order: 100000 + li * 10000 + s.order, sourceKey: s.key, linkId: id });
        var ref = '連動ルール ' + id + '（' + s.name + '・' + timing + (offset ? ' ずらし' + offset + 'ヶ月' : '') + '）';
        var at = function (a) { if (a >= genStart && a <= genEnd) pushRule(key, comp(a, a, amount, taxCode, ref)); };
        if (timing === '開始月') { if (s.start) at(addMonths(s.start, offset)); }
        else if (timing === '終了月') { if (s.end) at(addMonths(s.end, offset)); }
        else {
          var from = s.start ? addMonths(s.start, offset) : genStart;
          var to = s.end ? addMonths(s.end, offset) : genEnd;
          eachMonth(maxYm(from, genStart), minYm(to, genEnd), at);
        }
      });
    });

    // ------------------------------------------------ 明細（実績・予定・上書き）
    (data.details || []).forEach(function (d, i) {
      var key = str(d.itemId);
      if (!key) return;
      var r = rows[key];
      if (!r) { warnings.push('明細 ' + (str(d.id) || ('行' + (i + 2))) + ' の項目ID ' + key + ' が見つかりません'); return; }
      var amount = num(d.amount, null);
      if (amount === null) { warnings.push('明細 ' + (str(d.id) || ('行' + (i + 2))) + ' の金額が数値ではありません'); return; }
      var dt = normalizeDate(d.date), cashYm = normalizeYm(d.cashYm);
      var partner = r.partnerId ? partners[r.partnerId] : null;
      var site = partner ? partner.site : 0;
      var cash, accrual;
      if (cashYm) { cash = cashYm; accrual = dt ? dt.ym : addMonths(cashYm, -site); }
      else if (dt) { accrual = dt.ym; cash = addMonths(dt.ym, partner ? partner.site + (dt.day > partner.closing ? 1 : 0) : 0); }
      else { warnings.push('明細 ' + (str(d.id) || ('行' + (i + 2))) + ' は発生日か入出金月が必要です'); return; }
      var k = str(d.kind);
      var kind = /上書/.test(k) ? '上書き' : /予定|見込/.test(k) ? '予定' : /実績/.test(k) ? '実績' : (cash < base ? '実績' : '予定');
      var taxCode = r.kind === 'sys' ? '対象外' : (str(d.taxCode) || r.taxCode || '対象外');
      var m = detIdx[key] || (detIdx[key] = {});
      (m[cash] || (m[cash] = [])).push({
        accrual: accrual, cash: cash, amount: amount, taxCode: taxCode, kind: kind,
        ref: kind === '上書き' ? 'シミュレーター上の上書き' : '明細', detailId: str(d.id), memo: str(d.memo)
      });
    });

    // ------------------------------------------------ セル値の決定
    var final = {};
    function resolve(key, c, derived) {
      var d = (detIdx[key] && detIdx[key][c]) || [];
      var act = [], plan = [], over = [];
      d.forEach(function (x) { (x.kind === '実績' ? act : x.kind === '予定' ? plan : over).push(x); });
      var rc = (ruleIdx[key] && ruleIdx[key][c]) || [];
      if (c < base) {
        if (act.length || !derived) return { comps: act, source: act.length ? '実績' : '' };
        return { comps: rc, source: rc.length ? '自動計算' : '' };
      }
      if (over.length) return { comps: [over[over.length - 1]], source: '上書き' };
      if (act.length) return { comps: act.concat(plan), source: plan.length ? '実績+予定' : '実績' };
      return { comps: rc.concat(plan), source: rc.length ? (derived ? '自動計算' : (plan.length ? 'ルール+予定' : 'ルール')) : (plan.length ? '予定' : '') };
    }
    function resolveKeys(keys, derived) {
      keys.forEach(function (k) {
        final[k] = {};
        cashMonths.forEach(function (c) { final[k][c] = resolve(k, c, derived); });
      });
    }
    function cellSum(k, c) {
      var f = final[k] && final[k][c];
      if (!f) return 0;
      var s = 0;
      for (var j = 0; j < f.comps.length; j++) s += f.comps[j].amount;
      return s;
    }

    // 第1段：比率以外の項目＋連動生成行
    var phase1 = rowList.filter(function (r) {
      return ((r.kind === 'item' && r.calcType !== '比率') || r.kind === 'gen') && r.section !== '月初残高';
    }).map(function (r) { return r.key; });
    resolveKeys(phase1, false);

    // 第2段：比率（第1段の確定値 × 率）。比率の対象は計上✔の項目だけ
    var ratioItems = items.filter(function (x) { return x.calcType === '比率' && x.section !== '月初残高'; });
    ratioItems.forEach(function (x) {
      var targets = phase1.filter(function (k) {
        var r = rows[k];
        if (!r.active) return false;
        return x.ratioTargets.some(function (t) {
          if (t.charAt(0) === '#' || t.charAt(0) === '＃') return (r.tags || []).indexOf(t.slice(1)) >= 0;
          return r.accountId === t || r.key === t;
        });
      });
      if (!targets.length) { warnings.push('比率項目「' + x.name + '」の比率対象に該当する項目がありません（#タグ・科目ID・項目IDで指定）'); return; }
      var site = siteOf(x);
      cashMonths.forEach(function (c) {
        if (x.start && c < x.start) return;
        if (x.end && c > x.end) return;
        var b = 0;
        targets.forEach(function (k) { b += cellSum(k, c); });
        if (!b) return;
        pushRule(x.key, comp(c, addMonths(c, site), b * x.ratio, x.taxCode, '比率 ' + (Math.round(x.ratio * 100000) / 1000) + '% × ' + fmt(b)));
      });
    });
    resolveKeys(ratioItems.map(function (x) { return x.key; }), false);

    // 第3段：給与関連のシステム行
    resolveKeys([SYS.SOCIAL, SYS.WITHHOLD, SYS.RESIDENT], false);

    // 第4段：仮受・仮払消費税（確定したセルの構成要素ごとに税額を計算）
    var taxBaseKeys = rowList.filter(function (r) {
      if (!r.active || (r.section !== '入金' && r.section !== '出金')) return false;
      return r.kind !== 'sys' || r.key === SYS.SOCIAL || r.key === SYS.WITHHOLD || r.key === SYS.RESIDENT;
    }).map(function (r) { return r.key; });
    taxBaseKeys.forEach(function (k) {
      var r = rows[k], sec = accounts[r.accountId].section;
      cashMonths.forEach(function (c) {
        final[k][c].comps.forEach(function (x) {
          var t = taxRateOf(x.taxCode, P);
          if (!t || !x.amount) return;
          var c2 = comp(x.accrual, c, x.amount * t, '対象外', r.name + '（' + x.taxCode + '・税抜' + fmt(x.amount) + '・発生' + x.accrual + '）');
          pushRule(sec === '入金' ? SYS.TAX_IN : SYS.TAX_OUT, c2);
        });
      });
    });
    resolveKeys([SYS.TAX_IN, SYS.TAX_OUT], true);

    // 第5段：消費税の納付（中間・確定）
    function fyEndOf(ym) {
      var F = P.fiscalEndMonth;
      return (monthOf(ym) <= F ? yearOf(ym) : yearOf(ym) + 1) + '-' + pad(F);
    }
    function fyLabel(fyE) { return addMonths(fyE, -11) + '〜' + fyE + '期'; }
    var netByAccrual = {};
    [[SYS.TAX_IN, 1], [SYS.TAX_OUT, -1]].forEach(function (pair) {
      cashMonths.forEach(function (c) {
        final[pair[0]][c].comps.forEach(function (x) {
          if (x.accrual < W0) return; // 先月より前の発生分はパラメータ「当期窓前消費税累計」で扱う
          netByAccrual[x.accrual] = (netByAccrual[x.accrual] || 0) + pair[1] * x.amount;
        });
      });
    });
    var fyEnd0 = fyEndOf(W0);
    var fiscal = [];
    var prevFyEnd = addMonths(fyEnd0, -12);
    if (P.prevFinalPayment) {
      pushRule(SYS.TAX_PAY, comp(prevFyEnd, addMonths(prevFyEnd, 2), P.prevFinalPayment, '対象外',
        fyLabel(prevFyEnd) + ' 確定申告（パラメータ「前期確定申告納付額」）'));
    }
    var prevAnnual = P.prevAnnualTax, prevKnown = true;
    for (var k = 0; k < 10; k++) {
      var fyE = addMonths(fyEnd0, 12 * k), fyS = addMonths(fyE, -11);
      if (fyS > WL) break;
      var n = 0, each = 0, dues = [];
      if (prevKnown) {
        n = (k === 0 && P.interimOverride !== null) ? P.interimOverride : autoInterimCount(prevAnnual);
        if (n === 1) { dues = [7]; each = floor100(prevAnnual * 6 / 12); }
        else if (n === 3) { dues = [4, 7, 10]; each = floor100(prevAnnual * 3 / 12); }
        else if (n === 11) { dues = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]; each = floor100(prevAnnual / 12); }
        else if (n !== 0) { warnings.push('中間納付回数は 0・1・3・11 のいずれかです（' + n + ' は無視しました）'); n = 0; }
      } else {
        warnings.push(fyLabel(fyE) + ' の中間納付は前期の年税額が表示期間内で確定しないため計上していません');
      }
      var interimTotal = 0;
      dues.forEach(function (off, j) {
        interimTotal += each;
        pushRule(SYS.TAX_PAY, comp(fyE, addMonths(fyS, off), each, '対象外',
          fyLabel(fyE) + ' 中間納付 ' + (j + 1) + '/' + n + '（前期年税額' + fmt(prevAnnual) + 'が基準）'));
      });
      var N = k === 0 ? P.carryTax : 0;
      eachMonth(maxYm(fyS, W0), fyE, function (a) { N += netByAccrual[a] || 0; });
      var complete = fyE <= WL;
      var finalPay = null;
      if (complete) {
        finalPay = floor100(N - interimTotal);
        pushRule(SYS.TAX_PAY, comp(fyE, addMonths(fyE, 2), finalPay, '対象外',
          fyLabel(fyE) + ' 確定申告（年税額' + fmt(N) + '−中間' + fmt(interimTotal) + (finalPay < 0 ? '：還付' : '') + '）'));
      }
      fiscal.push({ fyStart: fyS, fyEnd: fyE, annualTax: Math.round(N), complete: complete, interimCount: n, interimEach: each, interimTotal: interimTotal, finalPayment: finalPay, finalDue: addMonths(fyE, 2) });
      prevAnnual = N;
      prevKnown = complete;
    }
    resolveKeys([SYS.TAX_PAY], false);

    // ------------------------------------------------ 出力行の組み立て
    var kindRank = { item: 0, gen: 1, sys: 2 };
    var outRows = [];
    rowList.forEach(function (r) {
      var acc = accounts[r.accountId];
      var values, sources;
      if (r.section === '月初残高') {
        // 月初残高：先月の列に「先月初の残高」（項目の金額）だけを置く
        values = months.map(function (c, j) { return j === 0 ? r.amount : 0; });
        sources = months.map(function () { return ''; });
      } else {
        if (!final[r.key]) resolveKeys([r.key], false);
        values = months.map(function (c) { return cellSum(r.key, c); });
        sources = months.map(function (c) { return final[r.key][c].source; });
      }
      var touched = values.some(function (v) { return v !== 0; }) || sources.some(function (s) { return s === '上書き' || s === '実績'; });
      if (r.kind === 'gen' && !touched) return;
      var p = r.partnerId && partners[r.partnerId];
      outRows.push({
        key: r.key, kind: r.kind, section: acc.section, accountId: acc.id, accountName: acc.name, major: acc.major,
        finance: acc.finance, itemName: r.name, desc: r.desc || '', partnerName: p ? p.name : '', tags: (r.tags || []).join(' '),
        calcType: r.calcType, included: !!r.active, values: values, sources: sources,
        _sort: [SECTION_ORDER[acc.section], acc.order, acc.idx, kindRank[r.kind], r.order]
      });
    });
    function bySort(a, b) {
      for (var j = 0; j < a._sort.length; j++) if (a._sort[j] !== b._sort[j]) return a._sort[j] < b._sort[j] ? -1 : 1;
      return 0;
    }
    outRows.sort(bySort);
    outRows.forEach(function (r) { delete r._sort; });

    // 科目一覧（項目がない科目も、グリッドに大分類・中分類の行として出す）
    var accList = Object.keys(accounts).map(function (id) {
      var a = accounts[id];
      return { id: a.id, section: a.section, major: a.major, name: a.name, finance: a.finance, desc: a.desc,
        _sort: [SECTION_ORDER[a.section], a.order, a.idx] };
    }).sort(bySort);
    accList.forEach(function (a) { delete a._sort; });

    // ------------------------------------------------ 集計
    var zeros = function () { return months.map(function () { return 0; }); };
    var inflow = zeros(), outflow = zeros(), equityIn = zeros(), debtIn = zeros(), cashEq = zeros();
    var opening0 = 0, hasOpeningItems = false;
    outRows.forEach(function (r) {
      if (r.section === '月初残高') { hasOpeningItems = true; if (r.included) opening0 += r.values[0]; return; }
      if (!r.included) return;
      r.values.forEach(function (v, j) {
        if (r.section === '入金') {
          inflow[j] += v;
          if (r.finance === '出資') equityIn[j] += v;
          if (r.finance === '融資') debtIn[j] += v;
        } else if (r.section === '出金') outflow[j] += v;
        else cashEq[j] += v;
      });
    });
    if (!hasOpeningItems) opening0 = P.openingBalance;
    var net = [], opening = [], closing = [], fc = [], fcExEquity = [], fcExEquityDebt = [];
    var inflowExEquity = [], inflowExEquityDebt = [], burn = [], runway = [];
    var cumEq = 0, cumDebt = 0;
    for (var j = 0; j < WINDOW; j++) {
      net.push(inflow[j] - outflow[j]);
      opening.push(j === 0 ? opening0 : closing[j - 1]);
      closing.push(opening[j] + net[j]);              // 現金・預金の月末残高
      fc.push(closing[j] + cashEq[j]);                 // 正味フリーキャッシュ＝現金・預金＋現金同等物
      cumEq += equityIn[j]; cumDebt += debtIn[j];
      fcExEquity.push(fc[j] - cumEq);
      fcExEquityDebt.push(fc[j] - cumEq - cumDebt);
      inflowExEquity.push(inflow[j] - equityIn[j]);
      inflowExEquityDebt.push(inflow[j] - equityIn[j] - debtIn[j]);
      // burn rate：今月以降、出金の加重平均（当月の重み1、1ヶ月遡るごとに0.1ずつ下げ、最大10ヶ月）
      if (j === 0) { burn.push(null); runway.push(null); continue; }
      var sw = 0, sv = 0;
      for (var q = Math.max(1, j - 9); q <= j; q++) { var w = 1 - 0.1 * (j - q); sw += w; sv += outflow[q] * w; }
      burn.push(sv / sw);
      runway.push(burn[j] > 0 ? fc[j] / burn[j] : null);
    }

    // CALC（展開結果）用の縦持ちデータ
    var long = [];
    outRows.forEach(function (r) {
      if (r.section === '月初残高') return;
      months.forEach(function (c) {
        final[r.key][c].comps.forEach(function (x) {
          long.push({ key: r.key, accountName: r.accountName, itemName: r.itemName, ym: c, amount: x.amount,
            source: x.kind, ref: x.ref, accrual: x.accrual, taxCode: x.taxCode, detailId: x.detailId || '', memo: x.memo || '' });
        });
      });
    });

    return {
      base: base, months: months, params: P, rows: outRows, accounts: accList,
      totals: {
        inflow: inflow, outflow: outflow, net: net, opening: opening, closing: closing, cashEq: cashEq,
        fc: fc, fcExEquity: fcExEquity, fcExEquityDebt: fcExEquityDebt, equityIn: equityIn, debtIn: debtIn,
        inflowExEquity: inflowExEquity, inflowExEquityDebt: inflowExEquityDebt, burn: burn, runway: runway
      },
      fiscal: fiscal, long: long, warnings: warnings
    };
  }

  return {
    compute: compute, SYS: SYS, WINDOW: WINDOW,
    addMonths: addMonths, normalizeYm: normalizeYm, normalizeDate: normalizeDate, num: num, str: str
  };
})();

if (typeof module !== 'undefined') module.exports = { CFEngine: CFEngine };
