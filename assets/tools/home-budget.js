/* Home & Personal Finance toolkit: online monthly household budget planner (tool id home-budget).
   Any currency and date style: amounts are formatted here with Intl (the shared money() helper is rupees only),
   so the preview, Word and Excel copies all show the chosen currency. Record-keeping aid, no advice. */
(function () {
  'use strict';
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  /* code, locale for the number style, label */
  var CUR = [['INR', 'en-IN', 'Indian rupee (₹)'], ['USD', 'en-US', 'US dollar ($)'], ['EUR', 'en-IE', 'Euro (€)'], ['GBP', 'en-GB', 'British pound (£)'],
    ['AED', 'en-AE', 'UAE dirham'], ['SAR', 'en-SA', 'Saudi riyal'], ['QAR', 'en-QA', 'Qatari riyal'], ['KWD', 'en-KW', 'Kuwaiti dinar'], ['OMR', 'en-OM', 'Omani rial'],
    ['SGD', 'en-SG', 'Singapore dollar'], ['MYR', 'en-MY', 'Malaysian ringgit'], ['AUD', 'en-AU', 'Australian dollar'], ['CAD', 'en-CA', 'Canadian dollar'],
    ['NZD', 'en-NZ', 'New Zealand dollar'], ['ZAR', 'en-ZA', 'South African rand'], ['NPR', 'en-NP', 'Nepalese rupee'], ['BDT', 'en-BD', 'Bangladeshi taka'],
    ['PKR', 'en-PK', 'Pakistani rupee'], ['LKR', 'en-LK', 'Sri Lankan rupee'], ['JPY', 'en-JP', 'Japanese yen']];
  var LABELS = CUR.map(function (c) { return c[0] + ' · ' + c[2]; });
  var DATES = ['DD-MM-YYYY', 'MM-DD-YYYY', 'YYYY-MM-DD'];
  var TYPES = ['Fixed', 'Variable', 'Savings'];

  function curOf(label) { var code = String(label || '').slice(0, 3); for (var i = 0; i < CUR.length; i++) if (CUR[i][0] === code) return CUR[i]; return CUR[0]; }
  function formatter(c) {
    var whole, part;
    try { whole = new Intl.NumberFormat(c[1], { style: 'currency', currency: c[0], maximumFractionDigits: 0, minimumFractionDigits: 0 });
          part = new Intl.NumberFormat(c[1], { style: 'currency', currency: c[0] }); }
    catch (e) { whole = part = { format: function (n) { return c[0] + ' ' + n.toFixed(2); } }; }
    return function (n) { n = r2(n); return (Number.isInteger(n) ? whole : part).format(n).replace(/^-/, '−'); };
  }
  function dateIn(iso, style) {
    var p = String(iso || '').split('-'); if (p.length !== 3) return iso || '';
    return style === 'MM-DD-YYYY' ? p[1] + '-' + p[2] + '-' + p[0] : style === 'YYYY-MM-DD' ? iso : p[2] + '-' + p[1] + '-' + p[0];
  }
  function monthIn(ym, loc) {
    var p = String(ym || '').split('-'); if (p.length < 2) return ym || '';
    try { return new Date(+p[0], +p[1] - 1, 1).toLocaleString(loc, { month: 'long', year: 'numeric' }); } catch (e) { return ym; }
  }
  function safe(s) { return String(s || 'budget').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 50) || 'budget'; }

  TKTools.register('home-budget', {
    name: 'budget', no: 'month',
    profile: [['name', 'Household name', 'The Sharma family'], ['cur', 'Currency', LABELS[0], 'select', LABELS],
      ['dfmt', 'Date style', DATES[0], 'select', DATES], ['target', 'Savings target (% of income)', '20', 'number']],
    fields: [['month', 'Month', L.today().slice(0, 7), 'month'], ['date', 'Prepared on', L.today(), 'date'], ['note', 'Note for the family (optional)', '', 'wide']],
    groups: [
      { key: 'income', title: 'Income this month', add: 'Add income',
        cols: [['src', 'Source', '', 'wide'], ['who', 'Earner', ''], ['amt', 'Amount', '', 'number']],
        sample: [{ src: 'Salary', who: 'Rohit', amt: '58000' }, { src: 'Salary', who: 'Neha', amt: '36000' }, { src: 'Tuition classes at home', who: 'Neha', amt: '4500' }] },
      { key: 'spend', title: 'Spending categories', add: 'Add category',
        cols: [['cat', 'Category', '', 'wide'], ['type', 'Type', 'Variable', 'select', TYPES], ['plan', 'Planned', '', 'number'], ['act', 'Actual', '', 'number']],
        sample: [{ cat: 'Rent / home loan EMI', type: 'Fixed', plan: '18000', act: '18000' }, { cat: 'Groceries', type: 'Variable', plan: '9000', act: '8300' },
          { cat: 'Milk, fruit and vegetables', type: 'Variable', plan: '4500', act: '4100' }, { cat: 'Electricity', type: 'Variable', plan: '2200', act: '2650' },
          { cat: 'Mobile and internet', type: 'Fixed', plan: '1500', act: '1497' }, { cat: 'School fees and tuition', type: 'Fixed', plan: '7500', act: '7500' },
          { cat: 'Fuel and transport', type: 'Variable', plan: '4000', act: '3600' }, { cat: 'House help', type: 'Fixed', plan: '11500', act: '11500' },
          { cat: 'Eating out and entertainment', type: 'Variable', plan: '3000', act: '3600' }, { cat: 'Medicines and doctor', type: 'Variable', plan: '1500', act: '1180' },
          { cat: 'Emergency fund', type: 'Savings', plan: '5000', act: '5000' }, { cat: 'Recurring deposit / SIP', type: 'Savings', plan: '15000', act: '15000' }] }],
    make: function (P, D, G) {
      var c = curOf(P.cur), m = formatter(c), rows = [], inc = 0, by = { Fixed: [0, 0], Variable: [0, 0], Savings: [0, 0] }, over = [];
      G.income.forEach(function (l) {
        if (!l.src && !num(l.amt)) return;
        inc += num(l.amt); rows.push([(l.src || 'Income') + (l.who ? ' (' + l.who + ')' : ''), 'Income', '', m(num(l.amt)), '']);
      });
      G.spend.forEach(function (l) {
        if (!l.cat && !num(l.plan) && !num(l.act)) return;
        var t = by[l.type] ? l.type : 'Variable', p = num(l.plan), a = num(l.act), d = r2(p - a);
        by[t][0] += p; by[t][1] += a;
        if (t !== 'Savings' && a > p) over.push((l.cat || 'Unnamed') + ' (' + m(a - p) + ' over)');
        rows.push([l.cat || 'Unnamed', t, m(p), m(a), d < 0 && t !== 'Savings' ? m(d) + ' over' : m(d)]);
      });
      inc = r2(inc);
      var spent = r2(by.Fixed[1] + by.Variable[1]), planned = r2(by.Fixed[0] + by.Variable[0]), saved = r2(by.Savings[1]), left = r2(inc - spent - saved);
      var target = num(P.target), rate = inc > 0 ? (inc - spent) / inc * 100 : null;
      var rateTxt = rate == null ? 'Add income to see it' : (Math.round(rate * 10) / 10) + '%';
      var status = rate == null ? '' : rate >= target ? 'On target' : 'Below target by ' + (Math.round((target - rate) * 10) / 10) + ' points';
      var month = monthIn(D.month, c[1]);
      return { file: 'Budget-' + safe(P.name) + '-' + (D.month || ''), title: 'MONTHLY HOUSEHOLD BUDGET' + (month ? '  ·  ' + month.toUpperCase() : ''),
        org: P.name || 'Household budget', sub: ['Amounts in ' + (c[2].indexOf('(') >= 0 ? c[2] : c[2] + ' (' + c[0] + ')')],
        kv: [['Household', P.name], ['Month', month], ['Prepared on', dateIn(D.date, P.dfmt)], ['Savings target', target + '%'], ['Total income', m(inc)], ['Status', status]],
        cols: [['Item', ''], ['Type', ''], ['Planned', 'r'], ['Actual', 'r'], ['Planned − actual', 'r']], rows: rows,
        totals: [['Total income', m(inc)], ['Fixed spending', m(by.Fixed[1])], ['Variable spending', m(by.Variable[1])],
          ['Spending: actual vs planned', m(spent) + ' of ' + m(planned)], ['Moved to savings', m(saved)], ['Left over', m(left)],
          ['Savings rate (income not spent)', rateTxt + ' · target ' + target + '%', 1]],
        words: '',
        notes: [over.length ? 'Over plan: ' + over.join(', ') + '.' : 'No spending category is over its plan.', D.note || '',
          'Savings rate = (income − fixed − variable spending) ÷ income. A planning and record-keeping aid, not financial advice.'].filter(Boolean),
        sign: [''] };
    }
  });
})();
