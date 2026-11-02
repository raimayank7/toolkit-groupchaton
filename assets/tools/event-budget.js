/* Events & Weddings toolkit: online event budget planner (any currency, Indian or international number style, date style). */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  var CURS = ['₹', '$', '€', '£', 'AED', 'SAR', 'QAR', 'KWD', 'OMR', 'SGD', 'MYR', 'A$', 'C$', 'NZ$', 'NPR', 'PKR', 'BDT', 'LKR', 'ZAR', 'KES', 'NGN', '¥'];
  var STYLES = ['Indian (12,34,567)', 'International (1,234,567)'];
  var DATES = ['31-12-2026', '31 Dec 2026', '12/31/2026', '2026-12-31'];
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var CATS = ['Venue & hall', 'Catering & food', 'Decoration & flowers', 'Photography & video', 'Attire & jewellery', 'Makeup, hair & mehndi',
    'Music, DJ & band', 'Invitations & stationery', 'Gifts & favours', 'Rituals & officiant', 'Transport', 'Guest stay & hotel',
    'Entertainment', 'AV, stage & lights', 'Printing & signage', 'Miscellaneous & tips', 'Contingency', 'Other'];
  function fmt(P, n) {
    var f = new Intl.NumberFormat(P.style === STYLES[1] ? 'en-US' : 'en-IN', { maximumFractionDigits: Number.isInteger(r2(n)) ? 0 : 2, minimumFractionDigits: Number.isInteger(r2(n)) ? 0 : 2 });
    return (n < 0 ? '−' : '') + (P.cur || '') + (P.cur ? ' ' : '') + f.format(Math.abs(r2(n)));
  }
  function date(P, iso) {
    var p = String(iso || '').split('-'); if (p.length !== 3) return iso || '';
    if (P.dstyle === DATES[1]) return +p[2] + ' ' + MON[+p[1] - 1] + ' ' + p[0];
    if (P.dstyle === DATES[2]) return p[1] + '/' + p[2] + '/' + p[0];
    if (P.dstyle === DATES[3]) return iso;
    return p[2] + '-' + p[1] + '-' + p[0];
  }
  function safe(s) { return String(s || 'Event').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 50) || 'Event'; }
  function inDays(n) { var d = new Date(); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); }

  TKTools.register('event-budget', {
    name: 'budget', no: 'event',
    profile: [['by', 'Prepared by (optional)', ''], ['cur', 'Currency', '₹', 'select', CURS], ['style', 'Number style', STYLES[0], 'select', STYLES], ['dstyle', 'Date style', DATES[0], 'select', DATES]],
    fields: [['event', 'Event name', "Riya & Arjun's Wedding", 'wide'], ['date', 'Event date', inDays(150), 'date'], ['venue', 'Venue / city', 'Green Meadows Resort, Ludhiana'],
      ['budget', 'Total budget', '2500000', 'number'], ['guests', 'Expected guests', '350', 'number']],
    groups: [{ key: 'lines', title: 'Expenses', add: 'Add expense',
      cols: [['cat', 'Category', CATS[0], 'select', CATS], ['item', 'Item / details', '', 'wide'], ['vendor', 'Vendor', ''], ['est', 'Estimated', '', 'number'], ['act', 'Actual (when booked)', '', 'number'], ['paid', 'Advance paid', '', 'number']],
      sample: [{ cat: 'Venue & hall', item: 'Banquet lawn, 2 days', vendor: 'Green Meadows Resort', est: '450000', act: '480000', paid: '150000' },
        { cat: 'Catering & food', item: 'Dinner, 350 plates', vendor: 'Shree Caterers', est: '385000', act: '', paid: '50000' },
        { cat: 'Decoration & flowers', item: 'Mandap, stage and lights', vendor: 'Bloom Decorators', est: '220000', act: '240000', paid: '60000' },
        { cat: 'Photography & video', item: 'Photos, film and album', vendor: 'Pixel Stories', est: '180000', act: '175000', paid: '50000' },
        { cat: 'Attire & jewellery', item: 'Bride and groom outfits', vendor: '', est: '340000', act: '', paid: '' },
        { cat: 'Contingency', item: 'Kept aside (about 10%)', vendor: '', est: '200000', act: '', paid: '' }] }],
    make: function (P, D, G) {
      var rows = [], by = {}, order = [], t = { est: 0, exp: 0, paid: 0, bal: 0 };
      G.lines.forEach(function (l) {
        if (!l.item && !num(l.est) && !num(l.act)) return;
        var est = num(l.est), hasAct = String(l.act == null ? '' : l.act).trim() !== '', exp = hasAct ? num(l.act) : est, paid = num(l.paid), bal = r2(exp - paid);
        t.est += est; t.exp += exp; t.paid += paid; t.bal += bal;
        if (!by[l.cat]) { by[l.cat] = 0; order.push(l.cat); } by[l.cat] += exp;
        rows.push([rows.length + 1, l.cat, l.item, l.vendor, fmt(P, est), hasAct ? fmt(P, exp) : '–', paid ? fmt(P, paid) : '–', bal <= 0 && exp > 0 ? 'Paid' : fmt(P, bal)]);
      });
      var budget = num(D.budget), guests = num(D.guests), left = r2(budget - t.exp);
      var tot = [['Estimated total', fmt(P, t.est)], ['Expected total (actual or estimate)', fmt(P, t.exp), 1], ['Advance paid', fmt(P, t.paid)], ['Balance due', fmt(P, t.bal), 1]];
      if (budget) tot.push([left < 0 ? 'Over budget by' : 'Budget left', fmt(P, Math.abs(left)), 1]);
      if (guests) tot.push(['Expected cost per guest', fmt(P, guests ? t.exp / guests : 0)]);
      var notes = [];
      if (order.length) notes.push('By category: ' + order.map(function (c) { return c + ' ' + fmt(P, by[c]) + (budget ? ' (' + Math.round(by[c] / budget * 100) + '%)' : ''); }).join('  ·  '));
      if (budget && left < 0) notes.push('The expected cost is more than the total budget. Look at the biggest categories first, or raise the budget.');
      notes.push('Expected cost = actual cost when booked, otherwise the estimate. Balance due = expected cost − advance paid.');
      return { file: 'Event-Budget-' + safe(D.event), title: 'EVENT BUDGET', org: D.event || 'Event budget',
        sub: [[D.venue, D.date ? date(P, D.date) : ''].filter(Boolean).join('  ·  ')],
        kv: [['Event date', date(P, D.date)], ['Venue', D.venue], ['Total budget', budget ? fmt(P, budget) : ''], ['Guests', guests ? String(guests) : ''], ['Prepared by', P.by], ['Currency', P.cur]],
        cols: [['#', 'c'], ['Category', ''], ['Item', ''], ['Vendor', ''], ['Estimated', 'r'], ['Actual', 'r'], ['Paid', 'r'], ['Balance', 'r']],
        rows: rows, totals: tot, notes: notes, sign: ['Prepared by' + (P.by ? ': ' + P.by : ''), 'Approved by'] };
    }
  });
})();
