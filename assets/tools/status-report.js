/* Weekly project status report maker (Project Management toolkit): overall RAG status, status by area, work done, milestones,
   risks and issues, next steps and an optional budget snapshot, in one table so it prints on a page. Any currency and date format.
   Uses the shared engine in assets/tools.js. */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  var RAG = ['Green', 'Amber', 'Red'], DOT = { Green: '🟢 Green', Amber: '🟡 Amber', Red: '🔴 Red' };
  var DATES = ['DD-MM-YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD', 'D Mon YYYY'];
  var CURS = ['USD', 'EUR', 'GBP', 'INR', 'AED', 'SAR', 'AUD', 'CAD', 'SGD', 'NZD', 'ZAR', 'NGN', 'KES', 'PKR', 'BDT', 'LKR', 'NPR', 'MYR', 'PHP', 'IDR', 'JPY', 'CNY', 'CHF'];
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MS = ['On track', 'At risk', 'Late', 'Done'];

  function iso(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function weekDay(n) { var d = new Date(); d.setDate(d.getDate() - ((d.getDay() + 6) % 7) + n); return iso(d); }   /* n = 0 Monday of this week */
  function fmtDate(s, f) {
    var p = String(s || '').split('-'); if (p.length !== 3) return s || '';
    if (f === 'MM/DD/YYYY') return p[1] + '/' + p[2] + '/' + p[0];
    if (f === 'YYYY-MM-DD') return s;
    if (f === 'D Mon YYYY') return (+p[2]) + ' ' + MON[+p[1] - 1] + ' ' + p[0];
    return p[2] + '-' + p[1] + '-' + p[0];
  }
  function money(n, cur) {
    try { return new Intl.NumberFormat('en', { style: 'currency', currency: cur, maximumFractionDigits: cur === 'JPY' ? 0 : 2 }).format(n); }
    catch (e) { return cur + ' ' + r2(n).toLocaleString('en'); }
  }

  TKTools.register('status-report', {
    name: 'status report', no: 'no',
    profile: [['project', 'Project name', 'Company website redesign', 'wide'], ['pm', 'Project manager', 'Priya Nair'], ['sponsor', 'Sponsor', 'Rajesh Kumar'],
      ['org', 'Organisation (optional)', ''], ['datef', 'Date format', 'DD-MM-YYYY', 'select', DATES], ['cur', 'Currency', 'USD', 'select', CURS]],
    fields: [['no', 'Report number', 'SR-04'], ['from', 'Period from', weekDay(0), 'date'], ['to', 'Period to', weekDay(4), 'date'],
      ['golive', 'Target finish (forecast)', '', 'date'],
      ['overall', 'Overall status', 'Amber', 'select', RAG], ['summary', 'Summary (progress, main problem, what you need)', 'Design is nearly done and the build has started. Design approval is two days late because the sponsor is travelling.', 'wide'],
      ['decide', 'Decisions or help needed (optional)', 'Name a deputy who can approve the design by Wednesday.', 'wide'],
      ['sched', 'Schedule', 'Amber', 'select', RAG], ['schedC', 'Schedule comment', 'Design approval 2 days late; build can start in parallel.', 'wide'],
      ['cost', 'Cost', 'Green', 'select', RAG], ['costC', 'Cost comment', 'Within budget; small overspend on design covered by contingency.', 'wide'],
      ['scope', 'Scope', 'Green', 'select', RAG], ['scopeC', 'Scope comment', 'One change request (Hindi pages) waiting for a decision.', 'wide'],
      ['qual', 'Quality', 'Green', 'select', RAG], ['qualC', 'Quality comment', '', 'wide'],
      ['budget', 'Budget (optional)', '19250', 'number'], ['spent', 'Spent to date', '6110', 'number'], ['fcast', 'Forecast at completion', '19370', 'number']],
    groups: [
      { key: 'done', title: 'Done this period', add: 'Add item', cols: [['item', 'What was done', '', 'wide']],
        sample: [{ item: 'Home page and key page mockups delivered' }, { item: 'Staging site and hosting set up' }, { item: 'Content received from 3 of 5 departments' }] },
      { key: 'ms', title: 'Milestones', add: 'Add milestone', cols: [['name', 'Milestone', '', 'wide'], ['plan', 'Planned date', '', 'date'], ['fc', 'Forecast date', '', 'date'], ['st', 'Status', 'On track', 'select', MS]],
        sample: [{ name: 'Design approved by sponsor', plan: weekDay(1), fc: weekDay(3), st: 'At risk' }, { name: 'Website live', plan: weekDay(35), fc: weekDay(35), st: 'On track' }] },
      { key: 'risks', title: 'Top risks and issues', add: 'Add risk or issue', cols: [['kind', 'Type', 'Risk', 'select', ['Risk', 'Issue']], ['text', 'Description', '', 'wide'], ['imp', 'Impact', 'Medium', 'select', ['High', 'Medium', 'Low']], ['owner', 'Owner', ''], ['act', 'Action', '', 'wide']],
        sample: [{ kind: 'Issue', text: 'Finance department has not sent its page content', imp: 'Medium', owner: 'Sana', act: 'Escalate at Monday meeting' },
                 { kind: 'Risk', text: 'Only one developer knows the hosting set-up', imp: 'High', owner: 'Vikram', act: 'Write a set-up note; pair a second developer' }] },
      { key: 'next', title: 'Next steps', add: 'Add next step', cols: [['act', 'Action', '', 'wide'], ['owner', 'Owner', ''], ['due', 'Due', '', 'date']],
        sample: [{ act: 'Design sign-off meeting', owner: 'Priya', due: weekDay(9) }, { act: 'Start building page templates', owner: 'Vikram', due: weekDay(8) }] }],
    make: function (P, D, G) {
      var f = P.datef || 'DD-MM-YYYY', dt = function (s) { return fmtDate(s, f); }, rows = [];
      [['Schedule', D.sched, D.schedC], ['Cost', D.cost, D.costC], ['Scope', D.scope, D.scopeC], ['Quality', D.qual, D.qualC]].forEach(function (a) {
        if (a[1]) rows.push(['Status', a[0] + (a[2] ? ': ' + a[2] : ''), '', DOT[a[1]] || a[1]]);
      });
      G.done.forEach(function (l) { if (l.item) rows.push(['Done', l.item, '', '✓ Done']); });
      G.ms.forEach(function (l) {
        if (!l.name) return;
        var when = [l.plan ? 'Plan ' + dt(l.plan) : '', l.fc && l.fc !== l.plan ? 'Forecast ' + dt(l.fc) : ''].filter(Boolean).join(' · ');
        rows.push(['Milestone', l.name, when, l.st]);
      });
      G.risks.forEach(function (l) { if (l.text) rows.push([l.kind || 'Risk', l.text + (l.act ? ' → ' + l.act : ''), l.owner || '', (l.imp || '') + ' impact']); });
      G.next.forEach(function (l) { if (l.act) rows.push(['Next step', l.act, [l.owner, l.due ? 'by ' + dt(l.due) : ''].filter(Boolean).join(' · '), '']); });
      var tot = [], b = num(D.budget);
      if (b) {
        var fc = num(D.fcast) || b, v = r2(b - fc);
        tot.push(['Budget', money(b, P.cur)], ['Spent to date', money(num(D.spent), P.cur) + ' (' + Math.round(num(D.spent) / b * 100) + '%)'], ['Forecast at completion', money(fc, P.cur)],
                 [v < 0 ? 'Overspend forecast' : 'Variance (saving)', money(Math.abs(v), P.cur) + ' (' + Math.round(Math.abs(v) / b * 1000) / 10 + '%)', 1]);
      }
      var notes = [];
      if (D.summary) notes.unshift('Summary: ' + D.summary);
      if (D.decide) notes.push('Decisions or help needed: ' + D.decide);
      notes.push('Key: 🟢 Green = on track. 🟡 Amber = at risk, the team has a plan to recover. 🔴 Red = off track, needs a decision or help from the sponsor.');
      return { file: 'Status-Report-' + (D.no || '').replace(/[^\w-]+/g, '') + '-' + (D.to || ''), title: 'PROJECT STATUS REPORT' + (D.no ? '  ·  ' + D.no : ''),
        org: P.project || 'Project name', sub: [P.org || ''],
        kv: [['Period', dt(D.from) + (D.to ? ' to ' + dt(D.to) : '')], ['Overall status', DOT[D.overall] || D.overall], ['Project manager', P.pm], ['Sponsor', P.sponsor],
             ['Report date', dt(L.today())], ['Target finish', D.golive ? dt(D.golive) : '']],
        cols: [['Section', ''], ['Item', ''], ['Owner / date', ''], ['Status', '']], rows: rows, totals: tot,
        notes: notes, sign: ['Project manager', 'Sponsor'] };
    }
  });
})();
