/* OPD / consultation fee receipt (Clinic & Pharmacy Admin toolkit). Payment record only: no clinical content.
   Registers with the shared engine in assets/tools.js; nothing leaves the browser. */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  function safe(s) { return String(s || 'receipt').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'receipt'; }
  TKTools.register('opd-receipt', {
    name: 'OPD receipt', no: 'no',
    profile: [['name', 'Clinic name', 'Sehat Family Clinic'], ['addr', 'Address', '45, Civil Lines, Ludhiana, Punjab 141001', 'wide'],
      ['phone', 'Phone', '0161 400 0000'], ['reg', 'Registration no. (optional)', '']],
    fields: [['no', 'Receipt number', 'OPD-0001'], ['date', 'Date', L.today(), 'date'], ['patient', 'Patient name', 'Ramesh Kumar', 'wide'],
      ['age', 'Age', '54'], ['sex', 'Sex', 'Male', 'select', ['', 'Male', 'Female', 'Other']], ['pid', 'Patient ID / UHID', 'P-0001'],
      ['pphone', 'Patient phone', ''], ['doctor', 'Doctor', 'Dr. A. Sharma'],
      ['vtype', 'Visit type', 'New consultation', 'select', ['New consultation', 'Follow-up', 'Review', 'Emergency', 'Procedure']],
      ['mode', 'Payment mode', 'Cash', 'select', ['Cash', 'UPI', 'Card', 'Cheque']], ['ref', 'Payment reference (UPI / card / cheque no.)', ''],
      ['disc', 'Discount (₹)', '', 'number'], ['note', 'Note on the receipt', 'Thank you. Please bring this receipt on your next visit.', 'wide']],
    groups: [{ key: 'charges', title: 'Charges', add: 'Add charge', cols: [['desc', 'Charge', '', 'wide'], ['qty', 'Qty', '1', 'number'], ['rate', 'Rate (₹)', '', 'number']],
      sample: [{ desc: 'Consultation fee', qty: '1', rate: '500' }, { desc: 'Registration / file charges', qty: '1', rate: '100' }] }],
    make: function (P, D, G) {
      var rows = [], sub = 0;
      G.charges.forEach(function (l) {
        if (!l.desc && !num(l.rate)) return;
        var q = String(l.qty == null ? '' : l.qty).trim() === '' ? 1 : num(l.qty), a = r2(q * num(l.rate)); sub += a;
        rows.push([rows.length + 1, l.desc || 'Charge', q, num(l.rate), a]);
      });
      sub = r2(sub);
      var disc = Math.min(sub, Math.max(0, num(D.disc))), paid = r2(sub - disc);
      var tot = [['Total charges', sub]];
      if (disc) tot.push(['Discount', -disc]);
      tot.push(['Amount received', paid, 1]);
      var who = [D.age ? D.age + ' yrs' : '', D.sex].filter(Boolean).join(' / ');
      var how = 'Paid by ' + (D.mode || 'Cash') + (D.ref && D.mode !== 'Cash' ? ' (ref. ' + D.ref + ')' : '') + '.';
      return { file: 'OPD-Receipt-' + safe(D.no), title: 'OPD RECEIPT', org: P.name,
        sub: [P.addr, [P.phone ? 'Phone: ' + P.phone : '', P.reg ? 'Reg. No.: ' + P.reg : ''].filter(Boolean).join('  ·  ')],
        kv: [['Receipt No.', D.no], ['Date', L.dmy(D.date)], ['Patient', D.patient], ['Age / Sex', who], ['Patient ID', D.pid], ['Phone', D.pphone],
             ['Doctor', D.doctor], ['Visit type', D.vtype]],
        cols: [['#', 'c'], ['Particulars', ''], ['Qty', 'r'], ['Rate', 'r$'], ['Amount', 'r$']], rows: rows, totals: tot, words: L.words(paid),
        notes: [how, D.note, 'This is a payment receipt only. It is not a prescription or a medical record.'].filter(Boolean),
        sign: ['Patient / attendant', 'For ' + (P.name || 'the clinic') + '\n\nAuthorised signatory'] };
    }
  });
})();
