/* Rent receipt maker (Landlords & Rentals toolkit): India (HRA) mode with landlord PAN and revenue stamp reminders,
   or any other country with its own currency and date format. One receipt for one or many months. */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  var CUR = [['USD', '$', 'US Dollars', 'Cents', 'en-US'], ['EUR', '€', 'Euros', 'Cents', 'en-IE'], ['GBP', '£', 'Pounds', 'Pence', 'en-GB'],
    ['CAD', 'C$', 'Canadian Dollars', 'Cents', 'en-CA'], ['AUD', 'A$', 'Australian Dollars', 'Cents', 'en-AU'], ['SGD', 'S$', 'Singapore Dollars', 'Cents', 'en-SG'],
    ['AED', 'AED', 'Dirhams', 'Fils', 'en-US'], ['SAR', 'SAR', 'Riyals', 'Halalas', 'en-US'], ['ZAR', 'R', 'Rand', 'Cents', 'en-ZA'],
    ['NGN', '₦', 'Naira', 'Kobo', 'en-NG'], ['KES', 'KSh', 'Kenyan Shillings', 'Cents', 'en-KE'], ['PKR', 'Rs', 'Pakistani Rupees', 'Paisa', 'en-US'],
    ['BDT', '৳', 'Taka', 'Poisha', 'en-US'], ['NPR', 'Rs', 'Nepalese Rupees', 'Paisa', 'en-US'], ['LKR', 'Rs', 'Sri Lankan Rupees', 'Cents', 'en-US']];
  var CUR_OPTS = CUR.map(function (c) { return c[0] + ' – ' + c[2] + ' (' + c[1] + ')'; });
  var MODES = ['India (HRA receipt)', 'Other country'];
  var DATES = ['DD-MM-YYYY', 'DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD'];
  var PAY = ['Bank transfer', 'UPI', 'Cash', 'Cheque', 'Card', 'Other'];
  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  function w2(n) { return n < 20 ? ONES[n] : TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : ''); }
  function w3(n) { return [Math.floor(n / 100) ? ONES[Math.floor(n / 100)] + ' Hundred' : '', n % 100 ? w2(n % 100) : ''].filter(Boolean).join(' '); }
  function intlWords(amount, c) {   /* international style: billion, million, thousand */
    amount = r2(Math.abs(amount)); var n = Math.floor(amount), p = Math.round((amount - n) * 100), out = [];
    if (!n && !p) return '';
    [[1e9, 'Billion'], [1e6, 'Million'], [1e3, 'Thousand'], [1, '']].forEach(function (g) {
      var k = Math.floor(n / g[0]) % 1000; if (g[0] === 1e9) k = Math.floor(n / 1e9);
      if (k) out.push(w3(k) + (g[1] ? ' ' + g[1] : ''));
    });
    return (out.length ? out.join(' ') : 'Zero') + ' ' + c[2] + (p ? ' and ' + w2(p) + ' ' + c[3] : '') + ' Only';
  }
  function iso2(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }
  function monthsBack(n) { var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - n); return d; }
  function fmtDate(iso, f) {
    if (!iso) return ''; var p = String(iso).split('-'); if (p.length !== 3) return iso;
    return f === 'YYYY-MM-DD' ? iso : f === 'MM/DD/YYYY' ? p[1] + '/' + p[2] + '/' + p[0] : f === 'DD/MM/YYYY' ? p[2] + '/' + p[1] + '/' + p[0] : p[2] + '-' + p[1] + '-' + p[0];
  }
  function monthLabel(ym) { if (!ym) return ''; var p = ym.split('-'); return new Date(+p[0], +p[1] - 1, 1).toLocaleString('en-GB', { month: 'long', year: 'numeric' }); }
  function safe(s) { return String(s || 'receipt').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'receipt'; }
  var F_IN = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });

  TKTools.register('rent-receipt', {
    name: 'rent receipt', no: 'no',
    profile: [['mode', 'Country mode', MODES[0], 'select', MODES], ['cur', 'Currency (other countries)', CUR_OPTS[0], 'select', CUR_OPTS],
      ['datef', 'Date format', DATES[0], 'select', DATES], ['tenant', 'Tenant name', 'Neha Arora'],
      ['landlord', 'Landlord name', 'Suresh Kumar Gupta'], ['pan', 'Landlord PAN (India) or tax ID', ''],
      ['laddr', 'Landlord address', '45, Shastri Nagar, Jaipur, Rajasthan 302016', 'wide'],
      ['prop', 'Rented property address', 'Flat 2B, Lotus Residency, Malviya Nagar, Jaipur 302017', 'wide']],
    fields: [['no', 'Receipt number', 'RR-001'], ['date', 'Receipt date', L.today(), 'date'], ['note', 'Note (optional)', '', 'wide']],
    groups: [{ key: 'months', title: 'Rent paid (one line per month)', add: 'Add month',
      cols: [['month', 'Month', '', 'month'], ['amt', 'Rent', '', 'number'], ['paid', 'Paid on', '', 'date'], ['mode', 'Paid by', PAY[0], 'select', PAY], ['ref', 'Reference (UTR / cheque no.)', '']],
      sample: [2, 1, 0].map(function (k) { var d = monthsBack(k); return { month: iso2(d), amt: '18000', paid: iso2(d) + '-01', mode: 'UPI', ref: 'UPI ' + (610422110 + k * 37) }; }) }],
    make: function (P, D, G) {
      var india = P.mode !== MODES[1], c = india ? ['INR', '₹', 'Rupees', 'Paise'] : (CUR[CUR_OPTS.indexOf(P.cur)] || CUR[0]);
      var NF = new Intl.NumberFormat(c[4] || 'en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      var fmt = function (n) { return c[1] + ' ' + NF.format(n); };
      var rows = [], total = 0, cash = 0, months = [];
      G.months.forEach(function (l) {
        if (!l.month && !num(l.amt)) return;
        var a = num(l.amt); total += a; if (l.mode === 'Cash') cash += a; if (l.month) months.push(l.month);
        rows.push([rows.length + 1, monthLabel(l.month), india ? a : fmt(a), fmtDate(l.paid, P.datef), l.mode, l.ref]);
      });
      total = r2(total); months.sort();
      var period = !months.length ? '' : months.length === 1 ? monthLabel(months[0]) : monthLabel(months[0]) + ' to ' + monthLabel(months[months.length - 1]);
      var yearly = months.length ? Math.round(total / months.length * 12) : 0;
      var amt = india ? L.money(total) : fmt(total), inWords = india ? L.words(total) : intlWords(total, c);
      var notes = ['Received with thanks from ' + (P.tenant || '[tenant]') + ' the sum of ' + amt + (inWords ? ' (' + inWords + ')' : '') +
        ' towards rent of the property at ' + (P.prop || '[property address]') + (period ? ' for ' + period : '') + '.'];
      if (india && !P.pan && (yearly > 100000 || total > 100000))
        notes.push('Landlord PAN missing: this rent works out to about ₹ ' + F_IN.format(yearly) + ' a year. Employers usually ask for the landlord\'s PAN when rent for the year is more than ₹1,00,000. Add it above, or ask your employer what they accept if the landlord has no PAN.');
      if (india && cash > 5000) notes.push('Cash rent above ₹5,000: many employers ask for a ₹1 revenue stamp on the receipt, signed across by the landlord.');
      if (D.note) notes.push(D.note);
      return { file: 'Rent-Receipt-' + safe(D.no), title: 'RENT RECEIPT', org: P.landlord || 'Landlord name',
        sub: [P.laddr, P.pan ? (india ? 'PAN: ' : 'Tax ID: ') + P.pan : ''],
        kv: [['Receipt No.', D.no], ['Date', fmtDate(D.date, P.datef)], ['Received from', P.tenant], ['Period', period],
             ['Property', P.prop], [india ? 'Landlord PAN' : 'Landlord tax ID', P.pan || (india ? 'Not given' : '')]],
        cols: [['#', 'c'], ['Month', ''], ['Rent', india ? 'r$' : 'r'], ['Paid on', ''], ['Paid by', ''], ['Reference', '']],
        rows: rows, totals: [['Total received', india ? total : fmt(total), 1]], words: inWords, notes: notes,
        sign: ['Tenant', india && cash > 5000 ? 'Landlord\'s signature (across revenue stamp)' : 'Landlord\'s signature'] };
    }
  });
})();
