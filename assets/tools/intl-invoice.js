/* International invoice (Freelancers & Agencies toolkit): any currency, any tax label and rate, date style, payment terms,
   bank / IBAN / SWIFT details. Money is formatted here (the shared engine's money() is rupees only), so amounts reach the
   preview, Excel and Word as ready text. */
(function () {
  var L = TKTools.lib, num = L.num;
  var CUR = [['USD', 'US dollar', 'US dollars', 'cents'], ['EUR', 'Euro', 'euros', 'cents'], ['GBP', 'British pound', 'pounds sterling', 'pence'],
    ['INR', 'Indian rupee', 'Indian rupees', 'paise'], ['AED', 'UAE dirham', 'UAE dirhams', 'fils'], ['CAD', 'Canadian dollar', 'Canadian dollars', 'cents'],
    ['AUD', 'Australian dollar', 'Australian dollars', 'cents'], ['SGD', 'Singapore dollar', 'Singapore dollars', 'cents'], ['NZD', 'New Zealand dollar', 'New Zealand dollars', 'cents'],
    ['CHF', 'Swiss franc', 'Swiss francs', 'centimes'], ['JPY', 'Japanese yen', 'Japanese yen', ''], ['ZAR', 'South African rand', 'South African rand', 'cents'],
    ['SAR', 'Saudi riyal', 'Saudi riyals', 'halalas'], ['QAR', 'Qatari riyal', 'Qatari riyals', 'dirhams'], ['MYR', 'Malaysian ringgit', 'Malaysian ringgit', 'sen'],
    ['PHP', 'Philippine peso', 'Philippine pesos', 'centavos'], ['PKR', 'Pakistani rupee', 'Pakistani rupees', 'paisa'], ['BDT', 'Bangladeshi taka', 'Bangladeshi taka', 'poisha'],
    ['LKR', 'Sri Lankan rupee', 'Sri Lankan rupees', 'cents'], ['NPR', 'Nepalese rupee', 'Nepalese rupees', 'paisa'], ['NGN', 'Nigerian naira', 'Nigerian naira', 'kobo'],
    ['KES', 'Kenyan shilling', 'Kenyan shillings', 'cents'], ['EGP', 'Egyptian pound', 'Egyptian pounds', 'piastres'], ['BRL', 'Brazilian real', 'Brazilian reais', 'centavos'],
    ['MXN', 'Mexican peso', 'Mexican pesos', 'centavos'], ['HKD', 'Hong Kong dollar', 'Hong Kong dollars', 'cents'], ['SEK', 'Swedish krona', 'Swedish kronor', 'öre'],
    ['NOK', 'Norwegian krone', 'Norwegian kroner', 'øre'], ['DKK', 'Danish krone', 'Danish kroner', 'øre'], ['PLN', 'Polish złoty', 'Polish złoty', 'grosze'],
    ['IDR', 'Indonesian rupiah', 'Indonesian rupiah', ''], ['THB', 'Thai baht', 'Thai baht', 'satang']];
  var CURL = CUR.map(function (c) { return c[0] + ' – ' + c[1]; });
  var DATES = ['DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD', 'DD.MM.YYYY', 'DD MMM YYYY', 'MMM D, YYYY'];
  var TERMS = ['Due on receipt', 'Net 7', 'Net 14', 'Net 15', 'Net 30', 'Net 45', 'Net 60'];
  var TAXES = ['No tax', 'VAT', 'GST', 'Sales tax', 'HST', 'IVA', 'MwSt', 'TVA', 'Tax'];
  var UNITS = ['hours', 'days', 'items', 'words', 'pages', 'months', 'project'];
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  /* a sensible first guess of currency and date style from the visitor's time zone and language */
  var tz = '', lang = '';
  try { tz = Intl.DateTimeFormat().resolvedOptions().timeZone || ''; lang = (navigator.language || '').toUpperCase(); } catch (e) {}
  function guessCur() {
    var z = { 'Asia/Kolkata': 'INR', 'Asia/Calcutta': 'INR', 'Europe/London': 'GBP', 'Asia/Dubai': 'AED', 'Asia/Singapore': 'SGD', 'Asia/Karachi': 'PKR', 'Asia/Dhaka': 'BDT',
      'Asia/Colombo': 'LKR', 'Asia/Kathmandu': 'NPR', 'Asia/Manila': 'PHP', 'Asia/Kuala_Lumpur': 'MYR', 'Asia/Riyadh': 'SAR', 'Asia/Qatar': 'QAR', 'Asia/Tokyo': 'JPY',
      'Africa/Lagos': 'NGN', 'Africa/Nairobi': 'KES', 'Africa/Johannesburg': 'ZAR', 'Africa/Cairo': 'EGP', 'Europe/Zurich': 'CHF', 'Europe/Stockholm': 'SEK', 'Europe/Oslo': 'NOK',
      'Europe/Copenhagen': 'DKK', 'Europe/Warsaw': 'PLN', 'Asia/Hong_Kong': 'HKD', 'Asia/Jakarta': 'IDR', 'Asia/Bangkok': 'THB', 'Pacific/Auckland': 'NZD', 'America/Sao_Paulo': 'BRL', 'America/Mexico_City': 'MXN' }[tz];
    if (!z && /^Australia\//.test(tz)) z = 'AUD';
    if (!z && /^America\/(Toronto|Vancouver|Edmonton|Winnipeg|Halifax|Regina|St_Johns)/.test(tz)) z = 'CAD';
    if (!z && /^Europe\//.test(tz)) z = 'EUR';
    var i = CUR.map(function (c) { return c[0]; }).indexOf(z || 'USD');
    return CURL[i < 0 ? 0 : i];
  }
  function guessDate() { return /^(America\/(New_York|Chicago|Denver|Los_Angeles|Phoenix|Anchorage|Detroit|Indiana|Boise)|Pacific\/Honolulu)/.test(tz) || (!tz && /-US$/.test(lang)) ? 'MM/DD/YYYY' : 'DD/MM/YYYY'; }

  function code(label) { return String(label || 'USD').slice(0, 3); }
  function curInfo(c) { for (var i = 0; i < CUR.length; i++) if (CUR[i][0] === c) return CUR[i]; return CUR[0]; }
  function decimals(c) { try { return new Intl.NumberFormat('en', { style: 'currency', currency: c }).resolvedOptions().maximumFractionDigits; } catch (e) { return 2; } }
  function rnd(n, d) { var f = Math.pow(10, d); return Math.round((n + (n < 0 ? -1e-9 : 1e-9)) * f) / f; }
  function fmt(c, n) {
    try { return new Intl.NumberFormat(c === 'INR' ? 'en-IN' : 'en-US', { style: 'currency', currency: c }).format(n); }
    catch (e) { return c + ' ' + n.toFixed(2); }
  }
  function fmtDate(iso, style) {
    var p = String(iso || '').split('-'); if (p.length !== 3) return iso || '';
    var y = p[0], m = p[1], d = p[2], mn = MON[+m - 1] || m;
    return { 'MM/DD/YYYY': m + '/' + d + '/' + y, 'YYYY-MM-DD': y + '-' + m + '-' + d, 'DD.MM.YYYY': d + '.' + m + '.' + y,
      'DD MMM YYYY': d + ' ' + mn + ' ' + y, 'MMM D, YYYY': mn + ' ' + (+d) + ', ' + y }[style] || (d + '/' + m + '/' + y);
  }
  function addDays(iso, n) {
    var p = String(iso || '').split('-'); if (p.length !== 3) return '';
    var t = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
    return t.getUTCFullYear() + '-' + ('0' + (t.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + t.getUTCDate()).slice(-2);
  }
  /* amount in words, international style (thousand, million, billion); rupees keep the Indian lakh / crore style */
  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  function w2(n) { return n < 20 ? ONES[n] : TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : ''); }
  function w3(n) { var h = Math.floor(n / 100), r = n % 100, o = []; if (h) o.push(ONES[h] + ' Hundred'); if (r) o.push(w2(r)); return o.join(' '); }
  function intWords(n) {
    if (!n) return 'Zero';
    var out = [];
    [[1e12, 'Trillion'], [1e9, 'Billion'], [1e6, 'Million'], [1e3, 'Thousand']].forEach(function (s) { if (n >= s[0]) { out.push(w3(Math.floor(n / s[0])) + ' ' + s[1]); n %= s[0]; } });
    if (n) out.push(w3(n));
    return out.join(' ');
  }
  function words(c, amount) {
    if (!amount) return '';
    if (c === 'INR') return L.words(amount);
    var info = curInfo(c), dec = decimals(c), a = rnd(Math.abs(amount), dec), major = Math.floor(a), minor = dec ? Math.round((a - major) * Math.pow(10, dec)) : 0;
    return 'Amount in words: ' + intWords(major) + ' ' + cap(info[2]) + (minor && info[3] ? ' and ' + intWords(minor) + ' ' + cap(info[3]) : '') + ' Only';
  }
  function cap(s) { return s.replace(/(^|\s)(\S)/g, function (m, a, b) { return a + b.toUpperCase(); }); }
  function safe(s) { return String(s || 'invoice').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'invoice'; }

  TKTools.register('intl-invoice', {
    name: 'invoice', no: 'no',
    profile: [['name', 'Your name or business', 'Alex Morgan Studio'], ['addr', 'Address', '24 Harbour Street, Suite 3, Bristol BS1 4XY, United Kingdom', 'wide'],
      ['taxid', 'Your tax / VAT / GST number (optional)', ''], ['email', 'Email', 'hello@alexmorgan.example'], ['phone', 'Phone', ''], ['web', 'Website or portfolio', ''],
      ['dfmt', 'Date style', guessDate(), 'select', DATES],
      ['holder', 'Account holder name', 'Alex Morgan'], ['bank', 'Bank name', ''], ['acct', 'Account number or IBAN', ''], ['swift', 'SWIFT / BIC', ''],
      ['bcode', 'Bank code (sort code, routing no., IFSC, BSB)', ''], ['other', 'Other ways to pay (optional, e.g. a payment link)', '', 'wide']],
    fields: [['cur', 'Currency', guessCur(), 'select', CURL], ['no', 'Invoice number', 'INV-2026-001'], ['date', 'Invoice date', L.today(), 'date'],
      ['terms', 'Payment terms', 'Net 14', 'select', TERMS], ['cname', 'Client name', 'Northwind Labs Inc.', 'wide'], ['caddr', 'Client address', '500 Market Street, San Francisco, CA 94105, USA', 'wide'],
      ['ctax', 'Client tax / VAT number (if any)', ''], ['po', 'Client reference / PO number', ''],
      ['taxl', 'Tax label', 'No tax', 'select', TAXES], ['taxr', 'Tax rate %', '', 'number'], ['disc', 'Discount %', '', 'number'],
      ['notes', 'Notes (e.g. reverse charge statement, thank-you line)', 'Thank you for your business. Please use the invoice number as the payment reference.', 'wide']],
    groups: [{ key: 'items', title: 'Work and items', add: 'Add line',
      cols: [['desc', 'Description', '', 'wide'], ['qty', 'Quantity', '1', 'number'], ['unit', 'Unit', 'hours', 'select', UNITS], ['rate', 'Rate', '', 'number'], ['tax', 'Tax applies', true, 'check']],
      sample: [{ desc: 'UI design: onboarding screens (2 rounds of revisions)', qty: '12.5', unit: 'hours', rate: '65', tax: true },
               { desc: 'Icon set, 24 icons, SVG and PNG', qty: '1', unit: 'project', rate: '900', tax: true },
               { desc: 'Client workshop (online)', qty: '1', unit: 'days', rate: '480', tax: true }] }],
    make: function (P, D, G) {
      var c = code(D.cur), dec = decimals(c), sub = 0, taxable = 0, rows = [];
      var rate = D.taxl === 'No tax' ? 0 : num(D.taxr), dp = Math.min(100, Math.max(0, num(D.disc)));
      G.items.forEach(function (l) {
        if (!l.desc && !num(l.rate)) return;
        var a = rnd(num(l.qty) * num(l.rate), dec);
        sub += a; if (l.tax !== false && rate) taxable += a;
        var row = [rows.length + 1, l.desc, num(l.qty), l.unit, fmt(c, num(l.rate)), fmt(c, a)];
        if (rate) row.splice(5, 0, l.tax !== false ? rate + '%' : '—');
        rows.push(row);
      });
      sub = rnd(sub, dec);
      var disc = rnd(sub * dp / 100, dec), tax = rnd(taxable * (1 - dp / 100) * rate / 100, dec), total = rnd(sub - disc + tax, dec);
      var tot = [['Subtotal', fmt(c, sub)]];
      if (disc) tot.push(['Discount (' + dp + '%)', '− ' + fmt(c, disc)]);
      if (rate) tot.push([D.taxl + ' (' + rate + '%)', fmt(c, tax)]);
      tot.push(['Total due (' + c + ')', fmt(c, total), 1]);
      var n = { 'Due on receipt': 0, 'Net 7': 7, 'Net 14': 14, 'Net 15': 15, 'Net 30': 30, 'Net 45': 45, 'Net 60': 60 }[D.terms] || 0;
      var due = n ? fmtDate(addDays(D.date, n), P.dfmt) : 'On receipt';
      var kv = [['Invoice no.', D.no], ['Invoice date', fmtDate(D.date, P.dfmt)], ['Bill to', D.cname], ['Due date', due], ['Address', D.caddr], ['Payment terms', D.terms]];
      var extra = [['Client tax no.', D.ctax], ['Reference / PO', D.po]].filter(function (x) { return x[1]; });
      if (extra.length === 1) extra.push(['', '']);
      kv = kv.concat(extra);
      var bank = [['Account name', P.holder], ['Bank', P.bank], ['Account / IBAN', P.acct], ['SWIFT / BIC', P.swift], ['Bank code', P.bcode]]
        .filter(function (x) { return x[1]; }).map(function (x) { return x[0] + ': ' + x[1]; });
      var notes = [];
      if (bank.length) notes.push('Payment details: ' + bank.join('  ·  '));
      if (bank.length || P.other) notes.push('Payment reference: ' + (D.no || 'invoice number') + '. Please pay ' + fmt(c, total) + ' by ' + due.replace('On receipt', 'return') + '.');
      if (P.other) notes.push('Other ways to pay: ' + P.other);
      if (D.notes) notes.push(D.notes);
      return { file: 'Invoice-' + safe(D.no), title: 'INVOICE', org: P.name,
        sub: [P.addr, [P.taxid ? 'Tax no.: ' + P.taxid : '', P.email, P.phone, P.web].filter(Boolean).join('  ·  ')],
        kv: kv, cols: [['#', 'c'], ['Description', ''], ['Qty', 'r'], ['Unit', ''], ['Rate', 'r']].concat(rate ? [[D.taxl, 'c']] : [], [['Amount', 'r']]),
        rows: rows, totals: tot, words: words(c, total), notes: notes, sign: ['', 'For ' + (P.name || '')] };
    }
  });
})();
