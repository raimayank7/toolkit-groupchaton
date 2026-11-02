/* Sales & CRM toolkit: quotation / proforma invoice / estimate in any currency, number style and date format, with optional tax
   (added, included or none), line and overall discounts, other charges, validity date and terms. Amounts are passed to the shared
   renderer as ready-formatted text, because the engine's own money() is rupees only. */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  /* code, symbol, name, major unit (words), minor unit, scale ('in' = lakh/crore, 'w' = million), decimals */
  var CUR = [['INR', '₹', 'Indian rupee', 'Rupees', 'Paise', 'in', 2], ['USD', '$', 'US dollar', 'US Dollars', 'Cents', 'w', 2],
    ['EUR', '€', 'Euro', 'Euros', 'Cents', 'w', 2], ['GBP', '£', 'Pound sterling', 'Pounds', 'Pence', 'w', 2],
    ['AED', 'AED', 'UAE dirham', 'UAE Dirhams', 'Fils', 'w', 2], ['SAR', 'SAR', 'Saudi riyal', 'Saudi Riyals', 'Halalas', 'w', 2],
    ['QAR', 'QAR', 'Qatari riyal', 'Qatari Riyals', 'Dirhams', 'w', 2], ['AUD', 'A$', 'Australian dollar', 'Australian Dollars', 'Cents', 'w', 2],
    ['CAD', 'C$', 'Canadian dollar', 'Canadian Dollars', 'Cents', 'w', 2], ['SGD', 'S$', 'Singapore dollar', 'Singapore Dollars', 'Cents', 'w', 2],
    ['NZD', 'NZ$', 'New Zealand dollar', 'New Zealand Dollars', 'Cents', 'w', 2], ['NPR', 'Rs', 'Nepalese rupee', 'Nepalese Rupees', 'Paisa', 'in', 2],
    ['PKR', 'Rs', 'Pakistani rupee', 'Pakistani Rupees', 'Paisa', 'in', 2], ['LKR', 'Rs', 'Sri Lankan rupee', 'Sri Lankan Rupees', 'Cents', 'w', 2],
    ['BDT', 'Tk', 'Bangladeshi taka', 'Taka', 'Poisha', 'in', 2], ['MYR', 'RM', 'Malaysian ringgit', 'Ringgit', 'Sen', 'w', 2],
    ['PHP', '₱', 'Philippine peso', 'Pesos', 'Centavos', 'w', 2], ['ZAR', 'R', 'South African rand', 'Rand', 'Cents', 'w', 2],
    ['NGN', '₦', 'Nigerian naira', 'Naira', 'Kobo', 'w', 2], ['KES', 'KSh', 'Kenyan shilling', 'Kenyan Shillings', 'Cents', 'w', 2],
    ['JPY', '¥', 'Japanese yen', 'Yen', '', 'w', 0]];
  var CUR_OPTS = CUR.map(function (c) { return c[0] + ' · ' + c[1] + ' · ' + c[2]; });
  var STYLES = { 'Auto (by currency)': '', '1,234,567.89': 'en-US', '12,34,567.89 (Indian)': 'en-IN', '1.234.567,89': 'de-DE', '1 234 567,89': 'fr-FR', '1\'234\'567.89': 'en-US-ch' };
  var DATES = ['DD-MM-YYYY', 'DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD', 'DD Mon YYYY'];
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MODES = ['Add tax to prices', 'Prices include tax', 'No tax'];
  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function curOf(label) { var code = String(label || '').slice(0, 3); for (var i = 0; i < CUR.length; i++) if (CUR[i][0] === code) return CUR[i]; return CUR[0]; }
  function w2(n) { return n < 20 ? ONES[n] : TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : ''); }
  function w3(n) { var h = Math.floor(n / 100), r = n % 100; return [h ? ONES[h] + ' Hundred' : '', r ? w2(r) : ''].filter(Boolean).join(' '); }
  function intl(n) {   /* million / billion scale */
    var out = [], sc = [[1e12, 'Trillion'], [1e9, 'Billion'], [1e6, 'Million'], [1e3, 'Thousand']];
    sc.forEach(function (s) { var k = Math.floor(n / s[0]); if (k) { out.push(w3(k) + ' ' + s[1]); n -= k * s[0]; } });
    if (n) out.push(w3(n));
    return out.join(' ');
  }
  function inWords(amount, c) {
    amount = r2(Math.abs(amount)); var n = Math.floor(amount), p = c[6] ? Math.round((amount - n) * 100) : 0;
    if (!n && !p) return '';
    if (c[5] === 'in') {   /* lakh / crore from the shared helper, with this currency's names */
      return L.words(amount).replace(/^Rupees /, c[3] + ' ').replace(/ Paise Only$/, ' ' + c[4] + ' Only');
    }
    return c[3] + ' ' + (n ? intl(n) : 'Zero') + (p ? ' and ' + w2(p) + ' ' + c[4] : '') + ' Only';
  }
  function fmtDate(iso, f) {
    if (!iso) return ''; var p = String(iso).split('-'); if (p.length !== 3) return iso;
    var y = p[0], m = p[1], d = p[2];
    return f === 'DD/MM/YYYY' ? d + '/' + m + '/' + y : f === 'MM/DD/YYYY' ? m + '/' + d + '/' + y : f === 'YYYY-MM-DD' ? iso :
      f === 'DD Mon YYYY' ? d + ' ' + MON[+m - 1] + ' ' + y : d + '-' + m + '-' + y;
  }
  function addDays(iso, n) {
    if (!iso) return ''; var p = iso.split('-'), dt = new Date(+p[0], +p[1] - 1, +p[2] + Math.round(n));
    return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
  }
  function safe(s) { return String(s || 'document').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'document'; }
  function qty(n) { return String(r2(n)); }
  function pct(n) { return n ? String(r2(n)) + '%' : ''; }

  TKTools.register('sales-quote', {
    name: 'quotation', no: 'no',
    profile: [['name', 'Business name', 'Apex Office Solutions'], ['addr', 'Address', '2nd Floor, Trade Centre, MG Road, Pune 411001', 'wide'],
      ['phone', 'Phone', '+91 98200 00000'], ['email', 'Email', 'sales@apex.example'], ['web', 'Website', ''],
      ['taxlabel', 'Tax ID name (GSTIN, VAT no., TRN …)', 'GSTIN'], ['taxid', 'Your tax ID', ''], ['bank', 'Payment details (bank, UPI, IBAN)', '', 'wide'],
      ['cur', 'Currency', CUR_OPTS[0], 'select', CUR_OPTS], ['numfmt', 'Number style', 'Auto (by currency)', 'select', Object.keys(STYLES)],
      ['datefmt', 'Date format', 'DD-MM-YYYY', 'select', DATES], ['taxname', 'Tax name (GST, VAT, sales tax …)', 'GST'],
      ['round', 'Round the total to a whole number', true, 'check']],
    fields: [['kind', 'Document', 'Quotation', 'select', ['Quotation', 'Proforma invoice', 'Estimate']], ['no', 'Number', 'QT-001'], ['date', 'Date', L.today(), 'date'],
      ['days', 'Valid for (days)', '15', 'number'], ['ref', 'Customer reference / enquiry no.', ''], ['cname', 'Customer name', 'Greenfield Hospital', 'wide'],
      ['ccontact', 'Contact person', 'Dr. Meera Nair'], ['cphone', 'Customer phone', ''], ['cemail', 'Customer email', ''], ['ctax', 'Customer tax ID (if any)', ''],
      ['caddr', 'Customer address', 'Hospital Road, Kochi, Kerala 682016', 'wide'],
      ['taxmode', 'Tax', MODES[0], 'select', MODES], ['odisc', 'Overall discount (amount)', '', 'number'],
      ['chlabel', 'Other charges name', 'Freight and packing'], ['charge', 'Other charges (amount, added after tax)', '', 'number'],
      ['pay', 'Payment terms', '50% advance with the order, balance before dispatch', 'wide'], ['deliv', 'Delivery', 'Within 7 working days of the confirmed order', 'wide'],
      ['terms', 'Other terms', 'Warranty as per manufacturer. Prices may change after the validity date.', 'wide']],
    groups: [{ key: 'items', title: 'Items', add: 'Add line',
      cols: [['desc', 'Description', '', 'wide'], ['qty', 'Qty', '1', 'number'], ['unit', 'Unit', 'Nos'], ['rate', 'Rate', '', 'number'], ['disc', 'Disc %', '', 'number'], ['tax', 'Tax %', '18', 'number']],
      sample: [{ desc: 'Laptop, Core i5, 16 GB RAM, 512 GB SSD, 3-year warranty', qty: '5', unit: 'Nos', rate: '61500', disc: '2', tax: '18' },
               { desc: 'Wireless keyboard and mouse set', qty: '5', unit: 'Sets', rate: '1450', tax: '18' },
               { desc: 'On-site installation and data transfer', qty: '1', unit: 'Job', rate: '2500', tax: '18' }] }],
    make: function (P, D, G) {
      var c = curOf(P.cur), loc = STYLES[P.numfmt] || (c[5] === 'in' ? 'en-IN' : 'en-US'), sw = loc === 'en-US-ch';
      var nf = new Intl.NumberFormat(sw ? 'en-US' : loc, { minimumFractionDigits: c[6], maximumFractionDigits: c[6] });
      function m(n) { var s = nf.format(Math.abs(n)); if (sw) s = s.replace(/,/g, '\''); return (n < 0 ? '- ' : '') + c[1] + ' ' + s; }
      function fd(iso) { return fmtDate(iso, P.datefmt); }
      var kind = D.kind || 'Quotation', mode = D.taxmode || MODES[0], tx = mode !== 'No tax', incl = mode === 'Prices include tax', tn = P.taxname || 'Tax';
      var lines = [], gross = 0, anyDisc = false;
      G.items.forEach(function (l) {
        if (!l.desc && !num(l.rate)) return;
        var amt = r2(num(l.qty) * num(l.rate) * (1 - num(l.disc) / 100));
        if (num(l.disc)) anyDisc = true;
        gross += amt; lines.push({ l: l, amt: amt });
      });
      gross = r2(gross);
      var od = Math.min(Math.max(0, num(D.odisc)), gross), f = gross ? (gross - od) / gross : 1, byRate = {}, rates = [];
      lines.forEach(function (x) {
        var rt = tx ? num(x.l.tax) : 0, a = x.amt * f;
        if (!rt) return;
        var t = incl ? a - a / (1 + rt / 100) : a * rt / 100;
        if (!(rt in byRate)) { byRate[rt] = 0; rates.push(rt); }
        byRate[rt] += t;
      });
      var tax = 0; rates.sort(function (a, b) { return a - b; }).forEach(function (rt) { byRate[rt] = r2(byRate[rt]); tax += byRate[rt]; });
      tax = r2(tax);
      var after = r2(gross - od), base = incl ? r2(after - tax) : after, charge = Math.max(0, num(D.charge));
      var raw = r2(base + tax + charge), gt = P.round ? Math.round(raw) : raw;
      var cols = [['#', 'c'], ['Description', ''], ['Qty', 'r'], ['Unit', ''], ['Rate', 'r']];
      if (anyDisc) cols.push(['Disc', 'r']);
      if (tx) cols.push([tn + ' %', 'r']);
      cols.push([incl ? 'Amount (incl. ' + tn + ')' : 'Amount', 'r']);
      var rows = lines.map(function (x, i) {
        var r = [i + 1, x.l.desc || '', qty(num(x.l.qty)), x.l.unit || '', m(num(x.l.rate))];
        if (anyDisc) r.push(pct(num(x.l.disc)));
        if (tx) r.push(pct(num(x.l.tax)) || '0%');
        r.push(m(x.amt)); return r;
      });
      var tot = [];
      if (od || tx || charge) tot.push(['Sub-total', m(gross)]);
      if (od) tot.push(['Less discount', m(-od)]);
      if (tx && (incl || od)) tot.push([incl ? 'Value before ' + tn : 'Taxable value', m(base)]);
      rates.forEach(function (rt) { tot.push([tn + ' @ ' + r2(rt) + '%' + (incl ? ' (included)' : ''), m(byRate[rt])]); });
      if (charge) tot.push([D.chlabel || 'Other charges', m(charge)]);
      if (gt !== raw) tot.push(['Round off', m(r2(gt - raw))]);
      tot.push([kind === 'Proforma invoice' ? 'Amount payable' : 'Total', m(gt), 1]);
      var until = num(D.days) > 0 ? addDays(D.date, num(D.days)) : '';
      var notes = ['All amounts in ' + c[2] + ' (' + c[0] + ').' + (incl ? ' Prices include ' + tn + '.' : '')];
      if (D.pay) notes.push('Payment terms: ' + D.pay);
      if (D.deliv) notes.push('Delivery: ' + D.deliv);
      if (D.terms) notes.push(D.terms);
      if (until && kind !== 'Proforma invoice') notes.push('This ' + kind.toLowerCase() + ' is valid until ' + fd(until) + '.');
      if (kind === 'Proforma invoice') notes.push('This is a proforma invoice, not a tax invoice. The final invoice will be issued when the goods or services are supplied.');
      if (P.bank) notes.push('Payment details: ' + P.bank);
      var no = kind === 'Proforma invoice' ? 'Proforma No.' : kind === 'Estimate' ? 'Estimate No.' : 'Quotation No.';
      return { file: kind.replace(/ /g, '-') + '-' + safe(D.no), title: kind.toUpperCase(), org: P.name,
        sub: [P.addr, [P.phone ? 'Phone: ' + P.phone : '', P.email, P.web, P.taxid ? (P.taxlabel || 'Tax ID') + ': ' + P.taxid : ''].filter(Boolean).join('  ·  ')],
        kv: [[no, D.no], ['Date', fd(D.date)], ['To', D.cname], [kind === 'Proforma invoice' ? 'Pay by' : 'Valid until', until ? fd(until) + ' (' + num(D.days) + ' days)' : ''],
             ['Contact', [D.ccontact, D.cphone, D.cemail].filter(Boolean).join(', ')], ['Your reference', D.ref], ['Address', D.caddr],
             [D.ctax ? 'Customer ' + (P.taxlabel || 'tax ID') : '', D.ctax || '']],
        cols: cols, rows: rows, totals: tot, words: gt ? inWords(gt, c) : '', notes: notes,
        sign: ['Accepted by customer (signature and date)', 'For ' + (P.name || '') + '\n\nAuthorised signatory'] };
    }
  });
})();
