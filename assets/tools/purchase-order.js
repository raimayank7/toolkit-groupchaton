/* Online purchase order maker (Purchase & Warehouse toolkit). Any currency and date format; tax label of your choice.
   Indian rupee keeps numbers as numbers (₹ format, lakh/crore words); other currencies are written as formatted text. */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  /* code, symbol, money words, minor unit words, number locale */
  var CUR = [['INR', '₹', 'Rupees', 'Paise', 'en-IN'], ['USD', '$', 'US Dollars', 'Cents', 'en-US'], ['EUR', '€', 'Euros', 'Cents', 'de-DE'],
    ['GBP', '£', 'Pounds Sterling', 'Pence', 'en-GB'], ['AED', 'AED', 'UAE Dirhams', 'Fils', 'en-US'], ['SAR', 'SAR', 'Saudi Riyals', 'Halalas', 'en-US'],
    ['SGD', 'S$', 'Singapore Dollars', 'Cents', 'en-SG'], ['AUD', 'A$', 'Australian Dollars', 'Cents', 'en-AU'], ['CAD', 'C$', 'Canadian Dollars', 'Cents', 'en-CA'],
    ['ZAR', 'R', 'Rand', 'Cents', 'en-ZA'], ['NGN', '₦', 'Naira', 'Kobo', 'en-NG'], ['KES', 'KSh', 'Kenyan Shillings', 'Cents', 'en-KE'],
    ['PKR', 'Rs', 'Pakistani Rupees', 'Paisa', 'en-PK'], ['BDT', '৳', 'Taka', 'Poisha', 'en-IN'], ['LKR', 'Rs', 'Sri Lankan Rupees', 'Cents', 'en-LK'],
    ['NPR', 'Rs', 'Nepalese Rupees', 'Paisa', 'en-IN'], ['MYR', 'RM', 'Ringgit', 'Sen', 'en-MY']];
  var CUR_OPTS = CUR.map(function (c) { return c[0] + ' (' + c[1] + ')'; });
  var DATES = ['DD-MM-YYYY', 'DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD', 'DD Mon YYYY'];
  var TERMS = ['Delivered to our site, freight paid by supplier', 'Ex-works: we collect from your premises', 'Door delivery, freight extra at actuals',
    'EXW (Incoterms 2020)', 'FCA (Incoterms 2020)', 'FOB (Incoterms 2020)', 'CFR (Incoterms 2020)', 'CIF (Incoterms 2020)', 'DAP (Incoterms 2020)', 'DDP (Incoterms 2020)'];
  var PAY = ['30 days from receipt of material and invoice', '15 days from receipt of material and invoice', '45 days from receipt of material and invoice',
    '60 days from receipt of material and invoice', '100% advance with this order', '50% advance, balance before dispatch', 'Against delivery', 'Letter of credit at sight'];

  /* sensible first-time defaults from the device's time zone (the user can change them; they are then remembered) */
  var tz = ''; try { tz = Intl.DateTimeFormat().resolvedOptions().timeZone || ''; } catch (e) {}
  var ZONE = { 'Asia/Kolkata': 'INR', 'Asia/Calcutta': 'INR', 'Europe/London': 'GBP', 'Asia/Dubai': 'AED', 'Asia/Riyadh': 'SAR', 'Asia/Singapore': 'SGD',
    'Africa/Johannesburg': 'ZAR', 'Africa/Lagos': 'NGN', 'Africa/Nairobi': 'KES', 'Asia/Karachi': 'PKR', 'Asia/Dhaka': 'BDT', 'Asia/Colombo': 'LKR',
    'Asia/Kathmandu': 'NPR', 'Asia/Katmandu': 'NPR', 'Asia/Kuala_Lumpur': 'MYR', 'America/Toronto': 'CAD', 'America/Vancouver': 'CAD' };
  var c0 = ZONE[tz] || (/^Australia\//.test(tz) ? 'AUD' : /^Europe\//.test(tz) ? 'EUR' : 'USD');
  var india = c0 === 'INR';
  var d0 = india ? 'DD-MM-YYYY' : /^America\//.test(tz) && c0 !== 'CAD' ? 'MM/DD/YYYY' : 'DD Mon YYYY';
  var t0 = { INR: 'GST', AUD: 'GST', SGD: 'GST', MYR: 'SST', CAD: 'GST/HST', USD: 'Sales tax' }[c0] || 'VAT';

  function curOf(P) { var k = String(P.cur || '').slice(0, 3); for (var i = 0; i < CUR.length; i++) if (CUR[i][0] === k) return CUR[i]; return CUR[1]; }
  function fmtDate(iso, f) {
    if (!iso) return ''; var p = String(iso).split('-'); if (p.length !== 3) return iso;
    var M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return f === 'DD/MM/YYYY' ? p[2] + '/' + p[1] + '/' + p[0] : f === 'MM/DD/YYYY' ? p[1] + '/' + p[2] + '/' + p[0] : f === 'YYYY-MM-DD' ? iso
      : f === 'DD Mon YYYY' ? p[2] + ' ' + M[+p[1] - 1] + ' ' + p[0] : p[2] + '-' + p[1] + '-' + p[0];
  }
  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  function w3(n) { var h = Math.floor(n / 100), t = n % 100, o = [];
    if (h) o.push(ONES[h] + ' Hundred'); if (t) o.push(t < 20 ? ONES[t] : TENS[Math.floor(t / 10)] + (t % 10 ? '-' + ONES[t % 10] : '')); return o.join(' '); }
  function wordsIntl(amount, c) {   /* international style: billion, million, thousand */
    amount = r2(Math.abs(amount)); var n = Math.floor(amount), p = Math.round((amount - n) * 100), out = [];
    if (!n && !p) return '';
    [[1e9, ' Billion'], [1e6, ' Million'], [1e3, ' Thousand'], [1, '']].forEach(function (s) {
      var k = Math.floor(n / s[0]) % 1000; if (k) out.push(w3(k) + s[1]);
    });
    return c[2] + ' ' + (out.length ? out.join(' ') : 'Zero') + (p ? ' and ' + w3(p) + ' ' + c[3] : '') + ' Only';
  }
  function safe(s) { return String(s || 'PO').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'PO'; }

  TKTools.register('purchase-order', {
    name: 'purchase order', no: 'no',
    profile: [['name', 'Company name', 'Apex Engineering Works'], ['addr', 'Address', 'Plot 18, Focal Point, Ludhiana, Punjab 141010', 'wide'],
      ['taxid', 'Tax ID (GSTIN / VAT / EIN)', ''], ['phone', 'Phone', ''], ['email', 'Email', ''], ['taxlabel', 'Tax label', t0],
      ['cur', 'Currency', CUR_OPTS[CUR.map(function (c) { return c[0]; }).indexOf(c0)], 'select', CUR_OPTS], ['datefmt', 'Date format', d0, 'select', DATES]],
    fields: [['no', 'PO number', 'PO-001'], ['date', 'PO date', L.today(), 'date'], ['sname', 'Supplier name', 'Bharat Steel Traders', 'wide'],
      ['saddr', 'Supplier address', 'G.T. Road, Ludhiana, Punjab', 'wide'], ['staxid', 'Supplier tax ID', ''], ['quote', 'Supplier quotation ref.', 'Q-2291 dated 28-10-2026'],
      ['deliv', 'Delivery date', '', 'date'], ['dterms', 'Delivery terms', TERMS[0], 'select', TERMS], ['dplace', 'Deliver to (address or port)', 'Main store, Plot 18, Focal Point, Ludhiana', 'wide'],
      ['pay', 'Payment terms', PAY[0], 'select', PAY], ['freight', 'Freight / other charges', '', 'number'], ['contact', 'Contact for this order', 'Purchase dept., 98140 00000'],
      ['terms', 'Other conditions', 'Send the mill test certificate with the material. Delivery between 9 am and 5 pm, Monday to Saturday.', 'wide']],
    groups: [{ key: 'items', title: 'Items', add: 'Add item',
      cols: [['code', 'Item code', ''], ['desc', 'Description', '', 'wide'], ['qty', 'Qty', '1', 'number'], ['unit', 'Unit', 'Nos'], ['rate', 'Rate', '', 'number'],
        ['disc', 'Disc %', '', 'number'], ['tax', 'Tax %', '18', 'number']],
      sample: [{ code: 'RM-001', desc: 'MS round bar 20 mm, IS 2062 E250', qty: '2500', unit: 'kg', rate: '68', tax: '18' },
        { code: 'RM-002', desc: 'MS plate 10 mm, IS 2062 E250', qty: '1200', unit: 'kg', rate: '72', disc: '2', tax: '18' }] }],
    make: function (P, D, G) {
      var c = curOf(P), inr = c[0] === 'INR', tl = P.taxlabel || 'Tax';
      var nf = new Intl.NumberFormat(c[4], { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      var m = function (v) { return inr ? r2(v) : c[1] + ' ' + nf.format(r2(v)); };
      var sub = 0, rows = [], byRate = {};
      G.items.forEach(function (l) {
        if (!l.desc && !num(l.rate)) return;
        var a = r2(num(l.qty) * num(l.rate) * (1 - num(l.disc) / 100)), t = r2(a * num(l.tax) / 100);
        sub += a; byRate[num(l.tax)] = (byRate[num(l.tax)] || 0) + t;
        rows.push([rows.length + 1, l.code || '', l.desc, num(l.qty), l.unit, m(num(l.rate)), num(l.disc) ? num(l.disc) + '%' : '', num(l.tax) + '%', m(a)]);
      });
      sub = r2(sub);
      var tot = [['Sub-total (before ' + tl + ')', m(sub)]], tax = 0;
      Object.keys(byRate).map(Number).sort(function (a, b) { return a - b; }).forEach(function (k) { if (!k) return; tax += byRate[k]; tot.push([tl + ' @ ' + k + '%', m(byRate[k])]); });
      tax = r2(tax);
      var fr = num(D.freight); if (fr) tot.push(['Freight / other charges', m(fr)]);
      var gt = r2(sub + tax + fr);
      tot.push(['Total order value (' + c[0] + ')', m(gt), 1]);
      var f = P.datefmt;
      return { file: 'Purchase-Order-' + safe(D.no), title: 'PURCHASE ORDER', org: P.name,
        sub: [P.addr, [P.taxid ? 'Tax ID: ' + P.taxid : '', P.phone ? 'Phone: ' + P.phone : '', P.email].filter(Boolean).join('  ·  ')],
        kv: [['PO No.', D.no], ['PO date', fmtDate(D.date, f)], ['Supplier', D.sname], ['Supplier tax ID', D.staxid], ['Address', D.saddr], ['Your quotation', D.quote],
             ['Delivery date', fmtDate(D.deliv, f)], ['Delivery terms', D.dterms], ['Deliver to', D.dplace], ['Payment terms', D.pay], ['Contact', D.contact], ['Currency', c[0]]],
        cols: [['#', 'c'], ['Code', ''], ['Description', ''], ['Qty', 'r'], ['Unit', ''], ['Rate', inr ? 'r$' : 'r'], ['Disc', 'r'], [tl + ' %', 'r'], ['Amount', inr ? 'r$' : 'r']],
        rows: rows, totals: tot, words: inr ? L.words(gt) : wordsIntl(gt, c),
        notes: ['Please quote our PO number on your invoice, delivery challan or packing list and all correspondence.',
                'Material must match the description, specification and quantity ordered. Rejected or excess material will be returned at the supplier\'s cost.',
                'Please confirm acceptance of this order and the delivery date within 2 working days.', D.terms].filter(Boolean),
        sign: ['Accepted by supplier (sign, stamp and return)', 'For ' + (P.name || '') + '\n\nAuthorised signatory'] };
    }
  });
})();
