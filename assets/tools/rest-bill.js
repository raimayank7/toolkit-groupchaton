/* Restaurant bill / KOT (kit: restaurant). Same order prints as a kitchen order ticket (no prices) or as a bill with
   discount, voluntary service charge, packing charge and GST (5% / 18% split into CGST + SGST, or none for composition). */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  function now() { var d = new Date(); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
  var GSTS = ['5% (restaurant)', '18% (hotel restaurant, with ITC)', 'No GST (composition / unregistered)'];
  TKTools.register('rest-bill', {
    name: 'bill', no: 'no',
    profile: [['name', 'Restaurant name', 'Spice Route Family Restaurant'], ['addr', 'Address', 'SCO 21, Sector 17, Chandigarh 160017', 'wide'], ['phone', 'Phone', '98140 00000'],
      ['gstin', 'GSTIN', '04ABCDE1234F1Z5'], ['fssai', 'FSSAI licence / registration no.', '']],
    fields: [['kind', 'Document', 'Bill', 'select', ['Bill', 'KOT (kitchen order ticket)']], ['no', 'Bill / KOT number', 'B-001'], ['date', 'Date', L.today(), 'date'], ['time', 'Time', now(), 'time'],
      ['otype', 'Order type', 'Dine-in', 'select', ['Dine-in', 'Takeaway', 'Delivery']], ['table', 'Table no.', 'T4'], ['covers', 'Covers (guests)', '2', 'number'], ['stew', 'Steward', 'Aman'],
      ['cname', 'Guest name / phone (optional)', ''], ['gst', 'GST', GSTS[0], 'select', GSTS], ['sc', 'Service charge % (voluntary, usually 0)', '', 'number'],
      ['disc', 'Discount %', '', 'number'], ['pack', 'Packing / delivery charge (₹)', '', 'number'], ['pay', 'Payment mode', 'UPI', 'select', ['Cash', 'UPI', 'Card', 'Swiggy', 'Zomato', 'Other', 'Not paid yet']],
      ['terms', 'Note on the bill', 'Thank you for dining with us. Please visit again!', 'wide']],
    groups: [{ key: 'items', title: 'Order', add: 'Add dish', cols: [['desc', 'Dish', '', 'wide'], ['qty', 'Qty', '1', 'number'], ['rate', 'Rate (₹)', '', 'number'], ['note', 'Kitchen note', '']],
      sample: [{ desc: 'Paneer Butter Masala', qty: '1', rate: '280' }, { desc: 'Dal Makhani', qty: '1', rate: '240', note: 'less spicy' }, { desc: 'Butter Naan', qty: '4', rate: '50' },
        { desc: 'Jeera Rice', qty: '1', rate: '180' }, { desc: 'Sweet Lassi', qty: '2', rate: '90', note: 'one without sugar' }] }],
    make: function (P, D, G) {
      var kot = /^KOT/.test(D.kind || ''), when = L.dmy(D.date) + (D.time ? '  ' + D.time : ''), dine = D.otype === 'Dine-in';
      var lines = G.items.filter(function (l) { return l.desc || num(l.rate); });
      var head = [P.phone ? 'Phone: ' + P.phone : '', P.gstin ? 'GSTIN: ' + P.gstin : '', P.fssai ? 'FSSAI Lic. No.: ' + P.fssai : ''].filter(Boolean).join('  ·  ');
      if (kot) {
        var q = 0, rows = lines.map(function (l, i) { q += num(l.qty); return [i + 1, l.desc, num(l.qty), l.note || '']; });
        return { file: 'KOT-' + (D.no || 'order').replace(/[^\w\-]+/g, ''), title: 'KOT  ·  KITCHEN ORDER TICKET', org: P.name, sub: [],
          kv: [['KOT No.', D.no], ['Date & time', when], ['Order type', D.otype], [dine ? 'Table' : 'Guest', dine ? D.table : D.cname], dine ? ['Covers', D.covers] : ['', ''], ['Steward', D.stew]],
          cols: [['#', 'c'], ['Dish', ''], ['Qty', 'r'], ['Kitchen note', '']], rows: rows, totals: [['Total quantity', String(q), 1]],
          notes: [], sign: ['Steward', 'Kitchen'] };
      }
      var sub = 0, rows2 = lines.map(function (l, i) { var a = r2(num(l.qty) * num(l.rate)); sub += a; return [i + 1, l.desc, num(l.qty), num(l.rate), a]; });
      sub = r2(sub);
      var dp = num(D.disc), disc = r2(sub * dp / 100), after = r2(sub - disc), sp = num(D.sc), sc = r2(after * sp / 100), pack = r2(num(D.pack));
      var rate = parseFloat(D.gst) || 0, taxable = r2(after + sc + pack), tax = r2(taxable * rate / 100), cg = r2(tax / 2), sg = r2(tax - cg);
      var raw = r2(taxable + tax), gt = Math.round(raw), pc = function (n) { return (Math.round(n * 100) / 100) + '%'; };
      var tot = [['Food & beverages', sub]];
      if (disc) tot.push(['Discount (' + pc(dp) + ')', -disc]);
      if (sc) tot.push(['Service charge (' + pc(sp) + ', voluntary)', sc]);
      if (pack) tot.push([D.otype === 'Delivery' ? 'Packing & delivery charge' : 'Packing charge', pack]);
      if (rate) tot.push(['Taxable value', taxable], ['CGST @ ' + pc(rate / 2), cg], ['SGST @ ' + pc(rate / 2), sg]);
      if (gt !== raw) tot.push(['Round off', r2(gt - raw)]);
      tot.push(['Amount payable', gt, 1]);
      var notes = [];
      if (!rate && P.gstin) notes.push('Composition taxable person, not eligible to collect tax on supplies.');
      if (sc) notes.push('Service charge is voluntary. It is added only with your consent and will be removed on request.');
      if (D.pay) notes.push('Payment: ' + D.pay);
      if (D.terms) notes.push(D.terms);
      return { file: 'Bill-' + (D.no || 'restaurant').replace(/[^\w\-]+/g, ''), title: rate ? 'TAX INVOICE' : (P.gstin ? 'BILL OF SUPPLY' : 'BILL'), org: P.name, sub: [P.addr, head],
        kv: [['Bill No.', D.no], ['Date & time', when], ['Order type', D.otype], [dine ? 'Table' : 'Guest', dine ? D.table : D.cname], dine ? ['Covers', D.covers] : ['', ''], ['Steward', D.stew],
             dine && D.cname ? ['Guest', D.cname] : ['', ''], ['', '']],
        cols: [['#', 'c'], ['Dish', ''], ['Qty', 'r'], ['Rate', 'r$'], ['Amount', 'r$']], rows: rows2, totals: tot, words: L.words(gt),
        notes: notes, sign: ['Guest signature', 'For ' + (P.name || '')] };
    }
  });
})();
