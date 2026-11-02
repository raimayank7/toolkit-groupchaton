/* Contractor estimate / quotation maker (Contractors & Construction toolkit): BOQ lines grouped by section, quantities typed as
   measurements (2×4.5×0.23 + 1×3×0.23), contingency and GST. Uses the shared engine in assets/tools.js. */
(function () {
  var L = TKTools.lib, num = L.num, r2 = L.r2;
  var SECTIONS = ['Earthwork', 'Concrete (PCC / RCC)', 'Reinforcement steel', 'Masonry', 'Plastering', 'Flooring & tiling', 'Doors & windows',
    'Painting', 'Plumbing & sanitary', 'Electrical', 'Waterproofing', 'Interior work', 'Miscellaneous'];
  var UNITS = ['cum', 'sqm', 'rmt', 'kg', 'MT', 'nos', 'point', 'sqft', 'cft', 'rft', 'LS', 'job', 'bag', 'litre', 'day'];

  /* "12×1.5×1.5×1.5 + 48×0.9×0.6 - 2×1.2×1.5" → number (NaN when it can't be read). No eval. */
  function qtyOf(s) {
    s = String(s == null ? '' : s).replace(/,/g, '').replace(/[×xX*]/g, '*').replace(/[−–]/g, '-').replace(/\s+/g, '');
    if (!s) return 0;
    var terms = s.match(/[+-]?[^+-]+/g), total = 0;
    if (!terms || terms.join('') !== s) return NaN;
    for (var i = 0; i < terms.length; i++) {
      var t = terms[i], sign = 1;
      if (t[0] === '-') { sign = -1; t = t.slice(1); } else if (t[0] === '+') t = t.slice(1);
      var p = 1, f = t.split('*');
      for (var j = 0; j < f.length; j++) { if (!/^\d*\.?\d+$/.test(f[j])) return NaN; p *= parseFloat(f[j]); }
      total += sign * p;
    }
    return Math.round(total * 1000) / 1000;
  }

  TKTools.register('contractor-estimate', {
    name: 'estimate', no: 'no',
    profile: [['name', 'Firm / contractor name', 'Sharma Builders & Contractors'], ['addr', 'Address', '45, Industrial Area-B, Ludhiana, Punjab 141003', 'wide'],
      ['phone', 'Phone', '98140 00000'], ['email', 'Email', ''], ['gstin', 'GSTIN (if registered)', ''], ['reg', 'Registration / licence No.', '']],
    fields: [['kind', 'Document', 'Estimate', 'select', ['Estimate', 'Quotation']], ['no', 'Number', 'EST-001'], ['date', 'Date', L.today(), 'date'],
      ['valid', 'Valid until', '', 'date'], ['client', 'Client name', 'Mr. Harpreet Singh', 'wide'], ['site', 'Site address', 'Plot 27, Green Avenue, Ludhiana', 'wide'],
      ['work', 'Name of work', 'Construction of ground-floor residence, 1,200 sq ft', 'wide'],
      ['cont', 'Contingency %', '3', 'number'], ['gst', 'GST %', '18', 'select', ['0', '5', '12', '18']],
      ['terms', 'Payment terms', '10% advance on order; running bills every month on measured work, payable within 7 days; 5% retained until 6 months after completion.', 'wide'],
      ['excl', 'Exclusions (client to arrange)', 'Water and electricity for construction, government approvals and fees, soil testing, and any work not listed above.', 'wide']],
    groups: [{ key: 'items', title: 'BOQ lines', add: 'Add BOQ line',
      cols: [['sec', 'Section', 'Earthwork', 'select', SECTIONS], ['desc', 'Item description', '', 'wide'], ['unit', 'Unit', 'cum', 'select', UNITS],
        ['qty', 'Quantity (or 2×4.5×0.23)', ''], ['rate', 'Rate (₹)', '', 'number']],
      sample: [{ sec: 'Earthwork', desc: 'Excavation for foundations in ordinary soil, disposal within 50 m', unit: 'cum', qty: '12×1.5×1.5×1.5 + 48×0.9×0.6', rate: '280' },
        { sec: 'Concrete (PCC / RCC)', desc: 'PCC 1:4:8 bed 100 mm thick below footings, including curing', unit: 'cum', qty: '12×1.5×1.5×0.1', rate: '5600' },
        { sec: 'Concrete (PCC / RCC)', desc: 'RCC M20 in footings, columns, beams and slab incl. shuttering (steel separate)', unit: 'cum', qty: '38.5', rate: '7800' },
        { sec: 'Reinforcement steel', desc: 'TMT Fe 500D reinforcement: cutting, bending, binding and placing', unit: 'kg', qty: '4200', rate: '78' },
        { sec: 'Masonry', desc: 'Brick masonry 230 mm in cement mortar 1:6, superstructure', unit: 'cum', qty: '46×0.23×3 - 6×1.2×0.23×1.5', rate: '6800' },
        { sec: 'Plastering', desc: '12 mm cement plaster 1:4 on internal walls', unit: 'sqm', qty: '620', rate: '290' }] }],
    make: function (P, D, G) {
      var order = [], by = {}, bad = 0;
      G.items.forEach(function (l) {
        if (!l.desc && !num(l.rate)) return;
        var q = qtyOf(l.qty), s = l.sec || 'Miscellaneous';
        if (isNaN(q)) { bad++; q = 0; }
        if (!by[s]) { by[s] = { lines: [], sum: 0 }; order.push(s); }
        var amt = r2(q * num(l.rate));
        by[s].lines.push([l.desc, l.unit, isNaN(qtyOf(l.qty)) ? 'check: ' + l.qty : q, num(l.rate), amt]); by[s].sum += amt;
      });
      var rows = [], n = 0, works = 0, many = order.length > 1, tot = [];
      order.forEach(function (s, si) {
        if (many) rows.push([String.fromCharCode(65 + si % 26), s.toUpperCase(), '', '', '', '']);
        by[s].lines.forEach(function (x) { rows.push([++n].concat(x)); });
        by[s].sum = r2(by[s].sum); works += by[s].sum;
        if (many) tot.push([String.fromCharCode(65 + si % 26) + '. ' + s, by[s].sum]);
      });
      works = r2(works);
      var cp = num(D.cont), cont = r2(works * cp / 100), before = r2(works + cont), gp = num(D.gst), gst = r2(before * gp / 100), raw = r2(before + gst), gt = Math.round(raw);
      tot.push(['Total cost of works', works, 1]);
      if (cont) tot.push(['Contingency @ ' + cp + '%', cont]);
      if (gst) tot.push(['GST @ ' + gp + '%', gst]);
      if (gt !== raw) tot.push(['Round off', r2(gt - raw)]);
      tot.push([D.kind === 'Quotation' ? 'Total quoted amount' : 'Estimated total', gt, 1]);
      var q = D.kind === 'Quotation';
      var notes = ['Quantities are ' + (q ? 'based on the drawings and BOQ provided' : 'estimated') + '; payment will be for actual quantities measured jointly at the rates above.'];
      if (D.terms) notes.push('Payment terms: ' + D.terms);
      if (D.excl) notes.push('Exclusions: ' + D.excl);
      if (q && D.valid) notes.push('This quotation is valid until ' + L.dmy(D.valid) + '. Rates of cement, steel and other materials may change after that date.');
      if (bad) notes.unshift(bad + ' quantity ' + (bad > 1 ? 'entries could' : 'entry could') + ' not be read and counted as 0. Use numbers with × and + only, for example 2×4.5×0.23.');
      return { file: (q ? 'Quotation-' : 'Estimate-') + String(D.no || 'document').replace(/[^\w\-]+/g, '-'), title: q ? 'QUOTATION' : 'ESTIMATE OF COST', org: P.name,
        sub: [P.addr, [P.phone ? 'Phone: ' + P.phone : '', P.email, P.gstin ? 'GSTIN: ' + P.gstin : '', P.reg ? 'Reg. No.: ' + P.reg : ''].filter(Boolean).join('  ·  ')],
        kv: [[q ? 'Quotation No.' : 'Estimate No.', D.no], ['Date', L.dmy(D.date)], ['Client', D.client], [q ? 'Valid until' : '', q ? L.dmy(D.valid) : ''],
             ['Site', D.site], ['Name of work', D.work]],
        cols: [['#', 'c'], ['Item description', ''], ['Unit', ''], ['Qty', 'r'], ['Rate', 'r$'], ['Amount', 'r$']],
        rows: rows, totals: tot, words: L.words(gt), notes: notes,
        sign: ['Accepted by client (signature and date)', 'For ' + (P.name || '') + '\n\nAuthorised signatory'] };
    }
  });
})();
