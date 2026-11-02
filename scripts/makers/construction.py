"""Contractors & Construction toolkit files (kit id: construction).
   python -X utf8 scripts/make_files.py --only construction"""
from make_files import *

ORANGE = 'EA580C'
UNITS = ['cum', 'sqm', 'rmt', 'kg', 'MT', 'nos', 'point', 'sqft', 'cft', 'rft', 'LS', 'job', 'bag', 'litre', 'day']
QTY = '#,##0.000;-#,##0.000;'
NUM = '#,##0.00;-#,##0.00;'

def dims(ws, r1, r2, c_nos, c_last, c_out):
    """Nos × L × B × D: product of the typed dimensions (empty ones ignored), negative Nos = deduction."""
    for r in range(r1, r2 + 1):
        cell(ws, '%s%d' % (c_out, r), '=IF(COUNT({a}{r}:{b}{r})=0,"",ROUND(PRODUCT({a}{r}:{b}{r}),3))'.format(a=c_nos, b=c_last, r=r), 'calc', QTY)

# ------------------------------------------------------------------ Estimate & BOQ
def estimate_boq():
    wb = new_wb('con-estimate-boq-excel')
    howto(wb, first=True)
    pj = sheet(wb, 'Project', [34, 46, 3, 28], tab=ORANGE)
    ms = sheet(wb, 'Measurements', [9, 40, 8, 10, 10, 10, 12])
    bq = sheet(wb, 'BOQ', [8, 20, 52, 8, 11, 12, 12, 12, 15], tab=ORANGE)
    ab = sheet(wb, 'Abstract', [7, 40, 18, 12], tab=ORANGE)
    lists_sheet(wb, extra={5: UNITS})
    # Project
    title(pj, 'Project details', 'Fill once. The BOQ and the abstract of cost use these.')
    rows = [('Name of work', 'Construction of ground-floor residence (G+0), 1,200 sq ft', None), ('Client / owner', 'Mr. Harpreet Singh', None),
            ('Site location', 'Plot 27, Green Avenue, Ludhiana', None), ('Prepared by', 'Sharma Builders & Contractors', None),
            ('Estimate No.', 'EST-2026-014', None), ('Date', TODAY, DATE), ('Rates based on', 'Market rates of %s' % TODAY.strftime('%B %Y'), None),
            ('Built-up area (sq ft)', 1200, '#,##0'), ('Overheads & profit % (if rates exclude them)', 0, '0.0%'),
            ('Contingency %', 0.03, '0.0%'), ('GST %', 18, '0')]
    for i, (k, v, f) in enumerate(rows, 3):
        label(pj, 'A%d' % i, k); cell(pj, 'B%d' % i, v, 'in', f, align='left')
    dv(pj, 'B13', '"0,5,12,18"', 'GST rate for this work: check with your tax adviser')
    cell(pj, 'A15', 'Percentages: type 3% (or 0.03). GST: pick 0, 5, 12 or 18.', size=9, color=MUTED)
    head(pj, 2, ['Sections (edit freely)'], col=4)
    SECS = ['Earthwork', 'Concrete (PCC / RCC)', 'Reinforcement steel', 'Masonry', 'Plastering', 'Flooring & tiling', 'Doors & windows',
            'Painting', 'Plumbing & sanitary', 'Electrical', 'Waterproofing', 'Miscellaneous']
    for i in range(3, 23): cell(pj, 'D%d' % i, SECS[i - 3] if i - 3 < len(SECS) else None, 'in')
    # Measurements
    title(ms, 'Measurements', 'Quantity = Nos × Length × Breadth × Depth/Height (empty cells are skipped). Negative Nos deducts.')
    head(ms, 3, ['Item No.', 'Description / location', 'Nos', 'Length (m)', 'Breadth (m)', 'Depth / height (m)', 'Quantity'])
    grid(ms, 4, 503, 'ABCDEF', {}, {'C': 'General', 'D': '0.000', 'E': '0.000', 'F': '0.000'})
    dims(ms, 4, 503, 'C', 'F', 'G')
    for i, row in enumerate([('1', 'Column footings F1', 12, 1.5, 1.5, 1.5), ('1', 'Wall trenches, all round', 1, 48, 0.9, 0.6),
                             ('2', 'PCC bed below footings', 12, 1.5, 1.5, 0.1), ('5', 'Outer walls 230 mm, ground floor', 1, 46, 0.23, 3.0),
                             ('5', 'Deduct door and window openings', -6, 1.2, 0.23, 1.5)], 4):
        for j, v in enumerate(row): ms.cell(i, j + 1, v)
    ms.freeze_panes = 'A4'
    # BOQ
    title(bq, 'Bill of quantities', 'Leave "Qty (typed)" empty to use the quantity from the Measurements sheet.')
    label(bq, 'H2', 'Total'); cell(bq, 'I2', '=SUM(I4:I203)', 'soft', INR, bold=True)
    head(bq, 3, ['Item No.', 'Section', 'Description of item', 'Unit', 'Qty (typed)', 'Qty from measurements', 'Quantity used', 'Rate', 'Amount'])
    grid(bq, 4, 203, 'ABCDEFGHI', {'F': 'calc', 'G': 'calc', 'I': 'calc'}, {'E': QTY, 'F': QTY, 'G': QTY, 'H': INR, 'I': INR},
         {'F': '=IF($A{r}="","",SUMIF(Measurements!$A:$A,$A{r}&"",Measurements!$G:$G))',
          'G': '=IF(AND($A{r}="",$C{r}=""),"",IF($E{r}<>"",$E{r},$F{r}))',
          'I': '=IF(OR($G{r}="",$H{r}=""),"",ROUND(N($G{r})*$H{r},2))'})
    for r in range(4, 204): bq['C%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
    items = [('1', 'Earthwork', 'Earthwork in excavation for foundations in ordinary soil, including dressing and disposal within 50 m', 'cum', None, 280),
             ('2', 'Concrete (PCC / RCC)', 'PCC 1:4:8 (M7.5) bed 100 mm thick below footings, including curing', 'cum', None, 5600),
             ('3', 'Concrete (PCC / RCC)', 'RCC M20 in footings, columns, beams and roof slab including shuttering and curing (steel paid separately)', 'cum', 38.5, 7800),
             ('4', 'Reinforcement steel', 'TMT steel Fe 500D reinforcement: cutting, bending, binding with wire and placing in position', 'kg', 4200, 78),
             ('5', 'Masonry', 'Brick masonry 230 mm thick in cement mortar 1:6 in superstructure', 'cum', None, 6800),
             ('6', 'Plastering', '12 mm cement plaster 1:4 on internal walls, finished smooth', 'sqm', 620, 290),
             ('7', 'Flooring & tiling', 'Vitrified tiles 600 × 600 mm on cement mortar bed, including grouting', 'sqm', 105, 1150),
             ('8', 'Doors & windows', 'UPVC sliding windows with toughened glass and mesh, fixed complete', 'sqm', 14, 5200),
             ('9', 'Painting', 'Two coats acrylic emulsion over one coat primer and wall putty', 'sqm', 780, 165),
             ('10', 'Plumbing & sanitary', 'Plumbing and sanitary work with CPVC lines and fittings, complete as per drawing', 'LS', 1, 145000),
             ('11', 'Electrical', 'Concealed electrical wiring with FR copper wire and modular switches', 'point', 64, 1450),
             ('12', 'Waterproofing', 'Roof waterproofing with polymer-modified cementitious coating, two coats', 'sqm', 115, 340)]
    for i, (no, sec, desc, unit, q, rate) in enumerate(items, 4):
        bq['A%d' % i], bq['B%d' % i], bq['C%d' % i], bq['D%d' % i], bq['H%d' % i] = no, sec, desc, unit, rate
        if q is not None: bq['E%d' % i] = q
        bq.row_dimensions[i].height = 30
    dv(bq, 'B4:B203', '=Project!$D$3:$D$22', 'Pick a section from the Project sheet')
    dv(bq, 'D4:D203', '=Lists!$E$1:$E$%d' % len(UNITS))
    bq.freeze_panes = 'D4'
    bq.page_setup.orientation = 'landscape'
    # Abstract
    ws = ab
    ws.merge_cells('A1:D1'); cell(ws, 'A1', 'ABSTRACT OF COST', bold=True, size=15, color=ORANGE, align='center')
    ws.merge_cells('A2:D2'); cell(ws, 'A2', '=Project!B3', bold=True, align='center', wrap=True)
    ws.row_dimensions[2].height = 30
    ws.merge_cells('A3:D3'); cell(ws, 'A3', '="Client: "&Project!B4&"   ·   Site: "&Project!B5&"   ·   "&Project!B7&" dated "&TEXT(Project!B8,"dd-mm-yyyy")', size=9, align='center', wrap=True)
    ws.row_dimensions[3].height = 28
    head(ws, 5, ['S. No.', 'Section of work', 'Amount', '% of works'])
    for i in range(20):
        r, s = 6 + i, 3 + i
        cell(ws, 'A%d' % r, '=IF(Project!D%d="","",%d)' % (s, i + 1), 'box', align='center')
        cell(ws, 'B%d' % r, '=IF(Project!D%d="","",Project!D%d)' % (s, s), 'box')
        cell(ws, 'C%d' % r, '=IF($B{r}="","",SUMIF(BOQ!$B$4:$B$203,$B{r},BOQ!$I$4:$I$203))'.format(r=r), 'calc', INR)
        cell(ws, 'D%d' % r, '=IF(OR($B{r}="",$C$27=0),"",N($C{r})/$C$27)'.format(r=r), 'calc', '0.0%')
    cell(ws, 'B26', 'Items without a section', 'box', size=9, color=MUTED)
    cell(ws, 'C26', '=ROUND(SUM(BOQ!$I$4:$I$203)-SUM(C6:C25),2)', 'calc', INR)
    rows = [('Cost of works (total of BOQ)', '=ROUND(SUM(C6:C26),2)'),
            ('="Add: overheads & profit @ "&TEXT(Project!B11,"0.0%")', '=ROUND(C27*N(Project!B11),2)'),
            ('="Add: contingency @ "&TEXT(Project!B12,"0.0%")', '=ROUND(C27*N(Project!B12),2)'),
            ('Total before GST', '=C27+C28+C29'),
            ('="Add: GST @ "&N(Project!B13)&"%"', '=ROUND(C30*N(Project!B13)/100,2)'),
            ('Round off', '=C33-C30-C31'),
            ('GRAND TOTAL', '=ROUND(C30+C31,0)')]
    for i, (k, f) in enumerate(rows, 27):
        big = k in ('Cost of works (total of BOQ)', 'Total before GST', 'GRAND TOTAL')
        cell(ws, 'B%d' % i, k, 'soft' if big else 'box', bold=big); cell(ws, 'C%d' % i, f, 'soft' if big else 'calc', INR, bold=big)
    ws.merge_cells('A35:D36'); cell(ws, 'A35', words_formula('C33'), bold=True, wrap=True)
    cell(ws, 'B37', 'Cost per sq ft of built-up area', 'box'); cell(ws, 'C37', '=IF(N(Project!B10)=0,"",ROUND(C33/Project!B10,0))', 'calc', INR0)
    ws.merge_cells('A39:D40')
    cell(ws, 'A39', 'This is an estimate based on the quantities and rates shown. Actual cost can change with site conditions, design changes and market rates.',
         wrap=True, size=9, color=MUTED)
    cell(ws, 'B43', '="Prepared by: "&Project!B6', size=10); cell(ws, 'C43', 'Signature', size=9, align='center')
    ws.print_area = 'A1:D44'
    protect(ws)
    save(wb, 'Construction-Estimate-BOQ.xlsx')

# ------------------------------------------------------------------ RA bill & measurement book
def ra_bill():
    wb = new_wb('con-ra-bill-excel')
    howto(wb, first=True)
    ct = sheet(wb, 'Contract', [40, 40, 3, 50], tab=ORANGE)
    it = sheet(wb, 'Items', [9, 54, 8, 12, 12, 15])
    mb = sheet(wb, 'MB', [12, 8, 9, 36, 7, 9, 9, 9, 11, 18], tab=ORANGE)
    bi = sheet(wb, 'Bill items', [8, 40, 7, 10, 11, 11, 11, 11, 14, 14, 15, 14], tab=ORANGE)
    bs = sheet(wb, 'Bill summary', [56, 20], tab=ORANGE)
    lists_sheet(wb, extra={5: UNITS})
    # Contract
    title(ct, 'Contract and this bill', 'Change "Bill No. (this bill)" each time you prepare a new RA bill.')
    rows = [('Name of work', 'Construction of G+2 commercial building, Phase 1', None), ('Employer / client', 'Gill Developers Pvt. Ltd.', None),
            ('Contractor', 'Sharma Builders & Contractors', None), ('Work order / agreement No.', 'WO/2026/031', None),
            ('Agreement date', TODAY - datetime.timedelta(days=75), DATE), ('Contract value (from Items)', '=SUM(Items!F4:F103)', INR),
            ('Bill No. (this bill)', 2, '0'), ('Bill date', TODAY, DATE), ('Work done from', TODAY - datetime.timedelta(days=30), DATE),
            ('Work done to', TODAY - datetime.timedelta(days=1), DATE)]
    for i, (k, v, f) in enumerate(rows, 3):
        label(ct, 'A%d' % i, k); cell(ct, 'B%d' % i, v, 'calc' if k.startswith('Contract value') else 'in', f, align='left')
    label(ct, 'A14', 'Deductions and taxes (check your agreement)')
    rows = [('GST % added to the bill', 18, '0'), ('Retention money % (on value of work)', 0.05, '0.0%'), ('Income-tax TDS % (on value of work)', 0.01, '0.0%'),
            ('GST TDS % (government clients; else 0)', 0, '0.0%'), ('Labour welfare cess % (on value of work)', 0.01, '0.0%'),
            ('Mobilisation advance recovered in this bill', 50000, INR), ('Other deductions in this bill (amount)', 0, INR), ('Other deductions: details', '', None)]
    for i, (k, v, f) in enumerate(rows, 15):
        label(ct, 'A%d' % i, k, bold=False); cell(ct, 'B%d' % i, v, 'in', f, align='left')
    dv(ct, 'B15', '"0,5,12,18"')
    for i, s in enumerate(['How the bill works:', '• Every MB row carries a bill number.', '• Rows with a lower number = up to previous bill.',
                           '• Rows with this bill\'s number = this bill.', '• Percentages: type 5% (or 0.05).', '• Deductions are worked out on the value of',
                           '  work in this bill, excluding GST.'], 3):
        cell(ct, 'D%d' % i, s, size=9, color=MUTED, bold=(i == 3))
    # Items
    title(it, 'Agreement items', 'As in the work order. Item numbers must match the MB.')
    head(it, 3, ['Item No.', 'Description of item', 'Unit', 'Agreement qty', 'Rate', 'Amount'])
    grid(it, 4, 103, 'ABCDEF', {'F': 'calc'}, {'D': QTY, 'E': INR, 'F': INR}, {'F': '=IF(OR($D{r}="",$E{r}=""),"",ROUND($D{r}*$E{r},2))'})
    for r in range(4, 104): it['B%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
    agr = [('1', 'Excavation for foundations in all soils, disposal within 50 m', 'cum', 420, 290),
           ('2', 'PCC M10 below foundations, 100 mm thick', 'cum', 28, 5900),
           ('3', 'RCC M25 in foundations, columns, beams and slabs incl. shuttering (steel separate)', 'cum', 310, 8200),
           ('4', 'TMT Fe 500D reinforcement, cut, bent, tied and placed', 'kg', 38000, 76),
           ('5', 'AAC block masonry 200 mm in thin-bed mortar', 'cum', 165, 6200),
           ('6', 'Internal cement plaster 12 mm, 1:4', 'sqm', 2600, 280)]
    for i, row in enumerate(agr, 4):
        for j, v in enumerate(row): it.cell(i, j + 1, v)
    dv(it, 'C4:C103', '=Lists!$E$1:$E$%d' % len(UNITS))
    it.freeze_panes = 'A4'
    # MB
    title(mb, 'Measurement book', 'One row per measurement. Quantity = Nos × L × B × D (empty cells skipped; negative Nos deducts).')
    head(mb, 3, ['Date', 'Bill No.', 'Item No.', 'Location / details', 'Nos', 'Length', 'Breadth', 'Depth / height', 'Quantity', 'Measured / checked by'])
    grid(mb, 4, 1003, 'ABCDEFGHJ', {}, {'A': DATE, 'B': '0', 'E': 'General', 'F': '0.000', 'G': '0.000', 'H': '0.000'})
    dims(mb, 4, 1003, 'E', 'H', 'I')
    dv(mb, 'C4:C1003', '=Items!$A$4:$A$103', 'Pick the agreement item')
    t = TODAY
    D = lambda n: t - datetime.timedelta(days=n)
    mbr = [(D(40), 1, '1', 'Footings F1 to F24, grid A-D', 24, 2.4, 2.4, 1.8), (D(40), 1, '1', 'Combined footing CF1', 2, 4.2, 2.6, 1.8),
           (D(38), 1, '2', 'PCC below footings F1-F24', 24, 2.4, 2.4, 0.1), (D(35), 1, '3', 'Footings F1-F24 (RCC)', 24, 2.2, 2.2, 0.45),
           (D(34), 1, '4', 'Footing steel as per bar bending schedule BBS-01', None, 5240, None, None, 'kg'),
           (D(20), 2, '3', 'Ground floor columns C1-C24', 24, 0.3, 0.45, 3.3), (D(15), 2, '3', 'First floor slab and beams, grid A-D/1-6', 1, 24, 18, 0.15),
           (D(15), 2, '4', 'Columns and slab steel, BBS-02 and BBS-03', None, 11850, None, None, 'kg'),
           (D(8), 2, '5', 'Ground floor walls, grid A and D', 2, 24, 0.2, 3.0), (D(8), 2, '5', 'Deduct openings, ground floor', -8, 1.5, 0.2, 2.1)]
    for i, row in enumerate(mbr, 4):
        for j, v in enumerate(row[:8]): mb.cell(i, j + 1, v)
        mb.cell(i, 10, 'Er. Kaur / ' + ('client rep.' if row[1] == 1 else 'R. Gill'))
    mb.freeze_panes = 'A4'
    mb.auto_filter.ref = 'A3:J1003'
    # Bill items
    ws = bi
    ws.page_setup.orientation = 'landscape'
    ws.merge_cells('A1:L1'); cell(ws, 'A1', '="RUNNING ACCOUNT BILL No. "&Contract!B9&"  ·  ABSTRACT OF QUANTITIES"', bold=True, size=14, color=ORANGE)
    ws.merge_cells('A2:L2'); cell(ws, 'A2', '=Contract!B3&"  ·  "&Contract!B6&"  ·  Contractor: "&Contract!B5', size=9, color=MUTED)
    head(ws, 3, ['Item No.', 'Description', 'Unit', 'Agreement qty', 'Rate', 'Qty up to previous bill', 'Qty this bill', 'Qty up to date',
                 'Amount up to previous', 'Amount this bill', 'Amount up to date', 'Check'])
    MBQ = 'SUMIFS(MB!$I$4:$I$1003,MB!$C$4:$C$1003,$A{r}&"",MB!$B$4:$B$1003,"%s"&Contract!$B$9)'
    for i in range(100):
        r, s = 4 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d&"")' % (s, s), 'box', align='center')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Items!B%d&"")' % (r, s), 'box', wrap=True)
        cell(ws, 'C%d' % r, '=IF($A%d="","",Items!C%d&"")' % (r, s), 'box', align='center')
        cell(ws, 'D%d' % r, '=IF($A%d="","",N(Items!D%d))' % (r, s), 'calc', QTY)
        cell(ws, 'E%d' % r, '=IF($A%d="","",N(Items!E%d))' % (r, s), 'calc', INR)
        cell(ws, 'F%d' % r, ('=IF($A{r}="","",' + MBQ % '<' + ')').format(r=r), 'calc', QTY)
        cell(ws, 'G%d' % r, ('=IF($A{r}="","",' + MBQ % '=' + ')').format(r=r), 'calc', QTY)
        cell(ws, 'H%d' % r, '=IF($A{r}="","",F{r}+G{r})'.format(r=r), 'calc', QTY)
        cell(ws, 'I%d' % r, '=IF($A{r}="","",ROUND(F{r}*E{r},2))'.format(r=r), 'calc', INR)
        cell(ws, 'J%d' % r, '=IF($A{r}="","",ROUND(G{r}*E{r},2))'.format(r=r), 'calc', INR)
        cell(ws, 'K%d' % r, '=IF($A{r}="","",I{r}+J{r})'.format(r=r), 'calc', INR)
        cell(ws, 'L%d' % r, '=IF($A{r}="","",IF(H{r}>D{r}+0.0005,"Excess "&TEXT(H{r}-D{r},"0.00"),"OK"))'.format(r=r), 'calc', align='center')
    ws.conditional_formatting.add('L4:L103', FormulaRule(formula=['LEFT($L4,6)="Excess"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    cell(ws, 'H104', 'Total', 'soft', bold=True)
    for col in 'IJK': cell(ws, '%s104' % col, '=SUM(%s4:%s103)' % (col, col), 'soft', INR, bold=True)
    ws.freeze_panes = 'C4'
    protect(ws)
    # Bill summary
    ws = bs
    ws.merge_cells('A1:B1'); cell(ws, 'A1', '="RUNNING ACCOUNT BILL No. "&Contract!B9', bold=True, size=15, color=ORANGE, align='center')
    ws.merge_cells('A2:B2'); cell(ws, 'A2', '=Contract!B3', bold=True, align='center', wrap=True)
    info = [('Employer', '=Contract!B4'), ('Contractor', '=Contract!B5'), ('Work order / agreement', '=Contract!B6&" dated "&TEXT(Contract!B7,"dd-mm-yyyy")'),
            ('Bill date', '=TEXT(Contract!B10,"dd-mm-yyyy")'), ('Work done during', '=TEXT(Contract!B11,"dd-mm-yyyy")&" to "&TEXT(Contract!B12,"dd-mm-yyyy")')]
    for i, (k, f) in enumerate(info, 3):
        cell(ws, 'A%d' % i, k, 'box'); cell(ws, 'B%d' % i, f, 'box', align='right')
    pct = lambda ref: 'TEXT(%s,"0.0%%")' % ref
    lines = [('Gross value of work done up to date', "='Bill items'!K104", 1),
             ('Less: value up to previous bill', "='Bill items'!I104", 0),
             ('Value of work done in this bill', '=B10-B11', 1),
             ('="Add: GST @ "&Contract!B15&"%"', '=ROUND(B12*N(Contract!B15)/100,2)', 0),
             ('Bill amount including GST', '=B12+B13', 1),
             ('="Less: retention money @ "&' + pct('Contract!B16'), '=ROUND(B12*N(Contract!B16),2)', 0),
             ('="Less: income-tax TDS @ "&' + pct('Contract!B17'), '=ROUND(B12*N(Contract!B17),0)', 0),
             ('="Less: GST TDS @ "&' + pct('Contract!B18'), '=ROUND(B12*N(Contract!B18),0)', 0),
             ('="Less: labour welfare cess @ "&' + pct('Contract!B19'), '=ROUND(B12*N(Contract!B19),0)', 0),
             ('Less: mobilisation advance recovery', '=N(Contract!B20)', 0),
             ('="Less: other deductions"&IF(Contract!B22="",""," ("&Contract!B22&")")', '=N(Contract!B21)', 0),
             ('Total deductions', '=SUM(B15:B20)', 1),
             ('NET AMOUNT PAYABLE', '=ROUND(B14-B21,0)', 1)]
    for i, (k, f, b) in enumerate(lines, 10):
        cell(ws, 'A%d' % i, k, 'soft' if b else 'box', bold=bool(b)); cell(ws, 'B%d' % i, f, 'soft' if b else 'calc', INR, bold=bool(b))
    ws.merge_cells('A24:B25'); cell(ws, 'A24', words_formula('B22'), bold=True, wrap=True)
    cell(ws, 'A27', 'Retention money held up to date (all bills)', 'box'); cell(ws, 'B27', '=ROUND(B10*N(Contract!B16),2)', 'calc', INR)
    cell(ws, 'A28', 'Value of work done up to date as % of contract value', 'box')
    cell(ws, 'B28', '=IF(N(Contract!B8)=0,"",B10/Contract!B8)', 'calc', '0.0%')
    ws.merge_cells('A30:B31')
    cell(ws, 'A30', 'Certified that the measurements of the work billed above were taken jointly and recorded in the measurement book, and that the work has been done as per the agreement and specifications.',
         wrap=True, size=9, color=MUTED)
    cell(ws, 'A35', 'Contractor ____________________', size=10); cell(ws, 'B35', 'Site engineer ______________', size=10)
    cell(ws, 'A38', 'Checked by (QS / billing) ____________________', size=10); cell(ws, 'B38', 'Approved for payment ________', size=10)
    ws.print_area = 'A1:B39'
    protect(ws)
    save(wb, 'Construction-RA-Bill-and-MB.xlsx')

# ------------------------------------------------------------------ Labour attendance & wages
def labour_wages():
    wb = new_wb('con-labour-wage-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [34, 30, 3, 22])
    wk = sheet(wb, 'Workers', [8, 24, 18, 12, 15, 26])
    at = sheet(wb, 'Attendance', [7, 20, 13] + [3.8] * 31 + [5, 5, 5, 7, 7], tab=ORANGE)
    ad = sheet(wb, 'Advances', [12, 9, 22, 13, 30])
    wg = sheet(wb, 'Wages', [7, 22, 13, 10, 8, 8, 12, 11, 12, 11, 13, 18], tab=ORANGE)
    # Settings
    title(st, 'Settings', 'Site, wage period and overtime rules. Trades list on the right.')
    start = TODAY - datetime.timedelta(days=TODAY.weekday())   # Monday of this week
    rows = [('Site / project', 'Gill Developers, G+2 commercial building', None), ('Contractor / labour agency', 'Sharma Builders & Contractors', None),
            ('Wage period starts', start, DATE), ('Days in the period (7 = weekly, up to 31)', 7, '0'),
            ('Hours in a working day', 8, '0'), ('Overtime rate (× ordinary hourly rate)', 2, '0.0')]
    for i, (k, v, f) in enumerate(rows, 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', f, align='left')
    label(st, 'A9', 'Period ends'); cell(st, 'B9', '=B5+MIN(31,MAX(1,N(B6)))-1', 'calc', DATE, align='left')
    cell(st, 'A11', 'Daily wages must not be below the minimum wage notified for each trade in your state.', size=9, color=MUTED)
    head(st, 2, ['Trades (edit freely)'], col=4)
    TR = ['Mason', 'Helper / beldar', 'Carpenter', 'Bar bender', 'Electrician', 'Plumber', 'Painter', 'Tile fixer', 'Welder', 'Machine operator', 'Watchman']
    for i in range(3, 23): cell(st, 'D%d' % i, TR[i - 3] if i - 3 < len(TR) else None, 'in')
    # Workers
    title(wk, 'Workers', 'One row per worker. The code is used on the Advances sheet.')
    head(wk, 3, ['Code', 'Name', 'Trade', 'Daily wage', 'Phone', 'Bank account / UPI'])
    grid(wk, 4, 103, 'ABCDEF', {}, {'D': INR})
    dv(wk, 'C4:C103', '=Settings!$D$3:$D$22')
    ppl = [('W01', 'Ramesh Kumar', 'Mason', 900), ('W02', 'Suresh Yadav', 'Mason', 900), ('W03', 'Bhola Ram', 'Helper / beldar', 550),
           ('W04', 'Sunita Devi', 'Helper / beldar', 550), ('W05', 'Manoj Paswan', 'Helper / beldar', 550), ('W06', 'Gurmeet Singh', 'Carpenter', 850),
           ('W07', 'Akhtar Ali', 'Bar bender', 800), ('W08', 'Dinesh Sahni', 'Electrician', 850)]
    for i, row in enumerate(ppl, 4):
        for j, v in enumerate(row): wk.cell(i, j + 1, v)
    wk.freeze_panes = 'A4'
    # Attendance
    ws = at
    ws.page_setup.orientation = 'landscape'
    ws['A1'] = 'Labour attendance'; ws['A1'].font = Font(size=16, bold=True, color=BLUE)
    cell(ws, 'A2', '=Settings!B3&"  ·  "&TEXT(Settings!B5,"dd-mm-yyyy")&" to "&TEXT(Settings!B9,"dd-mm-yyyy")&"  ·  P present, H half day, A absent"', size=10, color=MUTED)
    first, last = 4, 34   # columns D .. AH
    for n in range(31):
        col = L(first + n)
        cell(ws, '%s4' % col, '=IF(%d<MIN(31,MAX(1,N(Settings!$B$6))),Settings!$B$5+%d,"")' % (n, n), 'calc', 'd', align='center', bold=True, size=9)
        cell(ws, '%s5' % col, '=IF(%s4="","",LEFT(TEXT(%s4,"ddd"),2))' % (col, col), 'calc', align='center', size=8)
    head(ws, 4, ['Code', 'Name', 'Trade'])
    for c in 'ABC': ws.merge_cells('%s4:%s5' % (c, c))
    for k, lab_ in enumerate(['P', 'H', 'A', 'Days worked', 'OT hours']):
        col = L(35 + k)
        c = ws['%s4' % col]; c.value = lab_; c.font, c.fill, c.border = Font(bold=True, color='FFFFFF', size=9), F_HEAD, BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); ws.merge_cells('%s4:%s5' % (col, col))
    ws.row_dimensions[4].height = 30
    G1, G2 = L(first), L(last)
    for i in range(100):
        r, s = 6 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Workers!A%d="","",Workers!A%d&"")' % (s, s), 'box', size=9)
        cell(ws, 'B%d' % r, '=IF($A%d="","",Workers!B%d&"")' % (r, s), 'box', size=9)
        cell(ws, 'C%d' % r, '=IF($A%d="","",Workers!C%d&"")' % (r, s), 'box', size=8)
        for n in range(31): cell(ws, '%s%d' % (L(first + n), r), None, 'in', align='center')
        cell(ws, 'AI%d' % r, '=IF($A{r}="","",COUNTIF({a}{r}:{b}{r},"P"))'.format(r=r, a=G1, b=G2), 'calc', align='center')
        cell(ws, 'AJ%d' % r, '=IF($A{r}="","",COUNTIF({a}{r}:{b}{r},"H"))'.format(r=r, a=G1, b=G2), 'calc', align='center')
        cell(ws, 'AK%d' % r, '=IF($A{r}="","",COUNTIF({a}{r}:{b}{r},"A"))'.format(r=r, a=G1, b=G2), 'calc', align='center')
        cell(ws, 'AL%d' % r, '=IF($A{r}="","",AI{r}+AJ{r}/2)'.format(r=r), 'calc', '0.0', align='center')
        cell(ws, 'AM%d' % r, None, 'in', 'General', align='center')
    dv(ws, '%s6:%s105' % (G1, G2), '"P,H,A"')
    ws.conditional_formatting.add('%s4:%s105' % (G1, G2), FormulaRule(formula=['%s$4=""' % G1], fill=PatternFill('solid', fgColor='E2E8F0')))
    ws.conditional_formatting.add('%s6:%s105' % (G1, G2), CellIsRule(operator='equal', formula=['"A"'], font=Font(bold=True, color='B91C1C')))
    cell(ws, 'B106', 'Headcount on site', 'soft', bold=True); ws.merge_cells('B106:C106')
    for n in range(31):
        col = L(first + n)
        cell(ws, '%s106' % col, '=IF({c}$4="","",COUNTIF({c}6:{c}105,"P")+COUNTIF({c}6:{c}105,"H"))'.format(c=col), 'soft', align='center', size=9)
    # sample marks: every worker present most days of the first week, a few H/A, some overtime
    marks = ['PPPPPPA', 'PPPPHPA', 'PPAPPPA', 'PPPPPPA', 'PHPPPAA', 'PPPPPPA', 'PPPPPPP', 'APPPPPA']
    ot = [6, 4, 0, 0, 2, 8, 10, 3]
    for i, m in enumerate(marks):
        for n, ch in enumerate(m): ws['%s%d' % (L(first + n), 6 + i)] = ch
        if ot[i]: ws['AM%d' % (6 + i)] = ot[i]
    ws.freeze_panes = 'D6'
    protect(ws)
    # Advances
    title(ad, 'Advances to workers', 'Advances dated inside the wage period are deducted on the Wages sheet.')
    head(ad, 3, ['Date', 'Code', 'Name', 'Amount', 'Note'])
    grid(ad, 4, 503, 'ABCDE', {'C': 'calc'}, {'A': DATE, 'D': INR},
         {'C': '=IF($B{r}="","",IFERROR(INDEX(Workers!$B$4:$B$103,MATCH($B{r},Workers!$A$4:$A$103,0)),"Code not found"))'})
    dv(ad, 'B4:B503', '=Workers!$A$4:$A$103')
    for i, row in enumerate([(start + datetime.timedelta(days=1), 'W03', None, 500, 'Ration'), (start + datetime.timedelta(days=3), 'W07', None, 1000, 'Going home'),
                             (start + datetime.timedelta(days=4), 'W01', None, 800, '')], 4):
        ad.cell(i, 1, row[0]); ad.cell(i, 2, row[1]); ad.cell(i, 4, row[3]); ad.cell(i, 5, row[4])
    ad.freeze_panes = 'A4'
    # Wages
    ws = wg
    ws.page_setup.orientation = 'landscape'
    ws.merge_cells('A1:L1'); cell(ws, 'A1', '="WAGE SHEET  ·  "&TEXT(Settings!B5,"dd-mm-yyyy")&" to "&TEXT(Settings!B9,"dd-mm-yyyy")', bold=True, size=14, color=ORANGE)
    ws.merge_cells('A2:L2'); cell(ws, 'A2', '=Settings!B3&"  ·  "&Settings!B4', size=9, color=MUTED)
    head(ws, 4, ['Code', 'Name', 'Trade', 'Daily wage', 'Days worked', 'OT hours', 'Wages', 'OT amount', 'Gross', 'Advance', 'Net payable', 'Signature / thumb'])
    cell(ws, 'F3', 'Totals', 'soft', bold=True, align='right')
    for col in 'GHIJK': cell(ws, '%s3' % col, '=SUM(%s5:%s104)' % (col, col), 'soft', INR0, bold=True)
    for i in range(100):
        r, s, a = 5 + i, 4 + i, 6 + i
        cell(ws, 'A%d' % r, '=IF(Workers!A%d="","",Workers!A%d&"")' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Workers!B%d&"")' % (r, s), 'box')
        cell(ws, 'C%d' % r, '=IF($A%d="","",Workers!C%d&"")' % (r, s), 'box', size=9)
        cell(ws, 'D%d' % r, '=IF($A%d="","",N(Workers!D%d))' % (r, s), 'calc', INR0)
        cell(ws, 'E%d' % r, '=IF($A%d="","",N(Attendance!AL%d))' % (r, a), 'calc', '0.0', align='center')
        cell(ws, 'F%d' % r, '=IF($A%d="","",N(Attendance!AM%d))' % (r, a), 'calc', 'General', align='center')
        cell(ws, 'G%d' % r, '=IF($A{r}="","",ROUND(D{r}*E{r},0))'.format(r=r), 'calc', INR0)
        cell(ws, 'H%d' % r, '=IF($A{r}="","",ROUND(F{r}*D{r}/MAX(1,N(Settings!$B$7))*N(Settings!$B$8),0))'.format(r=r), 'calc', INR0)
        cell(ws, 'I%d' % r, '=IF($A{r}="","",G{r}+H{r})'.format(r=r), 'calc', INR0)
        cell(ws, 'J%d' % r, '=IF($A{r}="","",SUMIFS(Advances!$D$4:$D$503,Advances!$B$4:$B$503,$A{r},Advances!$A$4:$A$503,">="&Settings!$B$5,Advances!$A$4:$A$503,"<="&Settings!$B$9))'.format(r=r), 'calc', INR0)
        cell(ws, 'K%d' % r, '=IF($A{r}="","",I{r}-J{r})'.format(r=r), 'soft', INR0, bold=True)
        cell(ws, 'L%d' % r, None, 'box')
        ws.row_dimensions[r].height = 22
    ws.conditional_formatting.add('K5:K104', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'C5'
    protect(ws)
    save(wb, 'Construction-Labour-Wage-Register.xlsx')

# ------------------------------------------------------------------ Material register
def material_register():
    wb = new_wb('con-material-register-excel')
    howto(wb, first=True)
    mt = sheet(wb, 'Materials', [30, 9, 13, 13, 14])
    sp = sheet(wb, 'Suppliers', [30, 16, 20])
    iw = sheet(wb, 'Inward', [12, 28, 26, 16, 14, 11, 12, 14, 18], tab=ORANGE)
    iss = sheet(wb, 'Issue', [12, 28, 11, 26, 30, 18], tab=ORANGE)
    sk = sheet(wb, 'Stock', [30, 9, 11, 11, 11, 11, 13, 15, 10], tab=ORANGE)
    ss = sheet(wb, 'Supplier summary', [30, 12, 16])
    title(mt, 'Materials', 'Opening stock = what is in the store on the day you start this register.')
    head(mt, 3, ['Material', 'Unit', 'Opening stock', 'Reorder level', 'Standard rate'])
    grid(mt, 4, 103, 'ABCDE', {}, {'C': 'General', 'D': 'General', 'E': INR})
    mats = [('Cement OPC 53 grade', 'bag', 120, 100, 380), ('Cement PPC', 'bag', 40, 50, 360), ('TMT bar 8 mm Fe 500D', 'kg', 850, 500, 64),
            ('TMT bar 12 mm Fe 500D', 'kg', 1600, 1000, 62), ('TMT bar 16 mm Fe 500D', 'kg', 900, 800, 62), ('Coarse sand', 'cft', 400, 300, 55),
            ('Aggregate 20 mm', 'cft', 500, 300, 48), ('Red bricks (first class)', 'nos', 6000, 4000, 8.5), ('Binding wire', 'kg', 40, 25, 85),
            ('Shuttering plywood 12 mm', 'sheet', 30, 10, 1450), ('Waterproofing admixture', 'litre', 20, 10, 180)]
    for i, row in enumerate(mats, 4):
        for j, v in enumerate(row): mt.cell(i, j + 1, v)
    mt.freeze_panes = 'A4'
    title(sp, 'Suppliers', 'One row per supplier. Used in the Inward register drop-down.'); head(sp, 3, ['Supplier', 'Phone', 'GSTIN']); grid(sp, 4, 53, 'ABC', {})
    for i, row in enumerate([('Punjab Cement Agencies', '98140 12345', '03AAAPC1234A1Z1'), ('Ludhiana Steel Traders', '98150 22334', '03AAALS5678B1Z2'),
                             ('Sutlej Sand & Grit Suppliers', '98760 44556', ''), ('Bhatti Bricks Kiln', '98720 77889', '')], 4):
        for j, v in enumerate(row): sp.cell(i, j + 1, v)
    title(iw, 'Material received (inward)', 'Record the quantity actually received and checked, not just the challan quantity.')
    head(iw, 3, ['Date', 'Material', 'Supplier', 'Challan / bill No.', 'Vehicle No.', 'Quantity', 'Rate', 'Amount', 'Received by / remarks'])
    grid(iw, 4, 1003, 'ABCDEFGHI', {'H': 'calc'}, {'A': DATE, 'F': 'General', 'G': INR, 'H': INR}, {'H': '=IF(OR($F{r}="",$G{r}=""),"",ROUND($F{r}*$G{r},2))'})
    dv(iw, 'B4:B1003', '=Materials!$A$4:$A$103'); dv(iw, 'C4:C1003', '=Suppliers!$A$4:$A$53')
    d = MONTH1
    for i, row in enumerate([(d, 'Cement OPC 53 grade', 'Punjab Cement Agencies', 'CH-4471', 'PB10 GH 2231', 200, 375, None, 'Raju (store)'),
                             (d + datetime.timedelta(days=1), 'TMT bar 12 mm Fe 500D', 'Ludhiana Steel Traders', 'INV-889', 'PB10 FT 9087', 2480, 61.5, None, 'Weighed 2,480 kg (challan 2,500)'),
                             (d + datetime.timedelta(days=2), 'Coarse sand', 'Sutlej Sand & Grit Suppliers', '117', 'PB13 K 4410', 600, 54, None, 'Raju (store)'),
                             (d + datetime.timedelta(days=3), 'Red bricks (first class)', 'Bhatti Bricks Kiln', 'BK-302', 'Tractor PB10 7781', 8000, 8.4, None, '150 broken, not accepted')], 4):
        for j, v in enumerate(row):
            if v is not None: iw.cell(i, j + 1, v)
    iw.freeze_panes = 'A4'
    title(iss, 'Material issued', 'Every time material leaves the store.')
    head(iss, 3, ['Date', 'Material', 'Quantity', 'Issued to (gang / sub-contractor)', 'Work / location', 'Issued by / remarks'])
    grid(iss, 4, 1003, 'ABCDEF', {}, {'A': DATE, 'C': 'General'})
    dv(iss, 'B4:B1003', '=Materials!$A$4:$A$103')
    for i, row in enumerate([(d, 'Cement OPC 53 grade', 85, 'Mason gang (Ramesh)', 'Column casting, ground floor C1-C12', 'Raju'),
                             (d + datetime.timedelta(days=1), 'TMT bar 12 mm Fe 500D', 1350, 'Bar bender (Akhtar)', 'First floor beams B1-B8', 'Raju'),
                             (d + datetime.timedelta(days=2), 'Coarse sand', 260, 'Mason gang (Ramesh)', 'Brickwork, ground floor', 'Raju'),
                             (d + datetime.timedelta(days=3), 'Red bricks (first class)', 7200, 'Mason gang (Suresh)', 'Brickwork, ground floor grid A-D', 'Raju'),
                             (d + datetime.timedelta(days=3), 'Cement OPC 53 grade', 110, 'Mason gang (Suresh)', 'Brickwork mortar 1:6', 'Raju')], 4):
        for j, v in enumerate(row): iss.cell(i, j + 1, v)
    iss.freeze_panes = 'A4'
    ws = sk
    title(ws, 'Stock in the site store', 'Calculates itself from the Materials, Inward and Issue sheets.')
    label(ws, 'G2', 'Stock value'); cell(ws, 'H2', '=SUM(H5:H104)', 'soft', INR, bold=True)
    head(ws, 4, ['Material', 'Unit', 'Opening', 'Received', 'Issued', 'Balance', 'Average rate', 'Value', 'Status'])
    for i in range(100):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Materials!A%d="","",Materials!A%d&"")' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Materials!B%d&"")' % (r, s), 'box', align='center')
        cell(ws, 'C%d' % r, '=IF($A%d="","",N(Materials!C%d))' % (r, s), 'calc', 'General')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIF(Inward!$B$4:$B$1003,$A{r},Inward!$F$4:$F$1003))'.format(r=r), 'calc', 'General')
        cell(ws, 'E%d' % r, '=IF($A{r}="","",SUMIF(Issue!$B$4:$B$1003,$A{r},Issue!$C$4:$C$1003))'.format(r=r), 'calc', 'General')
        cell(ws, 'F%d' % r, '=IF($A{r}="","",C{r}+D{r}-E{r})'.format(r=r), 'calc', 'General')
        cell(ws, 'G%d' % r, ('=IF($A{r}="","",IF(D{r}>0,ROUND(SUMIF(Inward!$B$4:$B$1003,$A{r},Inward!$H$4:$H$1003)/D{r},2),N(Materials!E{s})))').format(r=r, s=s), 'calc', INR)
        cell(ws, 'H%d' % r, '=IF($A{r}="","",ROUND(MAX(0,F{r})*G{r},2))'.format(r=r), 'calc', INR)
        cell(ws, 'I%d' % r, '=IF($A{r}="","",IF(F{r}<0,"Short",IF(F{r}<=N(Materials!D{s}),"Reorder","OK")))'.format(r=r, s=s), 'calc', align='center')
    ws.conditional_formatting.add('I5:I104', CellIsRule(operator='equal', formula=['"Reorder"'], fill=PatternFill('solid', fgColor='FEF3C7'), font=Font(bold=True, color='B45309')))
    ws.conditional_formatting.add('I5:I104', CellIsRule(operator='equal', formula=['"Short"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'B5'
    protect(ws)
    title(ss, 'Deliveries by supplier', 'Compare with each supplier\'s bills before paying.')
    head(ss, 3, ['Supplier', 'Deliveries', 'Total value'])
    for i in range(50):
        r, s = 4 + i, 4 + i
        cell(ss, 'A%d' % r, '=IF(Suppliers!A%d="","",Suppliers!A%d&"")' % (s, s), 'box')
        cell(ss, 'B%d' % r, '=IF($A{r}="","",COUNTIF(Inward!$C$4:$C$1003,$A{r}))'.format(r=r), 'calc', align='center')
        cell(ss, 'C%d' % r, '=IF($A{r}="","",SUMIF(Inward!$C$4:$C$1003,$A{r},Inward!$H$4:$H$1003))'.format(r=r), 'calc', INR)
    cell(ss, 'A54', 'Total', 'soft', bold=True); cell(ss, 'C54', '=SUM(C4:C53)', 'soft', INR, bold=True)
    protect(ss)
    save(wb, 'Construction-Material-Register.xlsx')

# ------------------------------------------------------------------ Word: daily site diary
def site_diary():
    d = doc_new('con-site-diary-word')
    para(d, 'DAILY SITE DIARY / PROGRESS REPORT', bold=True, size=16, color=ORANGE, align='c', space=2)
    para(d, '[Contractor / company name]', bold=True, align='c', space=8)
    kv(d, [('Project / name of work', '[ ]'), ('Client', '[ ]'), ('Date', '[DD-MM-YYYY]'), ('Day No. (from start)', '[ ]'),
           ('Weather: morning', '[Clear / cloudy / rain]'), ('Weather: afternoon', '[ ]'), ('Working hours', '[08:00 to 17:00]'), ('Report by', '[Site engineer]')])
    para(d, '1. Manpower on site', bold=True, color=ORANGE)
    table(d, ['Trade', 'Agency / sub-contractor', 'Nos', 'Remarks'],
          [['Mason', '[ ]', '[ ]', ''], ['Helper / beldar', '[ ]', '[ ]', ''], ['Carpenter (shuttering)', '[ ]', '[ ]', ''], ['Bar bender', '[ ]', '[ ]', ''],
           ['Electrician / plumber', '[ ]', '[ ]', ''], ['Total on site', '', '[ ]', '']], [5, 6, 2, 4])
    para(d, '2. Machinery and equipment', bold=True, color=ORANGE)
    table(d, ['Machine', 'Nos', 'Hours worked', 'Hours idle / reason'], [['[Concrete mixer]', '[ ]', '[ ]', '[ ]'], ['[Vibrator]', '[ ]', '[ ]', '[ ]']], [6, 2, 3, 6], fill_rows=1)
    para(d, '3. Work done today', bold=True, color=ORANGE)
    table(d, ['#', 'Activity', 'Location (floor / grid)', 'Quantity', 'Unit'],
          [['1', '[e.g. RCC slab casting M25]', '[First floor, grid A-D/1-6]', '[ ]', '[cum]'], ['2', '', '', '', '']], [1, 6.5, 5, 2, 2], fill_rows=3)
    para(d, '4. Materials received and tests', bold=True, color=ORANGE)
    table(d, ['Material', 'Quantity', 'Supplier / challan No.', 'Remarks'], [['[Cement OPC 53]', '[ ] bags', '[ ]', '']], [5, 3, 5, 4], fill_rows=2)
    para(d, 'Tests done today: [slump test ___ mm; ___ cubes cast, ID ___; other]', size=10)
    para(d, '5. Instructions received and visitors', bold=True, color=ORANGE)
    table(d, ['Time', 'From (name, designation)', 'Instruction / purpose of visit'], [['[ ]', '[ ]', '[ ]']], [2, 5, 10], fill_rows=1)
    para(d, '6. Hindrances and delays', bold=True, color=ORANGE)
    table(d, ['Hindrance (rain, drawings, power, material, client change)', 'From – to', 'Hours lost', 'Work affected'], [['[ ]', '[ ]', '[ ]', '[ ]']], [7, 3, 2, 5], fill_rows=1)
    para(d, '7. Safety', bold=True, color=ORANGE)
    para(d, 'Toolbox talk held: [Yes / No, topic]   ·   PPE checked: [Yes / No]   ·   Incidents or near misses: [None / details]', size=10)
    para(d, '8. Plan for tomorrow', bold=True, color=ORANGE)
    para(d, '[Activities, labour and material needed, inspections to call]', space=16)
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = '\n\n____________________', '\n\n____________________'
    t.rows[1].cells[0].text, t.rows[1].cells[1].text = 'Site engineer / supervisor', 'Client\'s representative (seen)'
    dsave(d, 'Construction-Daily-Site-Diary.docx')

# ------------------------------------------------------------------ Word: work order / agreement draft
def work_order():
    d = doc_new('con-work-order-word')
    letterhead(d, '[Employer / Company Name]', '[Address] · [City, State PIN] · Phone [ ] · [email]')
    para(d, '[Starting draft: replace every [bracketed] item, delete clauses that do not apply, and have the final text checked by your lawyer or adviser before signing. Delete this note.]',
         size=9, italic=True, color='B45309', space=8)
    para(d, 'WORK ORDER AND CONTRACT AGREEMENT', bold=True, size=15, align='c', space=6)
    kv(d, [('Work order No.', '[WO/2026/001]'), ('Date', '[DD-MM-YYYY]'), ('Employer', '[Name, address]'), ('Contractor', '[Name, address, GSTIN, PAN]'),
           ('Site', '[Full site address]'), ('Name of work', '[e.g. Civil and structural work of G+2 residence]')])
    para(d, 'This agreement is made on [date] between [Employer name] ("the Employer") and [Contractor name] ("the Contractor"). The Employer gives and the Contractor accepts the work described below on the following terms.', space=8)
    def clause(head_, lines, style='List Bullet'):
        para(d, head_, bold=True, color=BLUE, space=2)
        for s in lines: d.add_paragraph(s, style=style)
    clause('1. Scope of work', ['The Contractor shall carry out [describe the work: e.g. excavation, foundations, RCC frame, brickwork and plaster] as per the approved drawings, the schedule of rates below and the specifications in Annexure A.',
                                'Work not listed in the schedule is outside the scope and will be done only on a written variation order (clause 9).'])
    clause('2. Materials', ['Supplied by the Employer: [cement, steel, ...]. The Contractor shall store them safely, account for them in a material register and return or pay for unexplained wastage beyond [ ]%.',
                            'Supplied by the Contractor: [all other materials, tools, shuttering, scaffolding, water for curing]. Materials shall be of the approved make and quality.'])
    para(d, '3. Schedule of rates', bold=True, color=BLUE, space=2)
    table(d, ['Item', 'Description of item', 'Unit', 'Est. qty', 'Rate (₹)', 'Amount (₹)'],
          [['1', '[Earthwork in excavation ...]', '[cum]', '[ ]', '[ ]', '[ ]'], ['2', '[RCC M25 ...]', '[cum]', '[ ]', '[ ]', '[ ]']], [1.2, 7.5, 1.5, 1.8, 2, 2.5], fill_rows=4)
    para(d, 'Quantities are estimates. Payment shall be for the actual quantities measured jointly at the rates above. Rates are [inclusive / exclusive] of GST and include all labour, tools, transport, taxes (other than GST), overheads and profit.', size=10)
    clause('4. Contract value and time', ['Estimated contract value: ₹ [ ] (Rupees [ ] only) [plus GST].',
                                          'Start date: [DD-MM-YYYY]. Completion period: [ ] months from the start date, that is by [DD-MM-YYYY].',
                                          'The Contractor shall follow the work programme in Annexure B and report progress [weekly].'])
    clause('5. Measurement and payment', ['Work shall be measured jointly and recorded in a measurement book signed by both parties.',
                                          'The Contractor may submit a running account (RA) bill every [month] for work done. The Employer shall pay within [15] days of a correct bill, after the deductions below.',
                                          'Retention money of [5]% of each bill shall be held and released after the defects liability period, against any defects found.',
                                          'Taxes, TDS and statutory deductions shall be made as required by law. Mobilisation advance, if any, of ₹ [ ] shall be recovered in [ ] equal instalments from the RA bills.',
                                          'The final bill shall be paid within [30] days of completion and joint final measurement.'])
    clause('6. Labour, safety and law', ['The Contractor is the employer of its workers and shall pay wages not lower than the notified minimum wages, maintain attendance and wage registers, and comply with labour laws that apply (including registration and welfare contributions where required).',
                                         'The Contractor shall provide helmets, safety shoes, harnesses for work at height and other protective equipment, keep the site tidy, and not engage child labour.',
                                         'The Contractor shall take insurance for its workers [and the works] as required by law and by the Employer.'])
    clause('7. Quality', ['Work shall follow the drawings, the specifications and good practice. Concrete cubes and other tests shall be done as directed by the Employer\'s engineer at the [Contractor\'s / Employer\'s] cost.',
                          'Work found defective shall be removed and redone by the Contractor at its own cost.'])
    clause('8. Delays', ['If the work is delayed for reasons within the Contractor\'s control, the Contractor shall pay compensation of [0.5]% of the contract value for every week of delay, up to a maximum of [5]%.',
                         'Delays caused by the Employer (drawings, materials, payments, site access) or by events beyond either party\'s control shall be recorded in the site diary, and the completion date shall be extended accordingly.'])
    clause('9. Variations and extra items', ['No change in scope or extra item shall be carried out without a written order from the Employer. Extra items shall be paid at rates agreed in writing before the work is done.'])
    clause('10. Defects liability', ['The Contractor shall repair, at its own cost, any defect that appears within [12] months after completion and is due to poor workmanship or materials supplied by the Contractor.'])
    clause('11. Termination', ['Either party may end this agreement by [15] days\' written notice if the other party seriously breaches it and does not remedy the breach within that notice. Work done until then shall be measured and settled.'])
    clause('12. Disputes', ['The parties shall first try to settle any dispute by discussion. Failing that, [it shall be referred to arbitration by a sole arbitrator under the Arbitration and Conciliation Act, 1996 / the courts at [city] shall have jurisdiction].'])
    clause('13. Special conditions', ['[Any other condition: working hours, storage area, water and electricity, watch and ward, cleaning, hand-over documents.]'])
    para(d, 'Annexures: A. Specifications   B. Work programme   C. Drawings list', size=10, space=12)
    para(d, 'Accepted and signed by both parties on the date above.', space=24)
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = '\n\n____________________\nFor the Employer\nName, designation', '\n\n____________________\nFor the Contractor\nName, designation'
    t.rows[1].cells[0].text, t.rows[1].cells[1].text = '\n\nWitness 1: ____________________\nName, address', '\n\nWitness 2: ____________________\nName, address'
    dsave(d, 'Construction-Work-Order-Draft.docx')

MAKERS = [estimate_boq, ra_bill, labour_wages, material_register, site_diary, work_order]
