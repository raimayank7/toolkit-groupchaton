"""Purchase & Warehouse toolkit files: PO register, GRN + QC log, supplier rating, multi-location inventory,
   ABC analysis + FIFO/average valuation (Excel) and the material requisition slip (Word).
   Global toolkit: money cells use a plain number format and the currency code typed on each workbook's Settings sheet."""
from make_files import *

AMT = '#,##0.00'
AMT0 = '#,##0'
QTY = '#,##0.##'
D2 = 'dd-mmm-yyyy'          # 02-Nov-2026: reads the same everywhere
RED = PatternFill('solid', fgColor='FECACA')
GREEN = PatternFill('solid', fgColor='BBF7D0')
AMBER = PatternFill('solid', fgColor='FEF3C7')
DAY = lambda n: TODAY + datetime.timedelta(days=n)

def settings(wb, rows, note=None, widths=(30, 26)):
    """Settings sheet: rows = [(label, value, fmt, dv-source or None)], values from B3 down."""
    ws = sheet(wb, 'Settings', list(widths))
    title(ws, 'Settings', note or 'Yellow cells: type or pick. Other sheets read these values.')
    for i, (k, v, fmt, src) in enumerate(rows, 3):
        label(ws, 'A%d' % i, k); cell(ws, 'B%d' % i, v, 'in', fmt, align='left')
        if src: dv(ws, 'B%d' % i, src)
    return ws

def cur(text):
    """Column heading with the currency code from Settings!B4, e.g. Rate (USD)."""
    return '="%s ("&Settings!$B$4&")"' % text

def hidden_lists(wb, cols):
    ws = wb.create_sheet('Lists')
    for col, vals in cols.items():
        for i, v in enumerate(vals, 1): ws.cell(i, col, v)
    ws.sheet_state = 'hidden'
    return ws

def rows_in(ws, r0, data, col=1):
    for i, row in enumerate(data, r0):
        for j, v in enumerate(row):
            if v is not None: ws.cell(i, col + j, v)

# ------------------------------------------------------------------ 1. Purchase order register
def po_register():
    wb = new_wb('pw-po-register-excel')
    howto(wb, first=True)
    settings(wb, [('Company name', 'Apex Engineering Works', None, None), ('Currency code', 'INR', None, None),
                  ('Position as on', '=TODAY()', D2, None)],
             'Leave "Position as on" as =TODAY(), or type a date to see the position on that day.')
    su = sheet(wb, 'Suppliers', [30, 22, 16, 12])
    od = sheet(wb, 'Orders', [11, 12, 24, 30, 7, 10, 11, 7, 12, 9, 13, 11, 13, 10, 10, 13, 14, 9], tab=BLUE)
    rc = sheet(wb, 'Receipts', [12, 11, 30, 10, 16, 26, 16], tab=BLUE)
    pd = sheet(wb, 'Pending by supplier', [30, 15, 15, 11, 11, 13])
    hidden_lists(wb, {1: ['Yes']})
    title(su, 'Suppliers', 'The supplier drop-down on the Orders sheet reads this list.')
    head(su, 3, ['Supplier', 'Contact person', 'Phone', 'Lead time (days)'])
    grid(su, 4, 203, 'ABCD', {})
    rows_in(su, 4, [('Bharat Steel Traders', 'Mr. Anil Bansal', '98140 12345', 10), ('Precision Bearings Co.', 'Ms. Ritu Kapoor', '98720 23456', 21),
                    ('Shree Packaging Industries', 'Mr. Deepak Jain', '99150 34567', 7)])
    su.freeze_panes = 'A4'
    # Orders: inputs A-J, calculated K-R
    title(od, 'Purchase order lines', 'One row per item of every PO. Grey columns calculate themselves.')
    head(od, 3, ['PO no.', 'PO date', 'Supplier', 'Item', 'Unit', 'Qty ordered', cur('Rate'), 'Tax %', 'Due date', 'Short-close?',
                 cur('Line value'), cur('Tax'), cur('Line total'), 'Received', 'Pending qty', cur('Pending value'), 'Status', 'Days late'])
    asof = 'Settings!$B$5'
    f = {'K': '=IF(OR($F{r}="",$G{r}=""),"",ROUND($F{r}*$G{r},2))',
         'L': '=IF($K{r}="","",ROUND($K{r}*N($H{r})/100,2))',
         'M': '=IF($K{r}="","",$K{r}+$L{r})',
         'N': '=IF($A{r}="","",SUMIFS(Receipts!$D:$D,Receipts!$B:$B,$A{r},Receipts!$C:$C,$D{r}))',
         'O': '=IF($A{r}="","",IF($J{r}="Yes",0,MAX(0,N($F{r})-$N{r})))',
         'P': '=IF(OR($O{r}="",$G{r}=""),"",ROUND($O{r}*$G{r}*(1+N($H{r})/100),2))',
         'Q': ('=IF($A{r}="","",IF($J{r}="Yes","Short-closed",IF($O{r}=0,"Closed",IF(AND($I{r}<>"",%s>$I{r}),"Overdue",'
               'IF($N{r}>0,"Part received","Open")))))' % asof),
         'R': '=IF($Q{r}="Overdue",%s-$I{r},"")' % asof}
    grid(od, 4, 503, 'ABCDEFGHIJKLMNOPQR', {c: 'calc' for c in 'KLMNOPQR'},
         {'B': D2, 'I': D2, 'F': QTY, 'G': AMT, 'K': AMT, 'L': AMT, 'M': AMT, 'N': QTY, 'O': QTY, 'P': AMT}, f)
    dv(od, 'C4:C503', '=Suppliers!$A$4:$A$203', 'Pick a supplier from the Suppliers sheet')
    dv(od, 'J4:J503', '=Lists!$A$1:$A$1', 'Yes = the balance will not come; stop showing it as pending')
    rows_in(od, 4, [('PO-101', DAY(-40), 'Bharat Steel Traders', 'MS round bar 20 mm', 'kg', 2500, 68, 18, DAY(-28)),
                    ('PO-101', DAY(-40), 'Bharat Steel Traders', 'MS plate 10 mm', 'kg', 1200, 72, 18, DAY(-28)),
                    ('PO-102', DAY(-30), 'Precision Bearings Co.', 'Ball bearing 6205-2RS', 'Nos', 400, 145, 18, DAY(-9)),
                    ('PO-102', DAY(-30), 'Precision Bearings Co.', 'Taper roller bearing 30206', 'Nos', 150, 410, 18, DAY(-9)),
                    ('PO-103', DAY(-12), 'Shree Packaging Industries', '5-ply corrugated box 600x400x400', 'Nos', 1000, 46, 5, DAY(-5)),
                    ('PO-104', DAY(-6), 'Bharat Steel Traders', 'MS angle 50x50x6', 'kg', 800, 70, 18, DAY(8)),
                    ('PO-104', DAY(-6), 'Bharat Steel Traders', 'GI sheet 1.2 mm', 'kg', 300, 88, 18, DAY(8), 'Yes')])
    od.conditional_formatting.add('A4:R503', FormulaRule(formula=['$Q4="Overdue"'], fill=PatternFill('solid', fgColor='FEE2E2')))
    od.conditional_formatting.add('A4:R503', FormulaRule(formula=['OR($Q4="Closed",$Q4="Short-closed")'], font=Font(color='94A3B8')))
    od.freeze_panes = 'E4'
    od.auto_filter.ref = 'A3:R503'
    od.page_setup.orientation = 'landscape'
    # Receipts
    title(rc, 'Receipts against POs', 'Item must be written exactly as on the PO line. "Match" must say OK.')
    head(rc, 3, ['Date', 'PO no.', 'Item', 'Qty received', 'GRN / challan no.', 'Note', 'Match'])
    grid(rc, 4, 1003, 'ABCDEFG', {'G': 'calc'}, {'A': D2, 'D': QTY},
         {'G': '=IF($B{r}="","",IF(COUNTIFS(Orders!$A:$A,$B{r},Orders!$D:$D,$C{r})>0,"OK","No such PO line"))'})
    dv(rc, 'C4:C1003', '=Orders!$D$4:$D$503')
    rows_in(rc, 4, [(DAY(-29), 'PO-101', 'MS round bar 20 mm', 2500, 'GRN-0451', 'Full quantity'),
                    (DAY(-26), 'PO-101', 'MS plate 10 mm', 700, 'GRN-0457', 'Balance 500 kg promised next week'),
                    (DAY(-11), 'PO-102', 'Ball bearing 6205-2RS', 250, 'GRN-0470', ''),
                    (DAY(-4), 'PO-103', '5-ply corrugated box 600x400x400', 1000, 'GRN-0482', ''),
                    (DAY(-3), 'PO-104', 'GI sheet 1.2 mm', 120, 'GRN-0485', 'Supplier cannot send the rest')])
    rc.conditional_formatting.add('G4:G1003', CellIsRule(operator='equal', formula=['"No such PO line"'], fill=RED, font=Font(bold=True, color='B91C1C')))
    rc.freeze_panes = 'A4'
    # Pending by supplier
    ws = pd
    title(ws, 'Pending by supplier')
    ws['A2'] = '="As on "&TEXT(%s,"dd-mmm-yyyy")&"  ·  amounts in "&Settings!$B$4' % asof; ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 4, ['Supplier', cur('Ordered value'), cur('Pending value'), 'Open lines', 'Overdue lines', 'Oldest delay (days)'])
    for i in range(200):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Suppliers!A%d="","",Suppliers!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A{r}="","",SUMIFS(Orders!$M$4:$M$503,Orders!$C$4:$C$503,$A{r}))'.format(r=r), 'calc', AMT)
        cell(ws, 'C%d' % r, '=IF($A{r}="","",SUMIFS(Orders!$P$4:$P$503,Orders!$C$4:$C$503,$A{r}))'.format(r=r), 'calc', AMT)
        cell(ws, 'D%d' % r, '=IF($A{r}="","",COUNTIFS(Orders!$C$4:$C$503,$A{r},Orders!$O$4:$O$503,">0"))'.format(r=r), 'calc', align='center')
        cell(ws, 'E%d' % r, '=IF($A{r}="","",COUNTIFS(Orders!$C$4:$C$503,$A{r},Orders!$Q$4:$Q$503,"Overdue"))'.format(r=r), 'calc', align='center')
        cell(ws, 'F%d' % r, '=IF(N($E{r})=0,"",IFERROR(_xlfn.AGGREGATE(14,6,Orders!$R$4:$R$503/(Orders!$C$4:$C$503=$A{r}),1),""))'.format(r=r), 'calc', align='center')
    label(ws, 'A3', 'Total')
    cell(ws, 'B3', '=SUM(B5:B204)', 'soft', AMT, bold=True); cell(ws, 'C3', '=SUM(C5:C204)', 'soft', AMT, bold=True)
    cell(ws, 'D3', '=SUM(D5:D204)', 'soft', bold=True, align='center'); cell(ws, 'E3', '=SUM(E5:E204)', 'soft', bold=True, align='center')
    ws.conditional_formatting.add('E5:E204', CellIsRule(operator='greaterThan', formula=['0'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'A5'
    protect(ws)
    save(wb, 'PW-Purchase-Order-Register.xlsx')

# ------------------------------------------------------------------ 2. GRN and quality check log
REASONS = ['Damaged in transit', 'Wrong item', 'Below specification', 'Expired / short shelf life', 'Wrong quantity packed', 'Poor packing', 'Other']

def grn_qc():
    wb = new_wb('pw-grn-qc-excel')
    howto(wb, first=True)
    settings(wb, [('Company name', 'Apex Engineering Works', None, None), ('Currency code', 'INR', None, None),
                  ('Target acceptance %', 0.98, '0%', None), ('Store / location', 'Main store, Plot 18, Focal Point', None, None)])
    su = sheet(wb, 'Suppliers', [32])
    lg = sheet(wb, 'GRN Log', [11, 12, 24, 10, 13, 28, 7, 11, 10, 10, 10, 22, 13, 9, 9, 8, 15, 13, 22], tab=BLUE)
    pr = sheet(wb, 'GRN Print', [5, 30, 8, 11, 11, 11, 11, 22], tab=BLUE)
    qs = sheet(wb, 'QC Summary', [30, 10, 13, 13, 13, 13, 12, 12])
    hidden_lists(wb, {1: REASONS})
    title(su, 'Suppliers', 'Used for the drop-down on the GRN Log and for the QC Summary.')
    head(su, 3, ['Supplier']); grid(su, 4, 203, 'A', {})
    rows_in(su, 4, [('Bharat Steel Traders',), ('Precision Bearings Co.',), ('Shree Packaging Industries',), ('Kumar Fasteners',)])
    # GRN Log
    title(lg, 'Goods received and quality check', 'One row per item received. Leave "Qty accepted" empty while the lot is under inspection.')
    head(lg, 3, ['GRN no.', 'Date', 'Supplier', 'PO no.', 'Challan / invoice no.', 'Item', 'Unit', cur('Rate'), 'Qty as per challan',
                 'Qty received (counted)', 'Qty accepted', 'Rejection reason', 'Inspected by',
                 'Rejected', 'Shortage', 'Excess', 'QC result', cur('Accepted value'), 'Action'])
    f = {'N': '=IF(OR($J{r}="",$K{r}=""),"",MAX(0,$J{r}-$K{r}))',
         'O': '=IF(OR($I{r}="",$J{r}=""),"",MAX(0,$I{r}-$J{r}))',
         'P': '=IF(OR($I{r}="",$J{r}=""),"",MAX(0,$J{r}-$I{r}))',
         'Q': '=IF($J{r}="","",IF($K{r}="","Pending QC",IF($K{r}>=$J{r},"Accepted",IF($K{r}=0,"Rejected","Partly accepted"))))',
         'R': '=IF(OR($K{r}="",$H{r}=""),"",ROUND($K{r}*$H{r},2))',
         'S': '=IF($J{r}="","",IF(OR(N($N{r})>0,N($O{r})>0),"Debit note / replacement",IF(N($P{r})>0,"Excess: return or accept","")))'}
    grid(lg, 4, 503, 'ABCDEFGHIJKLMNOPQRS', {c: 'calc' for c in 'NOPQRS'},
         {'B': D2, 'H': AMT, 'I': QTY, 'J': QTY, 'K': QTY, 'N': QTY, 'O': QTY, 'P': QTY, 'R': AMT}, f)
    dv(lg, 'C4:C503', '=Suppliers!$A$4:$A$203'); dv(lg, 'L4:L503', '=Lists!$A$1:$A$%d' % len(REASONS))
    rows_in(lg, 4, [('GRN-0470', DAY(-11), 'Precision Bearings Co.', 'PO-102', 'INV/2231', 'Ball bearing 6205-2RS', 'Nos', 145, 250, 250, 244, 'Below specification', 'R. Singh'),
                    ('GRN-0470', DAY(-11), 'Precision Bearings Co.', 'PO-102', 'INV/2231', 'Taper roller bearing 30206', 'Nos', 410, 100, 96, 96, None, 'R. Singh'),
                    ('GRN-0482', DAY(-4), 'Shree Packaging Industries', 'PO-103', 'DC-778', '5-ply corrugated box 600x400x400', 'Nos', 46, 1000, 1000, 1000, None, 'M. Kaur'),
                    ('GRN-0485', DAY(-3), 'Bharat Steel Traders', 'PO-104', 'BST/0912', 'GI sheet 1.2 mm', 'kg', 88, 120, 124, 124, None, 'M. Kaur'),
                    ('GRN-0486', DAY(-2), 'Kumar Fasteners', 'PO-105', 'KF-3391', 'Hex bolt M12x50 (box of 100)', 'Box', 620, 40, 40, 0, 'Wrong item', 'R. Singh'),
                    ('GRN-0487', DAY(-1), 'Bharat Steel Traders', 'PO-104', 'BST/0930', 'MS angle 50x50x6', 'kg', 70, 800, 800, None, None, None)])
    lg.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Rejected"'], fill=RED, font=Font(bold=True, color='B91C1C')))
    lg.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Partly accepted"'], fill=AMBER))
    lg.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Pending QC"'], font=Font(italic=True, color='B45309')))
    lg.conditional_formatting.add('S4:S503', FormulaRule(formula=['$S4<>""'], font=Font(bold=True, color='B91C1C')))
    lg.freeze_panes = 'B4'
    lg.auto_filter.ref = 'A3:S503'
    lg.page_setup.orientation = 'landscape'
    # GRN Print
    ws = pr
    ws.merge_cells('A1:H1'); cell(ws, 'A1', '=Settings!B3', bold=True, size=16, color=BLUE, align='center')
    ws.merge_cells('A2:H2'); cell(ws, 'A2', '=Settings!B6', size=9, align='center', color=MUTED)
    ws.merge_cells('A3:H3'); c = cell(ws, 'A3', 'GOODS RECEIVED NOTE', bold=True, size=13, align='center', color='FFFFFF'); c.fill = F_HEAD
    L = 'GRN Log'
    first = 'MATCH($C$5,\'%s\'!$A:$A,0)' % L
    hd = lambda col: '=IFERROR(INDEX(\'%s\'!$%s:$%s,%s)&"","")' % (L, col, col, first)
    ws.merge_cells('A5:B5'); label(ws, 'A5', 'GRN no. (pick)'); ws.merge_cells('C5:D5'); cell(ws, 'C5', 'GRN-0470', 'in')
    dv(ws, 'C5', '=\'%s\'!$A$4:$A$503' % L, 'Pick the GRN number to print')
    ws.merge_cells('A6:B6'); label(ws, 'A6', 'Date'); ws.merge_cells('C6:D6')
    cell(ws, 'C6', '=IFERROR(INDEX(\'%s\'!$B:$B,%s),"")' % (L, first), fmt=D2, align='left')
    ws.merge_cells('A7:B7'); label(ws, 'A7', 'Supplier'); ws.merge_cells('C7:D7'); cell(ws, 'C7', hd('C'))
    label(ws, 'E5', 'PO no.'); ws.merge_cells('F5:H5'); cell(ws, 'F5', hd('D'))
    label(ws, 'E6', 'Challan / invoice'); ws.merge_cells('F6:H6'); cell(ws, 'F6', hd('E'))
    label(ws, 'E7', 'Inspected by'); ws.merge_cells('F7:H7'); cell(ws, 'F7', hd('M'))
    head(ws, 9, ['#', 'Item', 'Unit', 'As per challan', 'Received', 'Accepted', 'Rejected', 'QC result / reason'])
    for k in range(1, 13):
        r = 9 + k
        idx = '_xlfn.AGGREGATE(15,6,ROW(\'{L}\'!$A$4:$A$503)/(\'{L}\'!$A$4:$A$503=$C$5),{k})'.format(L=L, k=k)
        cell(ws, 'A%d' % r, '=IF(ISERROR(%s),"",%d)' % (idx, k), 'box', align='center')
        pick = lambda col: '=IF($A{r}="","",INDEX(\'{L}\'!${c}:${c},{i}))'.format(r=r, L=L, c=col, i=idx)
        cell(ws, 'B%d' % r, pick('F'), 'box'); cell(ws, 'C%d' % r, pick('G'), 'box')
        for col, src in zip('DEFG', 'IJKN'): cell(ws, '%s%d' % (col, r), pick(src), 'box', QTY)
        cell(ws, 'H%d' % r, '=IF($A{r}="","",INDEX(\'{L}\'!$Q:$Q,{i})&IF(INDEX(\'{L}\'!$L:$L,{i})="",""," – "&INDEX(\'{L}\'!$L:$L,{i})))'.format(r=r, L=L, i=idx), 'box', size=9)
    cell(ws, 'B22', 'Total', 'soft', bold=True)
    for col in 'DEFG': cell(ws, '%s22' % col, '=SUM(%s10:%s21)' % (col, col), 'soft', QTY, bold=True)
    ws.merge_cells('A24:H25')
    cell(ws, 'A24', 'Received the above material subject to the quantities and quality results shown. Rejected material is held for return to the supplier. '
         'Accepted quantity only is taken into stock.', wrap=True, size=9, color=MUTED)
    for col, k in zip('BEH', ['Received by (stores)', 'Checked by (quality)', 'Approved by']):
        ws['%s29' % col] = '__________________'; cell(ws, '%s30' % col, k, size=9)
    ws.print_area = 'A1:H31'
    protect(ws)
    # QC summary
    ws = qs
    title(ws, 'Quality summary by supplier', '="Target acceptance: "&TEXT(Settings!B5,"0%")&". Below target is shown in red."')
    ws['A2'].value = '="Target acceptance: "&TEXT(Settings!B5,"0%")&". Below target is shown in red."'
    head(ws, 4, ['Supplier', 'GRN lines', 'Qty received', 'Qty accepted', 'Qty rejected', 'Acceptance %', 'Shortage qty', 'Open actions'])
    for i in range(200):
        r, s = 5 + i, 4 + i
        S = lambda col: 'SUMIFS(\'GRN Log\'!${c}$4:${c}$503,\'GRN Log\'!$C$4:$C$503,$A{r},\'GRN Log\'!$K$4:$K$503,"<>")'.format(c=col, r=r)
        cell(ws, 'A%d' % r, '=IF(Suppliers!A%d="","",Suppliers!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A{r}="","",COUNTIFS(\'GRN Log\'!$C$4:$C$503,$A{r},\'GRN Log\'!$J$4:$J$503,"<>"))'.format(r=r), 'calc', align='center')
        cell(ws, 'C%d' % r, '=IF($A%d="","",%s)' % (r, S('J')), 'calc', QTY)
        cell(ws, 'D%d' % r, '=IF($A%d="","",%s)' % (r, S('K')), 'calc', QTY)
        cell(ws, 'E%d' % r, '=IF($A%d="","",%s)' % (r, S('N')), 'calc', QTY)
        cell(ws, 'F%d' % r, '=IF(OR($A{r}="",N($C{r})=0),"",$D{r}/$C{r})'.format(r=r), 'calc', '0.0%')
        cell(ws, 'G%d' % r, '=IF($A{r}="","",SUMIFS(\'GRN Log\'!$O$4:$O$503,\'GRN Log\'!$C$4:$C$503,$A{r}))'.format(r=r), 'calc', QTY)
        cell(ws, 'H%d' % r, '=IF($A{r}="","",COUNTIFS(\'GRN Log\'!$C$4:$C$503,$A{r},\'GRN Log\'!$S$4:$S$503,"?*"))'.format(r=r), 'calc', align='center')
    ws.conditional_formatting.add('F5:F204', FormulaRule(formula=['AND(ISNUMBER($F5),$F5<Settings!$B$5)'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('F5:F204', FormulaRule(formula=['AND(ISNUMBER($F5),$F5>=Settings!$B$5)'], fill=GREEN))
    cell(ws, 'A3', 'Quantities count only lines where QC is complete (Qty accepted filled in).', size=9, color=MUTED)
    ws.freeze_panes = 'A5'
    protect(ws)
    save(wb, 'PW-GRN-and-Quality-Check-Log.xlsx')

# ------------------------------------------------------------------ 3. Supplier master and rating
CATS = ['Raw material', 'Packing material', 'Spares & consumables', 'Bought-out parts', 'Trading goods', 'Services', 'Capital equipment', 'Other']
TERMS = ['Advance', 'On delivery', '7 days', '15 days', '30 days', '45 days', '60 days', '90 days']
APPROVAL = ['Approved', 'On trial', 'On hold', 'Blocked']

def supplier_rating():
    wb = new_wb('pw-supplier-rating-excel')
    howto(wb, first=True)
    st = settings(wb, [('Company name', 'Apex Engineering Works', None, None), ('Currency code', 'INR', None, None),
                       ('Grace days before "late"', 0, '0', None), ('Weight: on time', 35, '0', None), ('Weight: quality', 35, '0', None),
                       ('Weight: in full', 15, '0', None), ('Weight: service score', 15, '0', None), ('Score for grade A', 85, '0', None), ('Score for grade B', 70, '0', None)],
                  'Weights can be any numbers; the score divides by their total.')
    cell(st, 'A13', 'Grade C = below the grade B score. Service score is your own 1 to 5 rating on the Suppliers sheet.', size=9, color=MUTED)
    sp = sheet(wb, 'Suppliers', [9, 28, 18, 18, 14, 24, 18, 18, 12, 10, 22, 11, 9, 13], tab=BLUE)
    dl = sheet(wb, 'Deliveries', [28, 11, 12, 12, 12, 10, 10, 10, 9, 8, 8, 10, 10], tab=BLUE)
    rt = sheet(wb, 'Rating', [28, 18, 11, 10, 10, 10, 9, 11, 11, 9, 8, 7], tab=BLUE)
    hidden_lists(wb, {1: CATS, 2: TERMS, 3: APPROVAL, 4: [1, 2, 3, 4, 5]})
    title(sp, 'Supplier master', 'One row per supplier. Yellow = type or pick.')
    head(sp, 3, ['Code', 'Supplier name', 'Category', 'Contact person', 'Phone', 'Email', 'City / country', 'Tax ID (GSTIN / VAT)',
                 'Payment terms', 'Standard lead time (days)', 'Minimum order / notes', 'Approval status', 'Service score (1-5)', 'Last reviewed'])
    grid(sp, 4, 203, 'ABCDEFGHIJKLMN', {}, {'N': D2})
    dv(sp, 'C4:C203', '=Lists!$A$1:$A$%d' % len(CATS)); dv(sp, 'I4:I203', '=Lists!$B$1:$B$%d' % len(TERMS))
    dv(sp, 'L4:L203', '=Lists!$C$1:$C$%d' % len(APPROVAL)); dv(sp, 'M4:M203', '=Lists!$D$1:$D$5')
    rows_in(sp, 4, [('S001', 'Bharat Steel Traders', 'Raw material', 'Anil Bansal', '98140 12345', 'sales@example.com', 'Ludhiana, India', '03AAACB1234C1Z5', '30 days', 10, 'Min. 500 kg per order', 'Approved', 4, DAY(-60)),
                    ('S002', 'Precision Bearings Co.', 'Bought-out parts', 'Ritu Kapoor', '98720 23456', 'orders@example.com', 'Delhi, India', '07AAACP2345D1Z6', '45 days', 21, '', 'Approved', 3, DAY(-90)),
                    ('S003', 'Shree Packaging Industries', 'Packing material', 'Deepak Jain', '99150 34567', 'deepak@example.com', 'Ludhiana, India', '03AAACS3456E1Z7', '15 days', 7, 'Printing extra 5 days', 'Approved', 5, DAY(-30)),
                    ('S004', 'Kumar Fasteners', 'Spares & consumables', 'Sunil Kumar', '98880 45678', 'kf@example.com', 'Jalandhar, India', '', 'On delivery', 5, '', 'On trial', 2, None)])
    sp.freeze_panes = 'C4'
    # Deliveries
    g = 'Settings!$B$5'
    title(dl, 'Deliveries', 'One row per delivery (or per PO line). Grey columns calculate themselves.')
    head(dl, 3, ['Supplier', 'PO no.', 'PO date', 'Promised date', 'Received date', 'Qty ordered', 'Qty received', 'Qty accepted',
                 'Days late', 'On time', 'In full', 'Actual lead time', 'Quality %'])
    f = {'I': '=IF(OR($D{r}="",$E{r}=""),"",MAX(0,$E{r}-$D{r}-%s))' % g,
         'J': '=IF(OR($D{r}="",$E{r}=""),"",IF($E{r}<=$D{r}+%s,1,0))' % g,
         'K': '=IF(OR($E{r}="",$F{r}=""),"",IF(N($G{r})>=$F{r},1,0))',
         'L': '=IF(OR($C{r}="",$E{r}=""),"",$E{r}-$C{r})',
         'M': '=IF(N($G{r})=0,"",N($H{r})/$G{r})'}
    grid(dl, 4, 1003, 'ABCDEFGHIJKLM', {c: 'calc' for c in 'IJKLM'}, {'C': D2, 'D': D2, 'E': D2, 'F': QTY, 'G': QTY, 'H': QTY, 'M': '0.0%'}, f)
    for c in 'JK': dl.conditional_formatting.add('%s4:%s1003' % (c, c), CellIsRule(operator='equal', formula=['0'], fill=RED))
    dv(dl, 'A4:A1003', '=Suppliers!$B$4:$B$203')
    data = [('Bharat Steel Traders', 'PO-088', -120, 10, 9, 2000, 2000, 2000), ('Bharat Steel Traders', 'PO-095', -80, 10, 12, 1500, 1500, 1480),
            ('Bharat Steel Traders', 'PO-101', -40, 12, 11, 3700, 3200, 3200), ('Bharat Steel Traders', 'PO-104', -6, 14, 3, 1100, 924, 924),
            ('Precision Bearings Co.', 'PO-090', -110, 21, 27, 300, 300, 296), ('Precision Bearings Co.', 'PO-102', -30, 21, 19, 550, 350, 340),
            ('Shree Packaging Industries', 'PO-093', -95, 7, 7, 2000, 2000, 2000), ('Shree Packaging Industries', 'PO-103', -12, 7, 8, 1000, 1000, 1000),
            ('Kumar Fasteners', 'PO-105', -9, 5, 7, 40, 40, 0)]
    rows_in(dl, 4, [(s, po, DAY(d), DAY(d + pr), DAY(d + ac), q, rq, aq) for s, po, d, pr, ac, q, rq, aq in data])
    dl.freeze_panes = 'B4'
    # Rating
    ws = rt
    title(ws, 'Supplier rating', 'Score out of 100 = weighted on time, quality, in full and service. Only suppliers with deliveries are scored.')
    head(ws, 3, ['Supplier', 'Category', 'Deliveries', 'On time %', 'In full %', 'Quality %', 'Service (1-5)', 'Actual lead time (avg days)',
                 'Standard lead time', 'Score', 'Grade', 'Rank'])
    W = '(Settings!$B$6+Settings!$B$7+Settings!$B$8+Settings!$B$9)'
    for i in range(200):
        r, s = 4 + i, 4 + i
        D = lambda col: "Deliveries!$%s$4:$%s$1003" % (col, col)
        cell(ws, 'A%d' % r, '=IF(Suppliers!B%d="","",Suppliers!B%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Suppliers!C%d&"")' % (r, s), 'box')
        cell(ws, 'C%d' % r, '=IF($A{r}="","",COUNTIFS({a},$A{r},{e},"<>"))'.format(r=r, a=D('A'), e=D('E')), 'calc', align='center')
        cell(ws, 'D%d' % r, '=IF(N($C{r})=0,"",SUMIFS({j},{a},$A{r})/$C{r})'.format(r=r, j=D('J'), a=D('A')), 'calc', '0%')
        cell(ws, 'E%d' % r, '=IF(N($C{r})=0,"",SUMIFS({k},{a},$A{r})/$C{r})'.format(r=r, k=D('K'), a=D('A')), 'calc', '0%')
        cell(ws, 'F%d' % r, '=IF(N($C{r})=0,"",IFERROR(SUMIFS({h},{a},$A{r})/SUMIFS({g},{a},$A{r}),""))'.format(r=r, h=D('H'), g=D('G'), a=D('A')), 'calc', '0.0%')
        cell(ws, 'G%d' % r, '=IF($A%d="","",IF(Suppliers!M%d="","",Suppliers!M%d))' % (r, s, s), 'calc', align='center')
        cell(ws, 'H%d' % r, '=IF(N($C{r})=0,"",IFERROR(AVERAGEIFS({l},{a},$A{r}),""))'.format(r=r, l=D('L'), a=D('A')), 'calc', '0.0')
        cell(ws, 'I%d' % r, '=IF($A%d="","",IF(Suppliers!J%d="","",Suppliers!J%d))' % (r, s, s), 'calc', align='center')
        cell(ws, 'J%d' % r, ('=IF(N($C{r})=0,"",ROUND((N($D{r})*Settings!$B$6+IF($F{r}="",1,$F{r})*Settings!$B$7+N($E{r})*Settings!$B$8'
                             '+IF($G{r}="",0.6,$G{r}/5)*Settings!$B$9)/%s*100,1))' % W).format(r=r), 'soft', '0.0', bold=True)
        cell(ws, 'K%d' % r, '=IF($J{r}="","",IF($J{r}>=Settings!$B$10,"A",IF($J{r}>=Settings!$B$11,"B","C")))'.format(r=r), 'calc', bold=True, align='center')
        cell(ws, 'L%d' % r, '=IF($J{r}="","",COUNTIF($J$4:$J$203,">"&$J{r})+1)'.format(r=r), 'calc', align='center')
    ws.conditional_formatting.add('K4:K203', CellIsRule(operator='equal', formula=['"A"'], fill=GREEN))
    ws.conditional_formatting.add('K4:K203', CellIsRule(operator='equal', formula=['"B"'], fill=AMBER))
    ws.conditional_formatting.add('K4:K203', CellIsRule(operator='equal', formula=['"C"'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('H4:H203', FormulaRule(formula=['AND(ISNUMBER($H4),ISNUMBER($I4),$H4>$I4)'], font=Font(bold=True, color='B91C1C')))
    cell(ws, 'A205', 'No service score yet counts as 3 of 5. A supplier with no accepted-quantity data counts as 100% quality. Red lead time = slower than promised.', size=9, color=MUTED)
    ws.freeze_panes = 'B4'
    protect(ws)
    save(wb, 'PW-Supplier-Master-and-Rating.xlsx')

# ------------------------------------------------------------------ 4. Multi-location inventory
NLOC = 8
LOCS = ['Main warehouse', 'Factory store', 'Delhi branch', 'Site: Phase 2']

def multi_location():
    wb = new_wb('pw-multi-location-excel')
    howto(wb, first=True)
    lc = sheet(wb, 'Locations', [30, 26])
    title(lc, 'Locations and settings', 'Up to 8 locations. Names become the column headings on the Opening and Stock sheets.')
    label(lc, 'A3', 'Company name'); cell(lc, 'B3', 'Apex Engineering Works', 'in')
    label(lc, 'A4', 'Currency code'); cell(lc, 'B4', 'INR', 'in')
    head(lc, 6, ['Location', 'Address / in-charge'])
    for i in range(NLOC):
        cell(lc, 'A%d' % (7 + i), LOCS[i] if i < len(LOCS) else None, 'in'); cell(lc, 'B%d' % (7 + i), None, 'in')
    lc['B7'], lc['B8'], lc['B9'], lc['B10'] = 'Plot 18, Focal Point · Mr. Gill', 'Plant 1 · Ms. Kaur', 'Naraina Industrial Area · Mr. Arora', 'Project site, Phase 2'
    it = sheet(wb, 'Items', [12, 32, 8, 12, 11, 11, 10])
    op = sheet(wb, 'Opening', [12, 32] + [13] * NLOC)
    mv = sheet(wb, 'Movements', [12, 12, 30, 10, 18, 18, 11, 16, 26], tab=BLUE)
    sk = sheet(wb, 'Stock', [12, 30, 7] + [11] * NLOC + [10, 10, 10, 11, 14], tab=BLUE)
    # Items
    title(it, 'Items', 'Short unique codes. Reorder and maximum levels are for all locations together.')
    head(it, 3, ['Item code', 'Item name', 'Unit', '="Unit cost ("&Locations!$B$4&")"', 'Reorder level', 'Maximum level', 'Lead time (days)'])
    grid(it, 4, 503, 'ABCDEFG', {}, {'D': AMT})
    items = [('RM-001', 'MS round bar 20 mm', 'kg', 68, 1500, 5000, 10), ('RM-002', 'MS plate 10 mm', 'kg', 72, 800, 3000, 10),
             ('BO-101', 'Ball bearing 6205-2RS', 'Nos', 145, 200, 800, 21), ('PK-201', '5-ply corrugated box 600x400x400', 'Nos', 46, 300, 1500, 7),
             ('SP-301', 'Hex bolt M12x50 (box of 100)', 'Box', 620, 10, 40, 5), ('CN-401', 'Cutting oil 20 L', 'Can', 2450, 4, 12, 7)]
    rows_in(it, 4, items)
    it.freeze_panes = 'A4'
    # Opening
    title(op, 'Opening stock by location', 'Stock held in each location on the day you start.')
    head(op, 3, ['Item code', 'Item name'] + ['=IF(Locations!$A$%d="","(location %d)",Locations!$A$%d)' % (7 + i, i + 1, 7 + i) for i in range(NLOC)])
    for i in range(500):
        r = 4 + i
        cell(op, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d)' % (r, r), 'box'); cell(op, 'B%d' % r, '=IF(Items!B%d="","",Items!B%d)' % (r, r), 'box')
        for j in range(NLOC): cell(op, '%s%d' % (L(3 + j), r), None, 'in', QTY)
    rows_in(op, 4, [(None, None, 2200, 400, 0, 0), (None, None, 1100, 150), (None, None, 260, 40, 60), (None, None, 900, 0, 120),
                    (None, None, 12, 6, 2, 3), (None, None, 3, 2)])
    op.freeze_panes = 'C4'
    # Movements
    title(mv, 'Stock movements', 'Receipt: fill To only. Issue: fill From only. Transfer: fill both.')
    head(mv, 3, ['Date', 'Item code', 'Item name', 'Quantity', 'From location', 'To location', 'Type', 'Reference', 'Note'])
    locs = '=Locations!$A$7:$A$%d' % (6 + NLOC)
    grid(mv, 4, 2003, 'ABCDEFGHI', {'C': 'calc', 'G': 'calc'}, {'A': D2, 'D': QTY},
         {'C': '=IF($B{r}="","",IFERROR(INDEX(Items!$B:$B,MATCH($B{r},Items!$A:$A,0)),"Unknown code"))',
          'G': '=IF($D{r}="","",IF(AND($E{r}="",$F{r}<>""),"Receipt",IF(AND($E{r}<>"",$F{r}=""),"Issue",IF(AND($E{r}<>"",$F{r}<>""),IF($E{r}=$F{r},"Same location","Transfer"),"Location missing"))))'})
    dv(mv, 'B4:B2003', '=Items!$A$4:$A$503'); dv(mv, 'E4:F2003', locs)
    rows_in(mv, 4, [(DAY(-9), 'RM-001', None, 2500, None, 'Main warehouse', None, 'GRN-0451', 'PO-101'),
                    (DAY(-8), 'RM-001', None, 900, 'Main warehouse', 'Factory store', None, 'TR-031', 'For job 2215'),
                    (DAY(-7), 'RM-001', None, 640, 'Factory store', None, None, 'MRS-1182', 'Job 2215'),
                    (DAY(-6), 'BO-101', None, 120, 'Main warehouse', 'Delhi branch', None, 'TR-032', ''),
                    (DAY(-5), 'BO-101', None, 150, 'Delhi branch', None, None, 'INV-5520', 'Sold to Gupta Motors'),
                    (DAY(-4), 'PK-201', None, 820, 'Main warehouse', None, None, 'MRS-1190', 'Dispatch packing'),
                    (DAY(-3), 'SP-301', None, 9, 'Main warehouse', 'Site: Phase 2', None, 'TR-033', ''),
                    (DAY(-2), 'CN-401', None, 2, 'Factory store', None, None, 'MRS-1194', 'Machine shop'),
                    (DAY(-1), 'RM-002', None, 40, 'Main warehouse', None, None, 'ADJ-07', 'Physical count shortage')])
    mv.conditional_formatting.add('G4:G2003', FormulaRule(formula=['OR($G4="Location missing",$G4="Same location")'], fill=RED))
    mv.conditional_formatting.add('C4:C2003', CellIsRule(operator='equal', formula=['"Unknown code"'], fill=RED))
    mv.freeze_panes = 'C4'
    # Stock
    ws = sk
    title(ws, 'Stock by location')
    ws['A2'] = '="Quantities by location · values in "&Locations!$B$4'; ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    lc0 = 4                                   # first location column (D)
    tcol, rcol, scol, ocol, vcol = [L(lc0 + NLOC + k) for k in range(5)]
    head(ws, 3, ['Item code', 'Item name', 'Unit'] + ['=Opening!%s3' % L(3 + j) for j in range(NLOC)] +
         ['Total', 'Reorder level', 'Status', 'Suggested order', '="Value ("&Locations!$B$4&")"'])
    MV = lambda c: 'Movements!$%s$4:$%s$2003' % (c, c)
    for i in range(500):
        r = 4 + i
        cell(ws, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d)' % (r, r), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Items!B%d)' % (r, r), 'box')
        cell(ws, 'C%d' % r, '=IF($A%d="","",Items!C%d&"")' % (r, r), 'box')
        for j in range(NLOC):
            col, oc = L(lc0 + j), L(3 + j)
            cell(ws, '%s%d' % (col, r), ('=IF(OR($A{r}="",Locations!$A${l}=""),"",N(Opening!{oc}{r})+SUMIFS({d},{b},$A{r},{f},Locations!$A${l})'
                                         '-SUMIFS({d},{b},$A{r},{e},Locations!$A${l}))').format(r=r, l=7 + j, oc=oc, d=MV('D'), b=MV('B'), e=MV('E'), f=MV('F')), 'calc', QTY)
        cell(ws, '%s%d' % (tcol, r), '=IF($A{r}="","",SUM({a}{r}:{b}{r}))'.format(r=r, a=L(lc0), b=L(lc0 + NLOC - 1)), 'soft', QTY, bold=True)
        cell(ws, '%s%d' % (rcol, r), '=IF($A%d="","",N(Items!E%d))' % (r, r), 'calc', QTY)
        cell(ws, '%s%d' % (scol, r), '=IF($A{r}="","",IF(COUNTIF({a}{r}:{b}{r},"<0")>0,"Negative",IF({t}{r}<={rl}{r},"Reorder","OK")))'.format(
            r=r, a=L(lc0), b=L(lc0 + NLOC - 1), t=tcol, rl=rcol), 'calc', align='center')
        cell(ws, '%s%d' % (ocol, r), '=IF($A{r}="","",IF({t}{r}<={rl}{r},MAX(0,N(Items!F{r})-{t}{r}),0))'.format(r=r, t=tcol, rl=rcol), 'calc', QTY)
        cell(ws, '%s%d' % (vcol, r), '=IF($A{r}="","",ROUND({t}{r}*N(Items!D{r}),2))'.format(r=r, t=tcol), 'calc', AMT)
    last = 503
    ws.conditional_formatting.add('%s4:%s%d' % (scol, scol, last), CellIsRule(operator='equal', formula=['"Reorder"'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('%s4:%s%d' % (scol, scol, last), CellIsRule(operator='equal', formula=['"Negative"'], fill=RED, font=Font(bold=True, color='7F1D1D')))
    ws.conditional_formatting.add('D4:%s%d' % (tcol, last), CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    # value by location
    fr = last + 2
    cell(ws, 'B%d' % fr, '="Stock value by location ("&Locations!$B$4&")"', 'soft', bold=True)
    for j in range(NLOC):
        col = L(lc0 + j)
        cell(ws, '%s%d' % (col, fr), '=IF(Locations!$A${l}="","",SUMPRODUCT({c}4:{c}{last},Items!$D$4:$D${last}))'.format(l=7 + j, c=col, last=last), 'soft', AMT0)
    cell(ws, '%s%d' % (tcol, fr), '=SUM(%s4:%s%d)' % (vcol, vcol, last), 'soft', AMT0, bold=True)
    label(ws, 'L2', 'Items to reorder'); cell(ws, 'N2', '=COUNTIF(%s4:%s%d,"Reorder")' % (scol, scol, last), 'soft', bold=True, align='center')
    ws.freeze_panes = 'C4'
    ws.page_setup.orientation = 'landscape'
    protect(ws)
    save(wb, 'PW-Multi-Location-Inventory.xlsx')

# ------------------------------------------------------------------ 5. ABC analysis and stock valuation
def abc_valuation():
    wb = new_wb('pw-abc-valuation-excel')
    howto(wb, first=True)
    settings(wb, [('Company name', 'Apex Engineering Works', None, None), ('Currency code', 'INR', None, None),
                  ('Period from', datetime.date(TODAY.year, 4, 1) if TODAY.month >= 4 else datetime.date(TODAY.year - 1, 4, 1), D2, None),
                  ('Period to', '=EDATE(B5,12)-1', D2, None), ('Class A up to (cumulative share)', 0.8, '0%', None), ('Class B up to', 0.95, '0%', None)],
             'Change the cut-offs to suit you. Common choices: 70/90, 80/95.')
    it = sheet(wb, 'Items', [12, 34, 8])
    rc = sheet(wb, 'Receipts', [12, 12, 28, 10, 12, 14, 18, 12, 14], tab=BLUE)
    is_ = sheet(wb, 'Issues', [12, 12, 28, 10, 26], tab=BLUE)
    va = sheet(wb, 'Valuation', [12, 30, 11, 14, 11, 11, 12, 15, 15, 13, 15], tab=BLUE)
    ab = sheet(wb, 'ABC', [12, 30, 16, 9, 12, 12, 7, 7, 3, 12, 9, 12, 12], tab=BLUE)
    title(it, 'Items', 'One row per item, with a unique code.'); head(it, 3, ['Item code', 'Item name', 'Unit']); grid(it, 4, 503, 'ABC', {})
    items = [('RM-001', 'MS round bar 20 mm', 'kg'), ('RM-002', 'MS plate 10 mm', 'kg'), ('BO-101', 'Ball bearing 6205-2RS', 'Nos'),
             ('BO-102', 'Taper roller bearing 30206', 'Nos'), ('PK-201', '5-ply corrugated box 600x400x400', 'Nos'), ('SP-301', 'Hex bolt M12x50 (box of 100)', 'Box'),
             ('CN-401', 'Cutting oil 20 L', 'Can'), ('CN-402', 'Welding electrode 3.15 mm (packet)', 'Pkt'), ('SP-302', 'Grinding wheel 180 mm', 'Nos'),
             ('PK-202', 'Stretch film roll', 'Roll')]
    rows_in(it, 4, items)
    it.freeze_panes = 'A4'
    lookup = '=IF($B{r}="","",IFERROR(INDEX(Items!$B:$B,MATCH($B{r},Items!$A:$A,0)),"Unknown code"))'
    # Receipts with FIFO layers
    title(rc, 'Receipts (oldest first)', 'Opening stock first, then purchases in date order. Unit cost = landed cost per unit.')
    head(rc, 3, ['Date', 'Item code', 'Item name', 'Quantity', cur('Unit cost'), cur('Value'), 'Reference', 'FIFO qty left', cur('FIFO value left')])
    R = 'Receipts'
    f = {'C': lookup, 'F': '=IF(OR($D{r}="",$E{r}=""),"",ROUND($D{r}*$E{r},2))',
         'H': ('=IF($B{r}="","",MAX(0,MIN(N($D{r}),SUMIF($B$4:$B$1003,$B{r},$D$4:$D$1003)-SUMIF(Issues!$B$4:$B$2003,$B{r},Issues!$D$4:$D$2003)'
               '-SUMIF($B{n}:$B$1004,$B{r},$D{n}:$D$1004))))'),
         'I': '=IF(OR($H{r}="",$E{r}=""),"",ROUND($H{r}*$E{r},2))'}
    for r in range(4, 1004):
        for col in 'ABCDEFGHI':
            tpl = f.get(col)
            cell(rc, '%s%d' % (col, r), tpl.format(r=r, n=r + 1) if tpl else None, 'calc' if col in 'CFHI' else 'in',
                 {'A': D2, 'D': QTY, 'E': AMT, 'F': AMT, 'H': QTY, 'I': AMT}.get(col))
    d0 = datetime.date(TODAY.year, 4, 1) if TODAY.month >= 4 else datetime.date(TODAY.year - 1, 4, 1)
    rd = lambda n: d0 + datetime.timedelta(days=n)
    rows_in(rc, 4, [(rd(0), 'RM-001', None, 1200, 64, None, 'Opening'), (rd(0), 'BO-101', None, 300, 138, None, 'Opening'),
                    (rd(0), 'PK-201', None, 500, 44, None, 'Opening'), (rd(0), 'CN-401', None, 6, 2300, None, 'Opening'),
                    (rd(12), 'RM-001', None, 3000, 66, None, 'GRN-0402'), (rd(20), 'RM-002', None, 2000, 70, None, 'GRN-0410'),
                    (rd(31), 'BO-102', None, 200, 395, None, 'GRN-0418'), (rd(45), 'SP-301', None, 30, 600, None, 'GRN-0425'),
                    (rd(60), 'RM-001', None, 2500, 68, None, 'GRN-0451'), (rd(64), 'CN-402', None, 120, 240, None, 'GRN-0455'),
                    (rd(70), 'BO-101', None, 400, 145, None, 'GRN-0470'), (rd(74), 'PK-201', None, 1000, 46, None, 'GRN-0482'),
                    (rd(80), 'SP-302', None, 50, 95, None, 'GRN-0489'), (rd(85), 'PK-202', None, 40, 310, None, 'GRN-0491'),
                    (rd(90), 'RM-002', None, 1200, 72, None, 'GRN-0497'), (rd(96), 'CN-401', None, 8, 2450, None, 'GRN-0503')])
    dv(rc, 'B4:B1003', '=Items!$A$4:$A$503')
    rc.freeze_panes = 'A4'
    # Issues
    title(is_, 'Issues', 'Every issue to production, a job, a department or a customer, in quantity.')
    head(is_, 3, ['Date', 'Item code', 'Item name', 'Quantity', 'Issued to / reference'])
    grid(is_, 4, 2003, 'ABCDE', {'C': 'calc'}, {'A': D2, 'D': QTY}, {'C': lookup})
    dv(is_, 'B4:B2003', '=Items!$A$4:$A$503')
    rows_in(is_, 4, [(rd(15), 'RM-001', None, 2600, 'Machine shop'), (rd(40), 'RM-002', None, 1300, 'Fabrication'), (rd(50), 'BO-101', None, 280, 'Assembly'),
                     (rd(55), 'BO-102', None, 120, 'Assembly'), (rd(62), 'PK-201', None, 900, 'Dispatch'), (rd(66), 'RM-001', None, 2100, 'Machine shop'),
                     (rd(72), 'CN-401', None, 9, 'Machine shop'), (rd(75), 'SP-301', None, 18, 'Maintenance'), (rd(82), 'CN-402', None, 70, 'Fabrication'),
                     (rd(88), 'BO-101', None, 300, 'Assembly'), (rd(92), 'SP-302', None, 22, 'Maintenance'), (rd(98), 'RM-002', None, 700, 'Fabrication')])
    is_.freeze_panes = 'A4'
    # Valuation
    ws = va
    title(ws, 'Stock valuation: FIFO and weighted average')
    ws['A2'] = '=Settings!B3&"  ·  "&TEXT(Settings!B5,"dd-mmm-yyyy")&" to "&TEXT(Settings!B6,"dd-mmm-yyyy")&"  ·  amounts in "&Settings!B4'
    ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 4, ['Item code', 'Item name', 'Received qty', cur('Received value'), 'Issued qty', 'Closing qty', cur('Average cost'),
                 cur('Closing value (average)'), cur('Closing value (FIFO)'), cur('FIFO − average'), cur('Consumption value')])
    for i in range(500):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Items!B%d)' % (r, s), 'box')
        cell(ws, 'C%d' % r, '=IF($A{r}="","",SUMIF(Receipts!$B$4:$B$1003,$A{r},Receipts!$D$4:$D$1003))'.format(r=r), 'calc', QTY)
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIF(Receipts!$B$4:$B$1003,$A{r},Receipts!$F$4:$F$1003))'.format(r=r), 'calc', AMT)
        cell(ws, 'E%d' % r, '=IF($A{r}="","",SUMIF(Issues!$B$4:$B$2003,$A{r},Issues!$D$4:$D$2003))'.format(r=r), 'calc', QTY)
        cell(ws, 'F%d' % r, '=IF($A{r}="","",$C{r}-$E{r})'.format(r=r), 'calc', QTY)
        cell(ws, 'G%d' % r, '=IF(OR($A{r}="",N($C{r})=0),"",ROUND($D{r}/$C{r},2))'.format(r=r), 'calc', AMT)
        cell(ws, 'H%d' % r, '=IF($G{r}="","",ROUND(MAX(0,$F{r})*$G{r},2))'.format(r=r), 'calc', AMT)
        cell(ws, 'I%d' % r, '=IF($A{r}="","",SUMIF(Receipts!$B$4:$B$1003,$A{r},Receipts!$I$4:$I$1003))'.format(r=r), 'calc', AMT)
        cell(ws, 'J%d' % r, '=IF(OR($H{r}="",$I{r}=""),"",$I{r}-$H{r})'.format(r=r), 'calc', AMT)
        cell(ws, 'K%d' % r, '=IF($G{r}="","",ROUND($E{r}*$G{r},2))'.format(r=r), 'calc', AMT)
    label(ws, 'A3', 'Total')
    for col in 'DHIJK': cell(ws, '%s3' % col, '=SUM(%s5:%s504)' % (col, col), 'soft', AMT, bold=True)
    ws.conditional_formatting.add('F5:F504', CellIsRule(operator='lessThan', formula=['0'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'C5'
    protect(ws)
    # ABC
    ws = ab
    title(ws, 'ABC analysis by consumption value', '="Class A up to "&TEXT(Settings!B7,"0%")&" of value, B up to "&TEXT(Settings!B8,"0%")&", C the rest."')
    ws['A2'].value = '="Class A up to "&TEXT(Settings!B7,"0%")&" of value, B up to "&TEXT(Settings!B8,"0%")&", C the rest."'
    head(ws, 3, ['Item code', 'Item name', cur('Consumption value'), 'Share', 'Cumulative before', 'Cumulative after', 'Rank', 'Class'])
    TOT = 'SUM(Valuation!$K$5:$K$504)'
    for i in range(500):
        r, s = 4 + i, 5 + i
        cell(ws, 'A%d' % r, '=IF(Valuation!A%d="","",Valuation!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Valuation!B%d)' % (r, s), 'box')
        cell(ws, 'C%d' % r, '=IF($A%d="","",N(Valuation!K%d))' % (r, s), 'calc', AMT)
        cell(ws, 'D%d' % r, '=IF(OR($A{r}="",{t}=0),"",$C{r}/{t})'.format(r=r, t=TOT), 'calc', '0.0%')
        cell(ws, 'E%d' % r, '=IF($D{r}="","",SUMIF($C$4:$C$503,">"&$C{r})/{t})'.format(r=r, t=TOT), 'calc', '0.0%')
        cell(ws, 'F%d' % r, '=IF($D{r}="","",SUMIF($C$4:$C$503,">="&$C{r})/{t})'.format(r=r, t=TOT), 'calc', '0.0%')
        cell(ws, 'G%d' % r, '=IF($D{r}="","",RANK($C{r},$C$4:$C$503))'.format(r=r), 'calc', align='center')
        cell(ws, 'H%d' % r, '=IF($D{r}="","",IF($C{r}=0,"C",IF($E{r}<Settings!$B$7,"A",IF($E{r}<Settings!$B$8,"B","C"))))'.format(r=r), 'calc', bold=True, align='center')
    for k, fl in (('A', GREEN), ('B', AMBER), ('C', PatternFill('solid', fgColor='E2E8F0'))):
        ws.conditional_formatting.add('H4:H503', CellIsRule(operator='equal', formula=['"%s"' % k], fill=fl))
    head(ws, 3, ['Class', 'Items', '% of items', '% of value'], col=10)
    for i, k in enumerate('ABC', 4):
        cell(ws, 'J%d' % i, k, 'box', bold=True, align='center')
        cell(ws, 'K%d' % i, '=COUNTIF($H$4:$H$503,J%d)' % i, 'calc', align='center')
        cell(ws, 'L%d' % i, '=IFERROR(K%d/COUNTIF($H$4:$H$503,"?"),0)' % i, 'calc', '0%')
        cell(ws, 'M%d' % i, '=IFERROR(SUMIF($H$4:$H$503,J%d,$C$4:$C$503)/SUM($C$4:$C$503),0)' % i, 'calc', '0%')
    ch = BarChart(); ch.type = 'col'; ch.style = 10; ch.title = 'Share of items vs share of value'
    ch.add_data(Reference(ws, min_col=12, max_col=13, min_row=3, max_row=6), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=10, min_row=4, max_row=6))
    ch.series[0].graphicalProperties.solidFill = 'A5B4FC'; ch.series[1].graphicalProperties.solidFill = '4F46E5'
    ch.y_axis.numFmt = '0%'; ch.y_axis.majorGridlines = None
    ch.height, ch.width = 7.5, 13
    ws.add_chart(ch, 'J9')
    ws.freeze_panes = 'C4'
    protect(ws)
    save(wb, 'PW-ABC-Analysis-and-Stock-Valuation.xlsx')

# ------------------------------------------------------------------ 6. Material requisition slip (Word)
def requisition():
    d = doc_new('pw-material-requisition-word')
    letterhead(d, '[Your Company Name]', '[Plant / store address] · Phone [ ] · [email]', gst=False)
    para(d, 'MATERIAL REQUISITION SLIP', bold=True, size=15, align='c', space=6)
    kv(d, [('MRS No.', 'MRS/[2026]/[001]'), ('Date', '[DD-MM-YYYY]'), ('Department', '[ ]'), ('Cost centre / job no.', '[ ]'),
           ('Required by (date)', '[DD-MM-YYYY]'), ('Deliver to', '[Section / machine / site]'), ('Purpose', '[Production / maintenance / project / office use]'), ('Priority', '[Normal / Urgent]')])
    table(d, ['#', 'Item code', 'Description', 'Unit', 'Qty requested', 'Qty issued', 'Balance to purchase', 'Bin / location', 'Remarks'],
          [['1', '[ ]', '[Item description, size, grade]', '[ ]', '[ ]', '', '', '', '']], [0.8, 1.8, 4.6, 1.1, 1.6, 1.5, 1.7, 1.6, 1.8], fill_rows=9)
    para(d, 'Store check (tick one):  ☐ Issued in full    ☐ Partly issued, balance to purchase    ☐ Not in stock, purchase requisition raised: PR No. [ ]', size=10, space=10)
    t = d.add_table(rows=2, cols=4); t.style = 'Table Grid'
    for i, k in enumerate(['Requested by', 'Approved by (head of department)', 'Issued by (store)', 'Received by']):
        t.rows[0].cells[i].text = k; shade(t.rows[0].cells[i], 'E8F0FE')
        t.rows[1].cells[i].text = '\n\nName:\nDate:'
        for c in (t.rows[0].cells[i], t.rows[1].cells[i]):
            for r in c.paragraphs[0].runs: r.font.size = Pt(9)
    para(d, space=6)
    para(d, 'Notes', bold=True, color=BLUE, space=2)
    for s in ['Material is issued only against an approved slip. One slip per department and cost centre.',
              'Unused material must be returned to the store with a material return note (overleaf).',
              'The store posts the issued quantity in the stock register on the same day.']:
        d.add_paragraph(s, style='List Bullet')
    d.add_page_break()
    letterhead(d, '[Your Company Name]', '[Plant / store address] · Phone [ ] · [email]', gst=False)
    para(d, 'MATERIAL RETURN NOTE', bold=True, size=15, align='c', space=6)
    kv(d, [('Return No.', 'MRN/[2026]/[001]'), ('Date', '[DD-MM-YYYY]'), ('Returned by (department)', '[ ]'), ('Against MRS No.', '[ ]'),
           ('Cost centre / job no.', '[ ]'), ('Reason', '[Excess / job cancelled / wrong item / defective]')])
    table(d, ['#', 'Item code', 'Description', 'Unit', 'Qty returned', 'Condition', 'Taken to bin / location'],
          [['1', '[ ]', '[ ]', '[ ]', '[ ]', '[Good / damaged / scrap]', '[ ]']], [0.8, 2, 5.5, 1.2, 2, 2.6, 2.5], fill_rows=6)
    para(d, 'Condition decides where the material goes: good material back to stock, damaged material to quarantine for inspection, scrap to the scrap yard.', size=9, italic=True, space=18)
    t = d.add_table(rows=2, cols=3); t.style = 'Table Grid'
    for i, k in enumerate(['Returned by', 'Received by (store)', 'Checked by (quality, if damaged)']):
        t.rows[0].cells[i].text = k; shade(t.rows[0].cells[i], 'E8F0FE'); t.rows[1].cells[i].text = '\n\nName:\nDate:'
    dsave(d, 'PW-Material-Requisition-Slip.docx')

MAKERS = [po_register, grn_qc, supplier_rating, multi_location, abc_valuation, requisition]
