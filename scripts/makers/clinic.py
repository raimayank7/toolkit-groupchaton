"""Clinic & Pharmacy Admin toolkit: OPD register, patient billing, pharmacy stock and expiry, lab samples, duty roster, equipment log.
   Admin records only (no diagnosis, dosage or treatment content). Every sheet with patient data carries a privacy line."""
from make_files import *
from openpyxl.worksheet.formula import ArrayFormula

RED = PatternFill('solid', fgColor='FECACA')
AMBER = PatternFill('solid', fgColor='FDE68A')
GREEN = PatternFill('solid', fgColor='DCFCE7')
RED_FONT = Font(bold=True, color='B91C1C')
PRIVATE = 'Private patient data: keep this file on your own computer and password-protect it (File › Info › Protect Workbook).'
DT = 'dd-mm-yyyy hh:mm'
D = datetime.timedelta

def rows_in(ws, r0, rows, col=1):
    for i, row in enumerate(rows, r0):
        for j, v in enumerate(row):
            if v is not None: ws.cell(i, col + j, v)

def arrayize(ws, col, r1, r2):
    """MIN(IF(...)) / MAX(IF(...)) over a range must be an array formula (Ctrl+Shift+Enter) to work in every Excel version"""
    for r in range(r1, r2 + 1):
        c = ws['%s%d' % (col, r)]
        c.value = ArrayFormula(c.coordinate, c.value)

def setting(ws, r, k, v, fmt=None):
    label(ws, 'A%d' % r, k); cell(ws, 'B%d' % r, v, 'in', fmt, align='left')

def list_block(ws, col, r1, heading, values, n, fmt=None, extra=None):
    """an editable yellow list on a Settings sheet: heading in row r1-1, n rows"""
    label(ws, '%s%d' % (col, r1 - 1), heading)
    for i in range(n):
        cell(ws, '%s%d' % (col, r1 + i), values[i] if i < len(values) else None, 'in', fmt)

# ------------------------------------------------------------------ OPD register
def opd_register():
    wb = new_wb('clinic-opd-register-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [24, 30, 3, 26, 3, 24, 12, 3, 30])
    vs = sheet(wb, 'Visits', [11, 7, 7, 10, 24, 6, 6, 13, 11, 18, 18, 10, 10, 10, 10, 10, 22], tab=BLUE)
    dw = sheet(wb, 'Day-wise', [13, 10, 9, 10, 14, 13, 13, 13, 13, 14, 13], tab=BLUE)
    dr = sheet(wb, 'Doctor-wise', [26, 10, 10, 15, 15, 15])
    # Settings
    title(st, 'Settings', 'Fill once a month. The yellow lists feed the drop-downs on the Visits sheet.')
    setting(st, 3, 'Clinic name', 'Sehat Family Clinic')
    setting(st, 4, 'Month (first day)', MONTH1, 'mmmm yyyy')
    setting(st, 5, 'Payment modes', 'Cash, UPI, Card, Cheque, Insurance')
    st['B5'].fill, st['B5'].protection = F_CALC, Protection(locked=True)
    list_block(st, 'D', 3, 'Doctors', ['Dr. A. Sharma', 'Dr. R. Kaur', 'Dr. S. Verma (visiting)'], 10)
    list_block(st, 'F', 3, 'Visit type', ['New consultation', 'Follow-up', 'Review (free)', 'Emergency', 'Procedure'], 10)
    list_block(st, 'G', 3, 'Usual fee', [500, 300, 0, 800, 1000], 10, INR0)
    st.merge_cells('I3:I8')
    cell(st, 'I3', 'Change the doctors, visit types and fees to match your clinic. A fee of 0 is fine for free reviews. '
         'The Fee column on the Visits sheet starts with the usual fee; type over it when needed.', wrap=True, size=9, color=MUTED)
    cell(st, 'A15', PRIVATE, size=9, color=MUTED)
    # Visits
    title(vs, 'OPD visits', PRIVATE)
    head(vs, 3, ['Date', 'Token', 'Time', 'Patient ID', 'Patient name', 'Age', 'Sex', 'Phone', 'New / Follow-up', 'Doctor', 'Visit type',
                 'Fee', 'Discount', 'Paid', 'Mode', 'Balance', 'Remarks'])
    R2 = 1503
    grid(vs, 4, R2, 'ABCDEFGHIJKLMNOPQ', {'P': 'calc'}, {'A': DATE, 'C': 'hh:mm', 'L': INR0, 'M': INR0, 'N': INR0, 'P': INR0},
         {'L': '=IF($K{r}="","",IFERROR(VLOOKUP($K{r},Settings!$F$3:$G$12,2,FALSE),""))',
          'P': '=IF(AND($L{r}="",$N{r}=""),"",N($L{r})-N($M{r})-N($N{r}))'})
    d = MONTH1
    t = lambda h, m: datetime.time(h, m)
    rows_in(vs, 4, [
        (d, 1, t(9, 40), 'P-0001', 'Ramesh Kumar', 54, 'M', '98140 11111', 'New', 'Dr. A. Sharma', 'New consultation', None, None, 500, 'Cash'),
        (d, 2, t(9, 55), 'P-0002', 'Sunita Devi', 47, 'F', '98140 22222', 'Follow-up', 'Dr. A. Sharma', 'Follow-up', None, None, 300, 'UPI'),
        (d, 3, t(10, 20), 'P-0003', 'Harpreet Singh', 31, 'M', '98140 33333', 'New', 'Dr. R. Kaur', 'New consultation', None, 100, 400, 'Cash'),
        (d, 4, t(11, 5), 'P-0001', 'Ramesh Kumar', 54, 'M', '98140 11111', 'Follow-up', 'Dr. A. Sharma', 'Review (free)', None, None, 0, 'Cash'),
        (d, 5, t(12, 30), 'P-0004', 'Asha Rani', 68, 'F', '98140 44444', 'New', 'Dr. S. Verma (visiting)', 'Procedure', None, None, 600, 'Card'),
        (d + D(days=1), 1, t(9, 35), 'P-0005', 'Mohit Arora', 26, 'M', '98140 55555', 'New', 'Dr. R. Kaur', 'New consultation', None, None, 500, 'UPI'),
        (d + D(days=1), 2, t(10, 10), 'P-0002', 'Sunita Devi', 47, 'F', '98140 22222', 'Follow-up', 'Dr. A. Sharma', 'Follow-up', None, None, 300, 'Cash'),
        (d + D(days=1), 3, t(19, 45), 'P-0006', 'Kiran Bala', 39, 'F', '98140 66666', 'New', 'Dr. A. Sharma', 'Emergency', None, None, 800, 'UPI')])
    vs['Q8'] = 'Balance to pay on next visit'
    dv(vs, 'G4:G%d' % R2, '"M,F,Other"')
    dv(vs, 'I4:I%d' % R2, '"New,Follow-up"')
    dv(vs, 'J4:J%d' % R2, '=Settings!$D$3:$D$12', 'Pick the doctor (list on the Settings sheet)')
    dv(vs, 'K4:K%d' % R2, '=Settings!$F$3:$F$12', 'Pick the visit type; the usual fee fills in')
    dv(vs, 'O4:O%d' % R2, '"Cash,UPI,Card,Cheque,Insurance"')
    vs.conditional_formatting.add('P4:P%d' % R2, CellIsRule(operator='greaterThan', formula=['0'], fill=RED, font=RED_FONT))
    vs.freeze_panes = 'F4'
    vs.auto_filter.ref = 'A3:Q%d' % R2
    vs.page_setup.orientation = 'landscape'
    # Day-wise
    title(dw, 'Day-wise patients and collection', '="Month: "&TEXT(Settings!B4,"mmmm yyyy")&"   ·   "&Settings!B3')
    dw['A2'].value = '="Month: "&TEXT(Settings!B4,"mmmm yyyy")&"   ·   "&Settings!B3'
    head(dw, 3, ['Date', 'Patients', 'New', 'Follow-up', 'Fees charged', 'Cash', 'UPI', 'Card', 'Other modes', 'Total collected', 'Balance due'])
    V = lambda col, extra='': 'SUMIFS(Visits!$%s:$%s,Visits!$A:$A,$A{r}%s)' % (col, col, extra)
    for n in range(31):
        r = 4 + n
        cell(dw, 'A%d' % r, '=IF(MONTH(Settings!$B$4+%d)=MONTH(Settings!$B$4),Settings!$B$4+%d,"")' % (n, n), 'box', DATE)
        f = {'B': '=IF($A{r}="","",COUNTIFS(Visits!$A:$A,$A{r},Visits!$E:$E,"<>"))',
             'C': '=IF($A{r}="","",COUNTIFS(Visits!$A:$A,$A{r},Visits!$I:$I,"New"))',
             'D': '=IF($A{r}="","",COUNTIFS(Visits!$A:$A,$A{r},Visits!$I:$I,"Follow-up"))',
             'E': '=IF($A{r}="","",%s-%s)' % (V('L'), V('M')),
             'F': '=IF($A{r}="","",%s)' % V('N', ',Visits!$O:$O,"Cash"'),
             'G': '=IF($A{r}="","",%s)' % V('N', ',Visits!$O:$O,"UPI"'),
             'H': '=IF($A{r}="","",%s)' % V('N', ',Visits!$O:$O,"Card"'),
             'I': '=IF($A{r}="","",J{r}-F{r}-G{r}-H{r})',
             'J': '=IF($A{r}="","",%s)' % V('N'),
             'K': '=IF($A{r}="","",SUMIFS(Visits!$P:$P,Visits!$A:$A,$A{r},Visits!$P:$P,">0"))'}
        for col, fx in f.items():
            cell(dw, '%s%d' % (col, r), fx.format(r=r), 'calc', None if col in 'BCD' else INR0, align='center' if col in 'BCD' else None)
    cell(dw, 'A35', 'Month total', 'soft', bold=True)
    for col in 'BCDEFGHIJK':
        cell(dw, '%s35' % col, '=SUM(%s4:%s34)' % (col, col), 'soft', None if col in 'BCD' else INR0, bold=True, align='center' if col in 'BCD' else None)
    dw.merge_cells('A37:K38')
    cell(dw, 'A37', 'Cash = what should be in the cash drawer from OPD fees that day. Other modes = cheque and insurance. '
         'Balance due = fees not yet paid by patients who visited that day.', wrap=True, size=9, color=MUTED)
    dw.freeze_panes = 'B4'
    protect(dw)
    # Doctor-wise
    title(dr, 'Doctor-wise visits and fees', 'For the visits on the Visits sheet (one file per month).')
    head(dr, 3, ['Doctor', 'Visits', 'New patients', 'Fees charged', 'Collected', 'Balance due'])
    for i in range(10):
        r, s = 4 + i, 3 + i
        cell(dr, 'A%d' % r, '=IF(Settings!D%d="","",Settings!D%d)' % (s, s), 'box')
        f = {'B': '=IF($A{r}="","",COUNTIF(Visits!$J:$J,$A{r}))', 'C': '=IF($A{r}="","",COUNTIFS(Visits!$J:$J,$A{r},Visits!$I:$I,"New"))',
             'D': '=IF($A{r}="","",SUMIFS(Visits!$L:$L,Visits!$J:$J,$A{r})-SUMIFS(Visits!$M:$M,Visits!$J:$J,$A{r}))',
             'E': '=IF($A{r}="","",SUMIFS(Visits!$N:$N,Visits!$J:$J,$A{r}))',
             'F': '=IF($A{r}="","",SUMIFS(Visits!$P:$P,Visits!$J:$J,$A{r},Visits!$P:$P,">0"))'}
        for col, fx in f.items(): cell(dr, '%s%d' % (col, r), fx.format(r=r), 'calc', None if col in 'BC' else INR0, align='center' if col in 'BC' else None)
    cell(dr, 'A14', 'Total', 'soft', bold=True)
    for col in 'BCDEF': cell(dr, '%s14' % col, '=SUM(%s4:%s13)' % (col, col), 'soft', None if col in 'BC' else INR0, bold=True, align='center' if col in 'BC' else None)
    protect(dr)
    save(wb, 'Clinic-OPD-Register.xlsx')

# ------------------------------------------------------------------ Patient billing register
SERVICES = [('Consultation (OPD)', 'OPD', 500), ('Registration / file charges', 'OPD', 100), ('Medical certificate', 'OPD', 200),
            ('Dressing (small)', 'Procedure', 150), ('Dressing (large)', 'Procedure', 350), ('Injection charges', 'Procedure', 50),
            ('Nebulisation', 'Procedure', 150), ('Suturing (minor)', 'Procedure', 800), ('Plaster application', 'Procedure', 1200),
            ('ECG', 'Diagnostics', 300), ('Blood sugar (glucometer)', 'Lab', 60), ('CBC', 'Lab', 350), ('Lipid profile', 'Lab', 600),
            ('X-ray chest PA view', 'Radiology', 450), ('Ultrasound abdomen', 'Radiology', 1200),
            ('Room charges: general ward (per day)', 'Ward', 1500), ('Room charges: private room (per day)', 'Ward', 3500),
            ('Nursing charges (per day)', 'Ward', 500), ('Doctor visit (inpatient)', 'Ward', 600), ('Oxygen (per hour)', 'Ward', 150),
            ('Day-care bed', 'Ward', 1000), ('Physiotherapy session', 'Physiotherapy', 400), ('Dental scaling', 'Dental', 1000),
            ('Ambulance (local)', 'Other', 800), ('Medicines and consumables (as per pharmacy bill)', 'Pharmacy', 0)]
DEPTS = ['OPD', 'Procedure', 'Diagnostics', 'Lab', 'Radiology', 'Ward', 'Physiotherapy', 'Dental', 'Pharmacy', 'Other']
MODES = ['Cash', 'UPI', 'Card', 'Cheque', 'Bank transfer', 'Insurance / TPA']

def billing_register():
    wb = new_wb('clinic-billing-register-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [24, 46])
    sv = sheet(wb, 'Services', [46, 16, 13])
    bl = sheet(wb, 'Bills', [12, 11, 24, 11, 13, 22, 12, 12, 12, 12, 12, 11], tab=BLUE)
    ln = sheet(wb, 'Bill lines', [12, 11, 44, 14, 7, 11, 12, 11, 12, 13], tab=BLUE)
    py = sheet(wb, 'Payments', [11, 12, 12, 16, 30, 13], tab=BLUE)
    rc = sheet(wb, 'Receipt', [5, 38, 7, 12, 12, 14, 3, 8], tab=BLUE)
    sm = sheet(wb, 'Summary', [34, 16, 16])
    lists_sheet(wb, extra={3: DEPTS, 5: MODES})
    # Settings
    title(st, 'Settings', 'These details print on every receipt.')
    for r, (k, v, fmt) in enumerate([('Clinic / hospital name', 'Sehat Nursing Home', None), ('Address', '45, Civil Lines, Ludhiana, Punjab 141001', None),
                                     ('Phone', '0161 400 0000', None), ('Registration no.', 'Clinical Establishment Reg. No. [ ]', None),
                                     ('Month for the summary', MONTH1, 'mmmm yyyy')], 3):
        setting(st, r, k, v, fmt)
    cell(st, 'A10', PRIVATE, size=9, color=MUTED)
    # Services
    title(sv, 'Services and rates', 'Your price list. The Bill lines sheet picks the rate from here.')
    head(sv, 3, ['Service', 'Department', 'Rate'])
    grid(sv, 4, 203, 'ABC', {}, {'C': INR0})
    rows_in(sv, 4, SERVICES)
    dv(sv, 'B4:B203', '=Lists!$C$1:$C$10')
    sv.freeze_panes = 'A4'
    # Bills
    title(bl, 'Bills', PRIVATE)
    head(bl, 3, ['Bill no.', 'Date', 'Patient name', 'Patient ID', 'Phone', 'Doctor / department', 'Gross', 'Discount', 'Net bill', 'Received', 'Balance', 'Status'])
    LN, PY = "'Bill lines'", 'Payments'
    grid(bl, 4, 503, 'ABCDEFGHIJKL', {c: 'calc' for c in 'GHIJKL'}, {'B': DATE, **{c: INR0 for c in 'GHIJK'}},
         {'G': '=IF($A{r}="","",SUMIFS(%s!$G:$G,%s!$A:$A,$A{r}))' % (LN, LN),
          'H': '=IF($A{r}="","",SUMIFS(%s!$H:$H,%s!$A:$A,$A{r}))' % (LN, LN),
          'I': '=IF($A{r}="","",G{r}-H{r})',
          'J': '=IF($A{r}="","",SUMIFS(Payments!$C:$C,Payments!$B:$B,$A{r}))',
          'K': '=IF($A{r}="","",I{r}-J{r})',
          'L': '=IF($A{r}="","",IF(I{r}=0,"No services",IF(K{r}<0,"Refund due",IF(K{r}=0,"Paid",IF(J{r}>0,"Part paid","Due")))))'})
    d = MONTH1
    rows_in(bl, 4, [('B-001', d, 'Ramesh Kumar', 'P-0001', '98140 11111', 'Dr. A. Sharma'),
                    ('B-002', d, 'Asha Rani', 'P-0004', '98140 44444', 'Day-care'),
                    ('B-003', d + D(days=1), 'Gurmeet Kaur', 'P-0007', '98140 77777', 'Ward (Dr. R. Kaur)')])
    bl.conditional_formatting.add('L4:L503', CellIsRule(operator='equal', formula=['"Due"'], fill=RED, font=RED_FONT))
    bl.conditional_formatting.add('L4:L503', CellIsRule(operator='equal', formula=['"Part paid"'], fill=AMBER))
    bl.conditional_formatting.add('L4:L503', CellIsRule(operator='equal', formula=['"Paid"'], fill=GREEN))
    bl.freeze_panes = 'D4'
    # Bill lines
    title(ln, 'Bill lines', 'One row per service. Same bill number for every line of a bill. Empty quantity counts as 1. Date and rate fill in; type over them if needed.')
    head(ln, 3, ['Bill no.', 'Date', 'Service', 'Department', 'Qty', 'Rate', 'Amount', 'Discount', 'Net', 'Key (for receipt)'])
    grid(ln, 4, 2003, 'ABCDEFGHIJ', {c: 'calc' for c in 'DGIJ'}, {'B': DATE, 'F': INR0, 'G': INR0, 'H': INR0, 'I': INR0},
         {'B': '=IF($A{r}="","",IFERROR(INDEX(Bills!$B:$B,MATCH($A{r},Bills!$A:$A,0)),""))',
          'D': '=IF($C{r}="","",IFERROR(VLOOKUP($C{r},Services!$A$4:$C$203,2,FALSE),"Other"))',
          'F': '=IF($C{r}="","",IFERROR(VLOOKUP($C{r},Services!$A$4:$C$203,3,FALSE),0))',
          'G': '=IF($C{r}="","",IF($E{r}="",1,N($E{r}))*N($F{r}))',
          'I': '=IF($C{r}="","",G{r}-N(H{r}))',
          'J': '=IF(OR($A{r}="",$C{r}=""),"",$A{r}&"#"&COUNTIFS($A$4:$A{r},$A{r},$C$4:$C{r},"<>"))'})
    rows_in(ln, 4, [('B-001', None, 'Consultation (OPD)'), ('B-001', None, 'ECG'), ('B-001', None, 'Blood sugar (glucometer)'),
                    ('B-002', None, 'Day-care bed'), ('B-002', None, 'Dressing (large)', None, 2), ('B-002', None, 'Injection charges', None, 3),
                    ('B-003', None, 'Room charges: general ward (per day)', None, 3), ('B-003', None, 'Nursing charges (per day)', None, 3),
                    ('B-003', None, 'Doctor visit (inpatient)', None, 4), ('B-003', None, 'Medicines and consumables (as per pharmacy bill)', None, 1, 2860),
                    ('B-003', None, 'X-ray chest PA view')])
    ln['H6'] = 0; ln['H14'] = 300
    dv(ln, 'A4:A2003', '=Bills!$A$4:$A$503', 'Pick the bill number (add the bill on the Bills sheet first)')
    dv(ln, 'C4:C2003', '=Services!$A$4:$A$203', 'Pick a service from the Services sheet')
    ln.freeze_panes = 'A4'
    # Payments
    title(py, 'Payments received', 'Every advance, part payment and final payment, against its bill number.')
    head(py, 3, ['Date', 'Bill no.', 'Amount', 'Mode', 'Reference / note', 'Key (for receipt)'])
    grid(py, 4, 2003, 'ABCDEF', {'F': 'calc'}, {'A': DATE, 'C': INR0},
         {'F': '=IF(OR($B{r}="",$C{r}=""),"",$B{r}&"#"&COUNTIFS($B$4:$B{r},$B{r},$C$4:$C{r},"<>"))'})
    rows_in(py, 4, [(d, 'B-001', 860, 'Cash', ''), (d, 'B-002', 1000, 'UPI', 'Advance, UPI ref 4512'),
                    (d + D(days=1), 'B-003', 5000, 'Card', 'Advance at admission'), (d + D(days=4), 'B-003', 4000, 'Cash', 'Part payment')])
    dv(py, 'B4:B2003', '=Bills!$A$4:$A$503')
    dv(py, 'D4:D2003', '=Lists!$E$1:$E$6')
    py.freeze_panes = 'A4'
    # Receipt
    ws = rc
    for r, f, sz in ((1, '=Settings!B3', 17), (2, '=Settings!B4', 10), (3, '="Phone: "&Settings!B5&"   ·   "&Settings!B6', 9)):
        ws.merge_cells('A%d:F%d' % (r, r)); cell(ws, 'A%d' % r, f, bold=(r == 1), size=sz, align='center', color=BLUE if r == 1 else INK)
    ws.merge_cells('A5:F5'); c = cell(ws, 'A5', 'PAYMENT RECEIPT', bold=True, size=13, align='center', color='FFFFFF'); c.fill = F_HEAD
    ws['H6'] = 'Row'; ws['H6'].font = Font(size=8, color=MUTED)
    cell(ws, 'H7', '=IFERROR(MATCH($B$7,Bills!$A:$A,0),"")', size=8, color=MUTED)
    B = lambda col: '=IF($H$7="","",INDEX(Bills!$%s:$%s,$H$7))' % (col, col)
    label(ws, 'A7', 'Bill no.'); cell(ws, 'B7', 'B-003', 'in', bold=True)
    label(ws, 'D7', 'Bill date'); ws.merge_cells('E7:F7'); cell(ws, 'E7', B('B'), fmt=DATE, align='left')
    label(ws, 'A8', 'Patient'); cell(ws, 'B8', B('C'), bold=True)
    label(ws, 'D8', 'Patient ID'); ws.merge_cells('E8:F8'); cell(ws, 'E8', B('D'), align='left')
    label(ws, 'A9', 'Doctor'); cell(ws, 'B9', B('F'))
    label(ws, 'D9', 'Phone'); ws.merge_cells('E9:F9'); cell(ws, 'E9', B('E'), align='left')
    ws.column_dimensions['A'].width = 11
    head(ws, 11, ['#', 'Service', 'Qty', 'Rate', 'Discount', 'Amount'])
    LNI = lambda col, r: '=IF($H%d="","",INDEX(%s!$%s:$%s,$H%d))' % (r, LN, col, col, r)
    for k in range(1, 16):
        r = 11 + k
        cell(ws, 'H%d' % r, '=IF($B$7="","",IFERROR(MATCH($B$7&"#%d",%s!$J:$J,0),""))' % (k, LN), size=8, color='FFFFFF')
        cell(ws, 'A%d' % r, '=IF($H%d="","",%d)' % (r, k), 'box', align='center')
        cell(ws, 'B%d' % r, LNI('C', r), 'box')
        cell(ws, 'C%d' % r, '=IF($H{r}="","",IF(INDEX({ln}!$E:$E,$H{r})="",1,INDEX({ln}!$E:$E,$H{r})))'.format(r=r, ln=LN), 'box', align='center')
        cell(ws, 'D%d' % r, LNI('F', r), 'box', INR0)
        cell(ws, 'E%d' % r, '=IF($H{r}="","",IF(N(INDEX({ln}!$H:$H,$H{r}))=0,"",INDEX({ln}!$H:$H,$H{r})))'.format(r=r, ln=LN), 'box', INR0)
        cell(ws, 'F%d' % r, LNI('I', r), 'box', INR0)
    tot = [('Gross', 'G'), ('Discount', 'H'), ('Net bill', 'I'), ('Received', 'J'), ('Balance', 'K')]
    for i, (k, col) in enumerate(tot, 28):
        ws.merge_cells('D%d:E%d' % (i, i)); label(ws, 'D%d' % i, k, bold=k in ('Net bill', 'Balance'))
        cell(ws, 'F%d' % i, B(col), 'calc', INR0, bold=k in ('Net bill', 'Balance'))
    ws['F30'].fill = F_SOFT
    label(ws, 'A28', 'Received in words')
    ws.merge_cells('A29:C31'); cell(ws, 'A29', words_formula('F31'), wrap=True, bold=True)
    head(ws, 34, ['#', 'Payment date', 'Mode', 'Reference / note', '', 'Amount'])
    ws.merge_cells('D34:E34')
    PYI = lambda col, r: '=IF($H%d="","",INDEX(Payments!$%s:$%s,$H%d))' % (r, col, col, r)
    for k in range(1, 7):
        r = 34 + k
        cell(ws, 'H%d' % r, '=IF($B$7="","",IFERROR(MATCH($B$7&"#%d",Payments!$F:$F,0),""))' % k, size=8, color='FFFFFF')
        cell(ws, 'A%d' % r, '=IF($H%d="","",%d)' % (r, k), 'box', align='center')
        cell(ws, 'B%d' % r, PYI('A', r), 'box', DATE, align='left')
        cell(ws, 'C%d' % r, PYI('D', r), 'box')
        ws.merge_cells('D%d:E%d' % (r, r)); cell(ws, 'D%d' % r, PYI('E', r), 'box', size=9)
        cell(ws, 'F%d' % r, PYI('C', r), 'box', INR0)
    ws.merge_cells('A42:F43')
    cell(ws, 'A42', 'Received with thanks. This receipt is proof of payment only; it is not a prescription or a medical record. Please keep it for your records.',
         wrap=True, size=9, color=MUTED)
    ws.merge_cells('D46:F46'); cell(ws, 'D46', '="For "&Settings!B3', bold=True, align='right')
    ws.merge_cells('D48:F48'); cell(ws, 'D48', 'Authorised signatory', align='right', size=9)
    cell(ws, 'A48', 'Patient / attendant', size=9)
    ws.print_area = 'A1:F49'
    protect(ws)
    # Summary
    title(sm, 'Billing summary', '="Month: "&TEXT(Settings!B7,"mmmm yyyy")')
    sm['A2'].value = '="Month: "&TEXT(Settings!B7,"mmmm yyyy")'
    inm = lambda rng: '%s,">="&Settings!$B$7,%s,"<="&EOMONTH(Settings!$B$7,0)' % (rng, rng)
    head(sm, 4, ['This month', 'Amount', 'Count'])
    items = [('Net billed (bills dated this month)', '=SUMIFS(Bills!$I:$I,%s)' % inm('Bills!$B:$B'), '=COUNTIFS(%s)' % inm('Bills!$B:$B')),
             ('Received (payments dated this month)', '=SUMIFS(Payments!$C:$C,%s)' % inm('Payments!$A:$A'), '=COUNTIFS(%s)' % inm('Payments!$A:$A'))]
    for i, (k, f, n) in enumerate(items, 5):
        cell(sm, 'A%d' % i, k, 'box'); cell(sm, 'B%d' % i, f, 'calc', INR0); cell(sm, 'C%d' % i, n, 'calc', align='center')
    head(sm, 8, ['All bills in this file', 'Amount', 'Bills'])
    for i, (k, st_) in enumerate([('Due (nothing received)', 'Due'), ('Part paid', 'Part paid'), ('Refund due', 'Refund due')], 9):
        cell(sm, 'A%d' % i, k, 'box')
        cell(sm, 'B%d' % i, '=SUMIFS(Bills!$K:$K,Bills!$L:$L,"%s")' % st_, 'calc', INR0)
        cell(sm, 'C%d' % i, '=COUNTIF(Bills!$L:$L,"%s")' % st_, 'calc', align='center')
    cell(sm, 'A12', 'Total still to collect', 'soft', bold=True); cell(sm, 'B12', '=SUMIF(Bills!$K:$K,">0")', 'soft', INR0, bold=True)
    head(sm, 14, ['Received this month by mode', 'Amount', 'Payments'])
    for i, m in enumerate(MODES, 15):
        cell(sm, 'A%d' % i, m, 'box')
        cell(sm, 'B%d' % i, '=SUMIFS(Payments!$C:$C,Payments!$D:$D,$A%d,%s)' % (i, inm('Payments!$A:$A')), 'calc', INR0)
        cell(sm, 'C%d' % i, '=COUNTIFS(Payments!$D:$D,$A%d,%s)' % (i, inm('Payments!$A:$A')), 'calc', align='center')
    head(sm, 22, ['Billed this month by department', 'Net amount', 'Lines'])
    for i, dpt in enumerate(DEPTS, 23):
        cell(sm, 'A%d' % i, dpt, 'box')
        cell(sm, 'B%d' % i, "=SUMIFS('Bill lines'!$I:$I,'Bill lines'!$D:$D,$A%d,%s)" % (i, inm("'Bill lines'!$B:$B")), 'calc', INR0)
        cell(sm, 'C%d' % i, "=COUNTIFS('Bill lines'!$D:$D,$A%d,%s)" % (i, inm("'Bill lines'!$B:$B")), 'calc', align='center')
    protect(sm)
    save(wb, 'Clinic-Patient-Billing-Register.xlsx')

# ------------------------------------------------------------------ Pharmacy stock & expiry
def pharmacy_stock():
    wb = new_wb('pharmacy-stock-expiry-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 22, 3, 60])
    it = sheet(wb, 'Items', [34, 12, 16, 10, 9, 10, 11, 13, 10])
    bt = sheet(wb, 'Batches', [34, 12, 12, 22, 12, 12, 10, 11, 10, 10, 10, 10, 12, 12, 22], tab=BLUE)
    iss = sheet(wb, 'Issues', [12, 34, 12, 9, 26, 16], tab=BLUE)
    ex = sheet(wb, 'Expiry summary', [34, 12, 16], tab=BLUE)
    title(st, 'Settings', 'Fill once. Other sheets read these values.')
    setting(st, 3, 'Pharmacy / clinic name', 'Sehat Clinic Dispensary')
    setting(st, 4, 'Stock and expiry as on', '=TODAY()', DATE)
    setting(st, 5, 'Near-expiry warning (days)', 90)
    st.merge_cells('D3:D7')
    cell(st, 'D3', 'Leave "as on" as =TODAY(), or type a date to see the position on that day. A batch turns amber when its expiry date '
         'is within the warning days, and red once it has expired.', wrap=True, size=9, color=MUTED)
    cell(st, 'A9', 'Patient names in the Issues sheet are private: keep this file on your own computer.', size=9, color=MUTED)
    # Items
    title(it, 'Items', 'One row per medicine or product. Balance adds up all its batches.')
    head(it, 3, ['Item', 'Form', 'Pack', 'Count in', 'Rack', 'Reorder level', 'Balance', 'Nearest expiry (in stock)', 'Status'])
    grid(it, 4, 503, 'ABCDEFGHI', {c: 'calc' for c in 'GHI'}, {'H': DATE},
         {'G': '=IF($A{r}="","",SUMIFS(Batches!$K:$K,Batches!$A:$A,$A{r}))',
          'H': '=IF($A{r}="","",IFERROR(1/(1/MIN(IF((Batches!$A$4:$A$1503=$A{r})*(Batches!$C$4:$C$1503<>"")*(Batches!$K$4:$K$1503>0),Batches!$C$4:$C$1503))),""))',
          'I': '=IF($A{r}="","",IF(G{r}<=N(F{r}),"Reorder","OK"))'})
    arrayize(it, 'H', 4, 503)
    rows_in(it, 4, [('Paracetamol 500 mg tablets', 'Tablet', 'Strip of 10', 'Strip', 'A1', 30),
                    ('Cetirizine 10 mg tablets', 'Tablet', 'Strip of 10', 'Strip', 'A1', 20),
                    ('Amoxicillin 500 mg capsules', 'Capsule', 'Strip of 10', 'Strip', 'A2', 20),
                    ('ORS sachets 21 g', 'Powder', 'Sachet', 'Sachet', 'B1', 50),
                    ('Cough syrup 100 ml', 'Syrup', 'Bottle', 'Bottle', 'B2', 10),
                    ('Disposable syringe 5 ml', 'Surgical', 'Box of 100', 'Piece', 'C1', 200),
                    ('Absorbent cotton roll 500 g', 'Consumable', 'Roll', 'Roll', 'C2', 5),
                    ('Normal saline 500 ml', 'Injection', 'Bottle', 'Bottle', 'C3', 24)])
    dv(it, 'B4:B503', '"Tablet,Capsule,Syrup,Injection,Ointment,Drops,Powder,Inhaler,Surgical,Consumable,Other"')
    it.conditional_formatting.add('I4:I503', CellIsRule(operator='equal', formula=['"Reorder"'], fill=RED, font=RED_FONT))
    it.freeze_panes = 'B4'
    # Batches
    title(bt, 'Batches', 'One row per batch received. Expiry "03/2027" = type 31-03-2027. Grey columns calculate themselves.')
    head(bt, 3, ['Item', 'Batch no.', 'Expiry date', 'Supplier', 'Invoice no.', 'Received on', 'Qty received', 'Purchase rate', 'MRP',
                 'Issued', 'Balance', 'Days to expiry', 'Status', 'Stock value', 'Remarks'])
    grid(bt, 4, 1503, 'ABCDEFGHIJKLMNO', {c: 'calc' for c in 'JKLMN'}, {'C': DATE, 'F': DATE, 'H': INR, 'I': INR, 'N': INR},
         {'J': '=IF($A{r}="","",SUMIFS(Issues!$D:$D,Issues!$B:$B,$A{r},Issues!$C:$C,$B{r}))',
          'K': '=IF($A{r}="","",N($G{r})-J{r})',
          'L': '=IF(OR($A{r}="",$C{r}=""),"",$C{r}-Settings!$B$4)',
          'M': '=IF($A{r}="","",IF(K{r}<=0,"Nil stock",IF($C{r}="","OK",IF($C{r}<Settings!$B$4,"Expired",IF(L{r}<=Settings!$B$5,"Near expiry","OK")))))',
          'N': '=IF($A{r}="","",MAX(0,K{r})*N($H{r}))'})
    t = TODAY
    eom = lambda dd: datetime.date(dd.year + (dd.month // 12), dd.month % 12 + 1, 1) - D(days=1)
    rows_in(bt, 4, [('Paracetamol 500 mg tablets', 'PCM2391', eom(t - D(days=20)), 'Punjab Medical Agencies', 'PMA/1182', t - D(days=400), 40, 9.5, 15),
                    ('Paracetamol 500 mg tablets', 'PCM2544', eom(t + D(days=430)), 'Punjab Medical Agencies', 'PMA/1460', t - D(days=30), 100, 9.8, 15.5),
                    ('Cetirizine 10 mg tablets', 'CTZ0871', eom(t + D(days=45)), 'City Pharma Distributors', 'CPD/552', t - D(days=300), 50, 14, 22),
                    ('Amoxicillin 500 mg capsules', 'AMX1129', eom(t + D(days=200)), 'City Pharma Distributors', 'CPD/601', t - D(days=90), 60, 62, 98),
                    ('ORS sachets 21 g', 'ORS5510', eom(t + D(days=75)), 'Punjab Medical Agencies', 'PMA/1390', t - D(days=150), 120, 16, 21),
                    ('Cough syrup 100 ml', 'CSY0042', eom(t + D(days=25)), 'Wellness Surgicals', 'WS/77', t - D(days=320), 12, 68, 105),
                    ('Disposable syringe 5 ml', 'SYR8812', eom(t + D(days=900)), 'Wellness Surgicals', 'WS/91', t - D(days=60), 500, 4.2, 7),
                    ('Absorbent cotton roll 500 g', 'COT0315', None, 'Wellness Surgicals', 'WS/91', t - D(days=60), 10, 145, 210),
                    ('Normal saline 500 ml', 'NS77801', eom(t + D(days=500)), 'City Pharma Distributors', 'CPD/640', t - D(days=10), 48, 22, 35)])
    bt['O4'] = 'Keep apart; return to supplier'
    dv(bt, 'A4:A1503', '=Items!$A$4:$A$503', 'Pick the item from the Items sheet')
    bt.conditional_formatting.add('A4:O1503', FormulaRule(formula=['$M4="Expired"'], fill=RED, font=RED_FONT))
    bt.conditional_formatting.add('A4:O1503', FormulaRule(formula=['$M4="Near expiry"'], fill=AMBER))
    bt.freeze_panes = 'C4'
    bt.auto_filter.ref = 'A3:O1503'
    # Issues
    title(iss, 'Sales and issues', 'One row per item sold or issued. Pick the batch it came from: the oldest expiry first.')
    head(iss, 3, ['Date', 'Item', 'Batch no.', 'Qty', 'Bill no. / patient ref', 'Check'])
    grid(iss, 4, 2003, 'ABCDEF', {'F': 'calc'}, {'A': DATE},
         {'F': '=IF($B{r}="","",IF(COUNTIFS(Batches!$A:$A,$B{r},Batches!$B:$B,$C{r})=0,"Batch not found","OK"))'})
    rows_in(iss, 4, [(t - D(days=6), 'Paracetamol 500 mg tablets', 'PCM2544', 6, 'B-1021'), (t - D(days=6), 'ORS sachets 21 g', 'ORS5510', 10, 'B-1021'),
                     (t - D(days=5), 'Cetirizine 10 mg tablets', 'CTZ0871', 4, 'B-1025'), (t - D(days=4), 'Disposable syringe 5 ml', 'SYR8812', 40, 'OPD use'),
                     (t - D(days=3), 'Amoxicillin 500 mg capsules', 'AMX1129', 5, 'B-1033'), (t - D(days=2), 'Normal saline 500 ml', 'NS77801', 6, 'Day-care ward'),
                     (t - D(days=1), 'Cough syrup 100 ml', 'CSY0042', 3, 'B-1040'), (t - D(days=1), 'Paracetamol 500 mg tablets', 'PCM2544', 12, 'B-1041')])
    dv(iss, 'B4:B2003', '=Items!$A$4:$A$503')
    dv(iss, 'C4:C2003', '=Batches!$B$4:$B$1503', 'Pick the batch number (see the Batches sheet)')
    iss.conditional_formatting.add('F4:F2003', CellIsRule(operator='equal', formula=['"Batch not found"'], fill=RED, font=RED_FONT))
    iss.freeze_panes = 'A4'
    # Expiry summary
    title(ex, 'Expiry summary', '="As on "&TEXT(Settings!B4,"dd-mm-yyyy")&"   ·   batches with stock left"')
    ex['A2'].value = '="As on "&TEXT(Settings!B4,"dd-mm-yyyy")&"   ·   batches with stock left"'
    head(ex, 4, ['Expiry', 'Batches', 'Stock value'])
    bands = [('Already expired', '"<0"', None), ('Expiring in 0 to 30 days', '">=0"', '"<=30"'), ('Expiring in 31 to 60 days', '">=31"', '"<=60"'),
             ('Expiring in 61 to 90 days', '">=61"', '"<=90"'), ('More than 90 days away', '">90"', None)]
    for i, (k, a, b) in enumerate(bands, 5):
        crit = 'Batches!$L:$L,%s%s,Batches!$K:$K,">0"' % (a, (',Batches!$L:$L,%s' % b) if b else '')
        cell(ex, 'A%d' % i, k, 'box')
        cell(ex, 'B%d' % i, '=COUNTIFS(%s)' % crit, 'calc', align='center')
        cell(ex, 'C%d' % i, '=SUMIFS(Batches!$N:$N,%s)' % crit, 'calc', INR)
    cell(ex, 'A10', 'No expiry date entered', 'box')
    cell(ex, 'B10', '=COUNTIFS(Batches!$A:$A,"<>",Batches!$C:$C,"",Batches!$K:$K,">0")', 'calc', align='center')
    cell(ex, 'C10', '=SUMIFS(Batches!$N:$N,Batches!$A:$A,"<>",Batches!$C:$C,"",Batches!$K:$K,">0")', 'calc', INR)
    cell(ex, 'A11', 'Total stock value', 'soft', bold=True); cell(ex, 'C11', '=SUM(Batches!$N$4:$N$1503)', 'soft', INR, bold=True)
    ex.conditional_formatting.add('B5:C5', FormulaRule(formula=['$B$5>0'], fill=RED, font=RED_FONT))
    ex.conditional_formatting.add('B6:C6', FormulaRule(formula=['$B$6>0'], fill=AMBER))
    cell(ex, 'A13', 'Items at or below reorder level', 'box'); cell(ex, 'B13', '=COUNTIF(Items!$I:$I,"Reorder")', 'calc', align='center')
    ex.merge_cells('A15:C17')
    cell(ex, 'A15', 'Plan returns to suppliers for the 0 to 90 day bands in time; many suppliers accept returns only before the expiry month. '
         'Expired stock must be kept apart from saleable stock until it is returned or disposed of as your local rules require.', wrap=True, size=9, color=MUTED)
    protect(ex)
    save(wb, 'Pharmacy-Stock-Expiry-Register.xlsx')

# ------------------------------------------------------------------ Lab sample register
TESTS = [('CBC (complete blood count)', 'EDTA blood', 6, '', 350), ('ESR', 'EDTA blood', 3, '', 120), ('Blood sugar fasting', 'Fluoride blood', 2, '', 60),
         ('Blood sugar post-meal', 'Fluoride blood', 2, '', 60), ('HbA1c', 'EDTA blood', 24, '', 500), ('Lipid profile', 'Serum', 12, '', 600),
         ('Liver function test (LFT)', 'Serum', 12, '', 700), ('Kidney function test (KFT)', 'Serum', 12, '', 700), ('Thyroid profile (T3, T4, TSH)', 'Serum', 24, '', 550),
         ('Serum electrolytes', 'Serum', 12, '', 450), ('Uric acid', 'Serum', 12, '', 200), ('CRP', 'Serum', 6, '', 400),
         ('Vitamin D (25-OH)', 'Serum', 48, 'Reference lab', 1400), ('Vitamin B12', 'Serum', 48, 'Reference lab', 900),
         ('Urine routine', 'Urine', 4, '', 150), ('Urine culture', 'Urine (sterile container)', 72, 'Reference lab', 700),
         ('Stool routine', 'Stool', 6, '', 150), ('Blood group and Rh', 'EDTA blood', 2, '', 100), ('Dengue NS1 antigen', 'Serum', 6, '', 600),
         ('Malaria antigen', 'EDTA blood', 2, '', 400)]
STATUS = ['Collected', 'Sent to lab', 'Processing', 'Report ready', 'Delivered', 'Recollect needed']

def lab_register():
    wb = new_wb('lab-sample-register-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 26, 3, 60])
    ts = sheet(wb, 'Tests', [34, 24, 13, 18, 11])
    rg = sheet(wb, 'Register', [12, 11, 7, 22, 10, 5, 6, 13, 18, 30, 18, 10, 10, 15, 16, 12, 12, 16, 12], tab=BLUE)
    sm = sheet(wb, 'Summary', [34, 12, 14], tab=BLUE)
    title(st, 'Settings', 'Fill once. Other sheets read these values.')
    setting(st, 3, 'Lab / collection centre name', 'Sehat Diagnostics Collection Centre')
    setting(st, 4, 'Warn when a report is due within (hours)', 4)
    setting(st, 5, 'Month for the summary', '=DATE(YEAR(TODAY()),MONTH(TODAY()),1)', 'mmmm yyyy')
    st.merge_cells('D3:D6')
    cell(st, 'D3', 'Overdue and due-soon alerts use the computer\'s current date and time, so they update each time the file is opened. '
         'The register tracks samples and reports only; it does not store test results.', wrap=True, size=9, color=MUTED)
    cell(st, 'A8', PRIVATE, size=9, color=MUTED)
    title(ts, 'Tests', 'Turnaround = hours from collection until the report is usually ready. Change them to match your lab.')
    head(ts, 3, ['Test', 'Sample type', 'Turnaround (hours)', 'Outside lab (if sent out)', 'Charge'])
    grid(ts, 4, 103, 'ABCDE', {}, {'E': INR0})
    rows_in(ts, 4, TESTS)
    ts.freeze_panes = 'A4'
    title(rg, 'Sample register', PRIVATE)
    head(rg, 3, ['Sample ID', 'Collected on', 'Time', 'Patient name', 'Patient ID', 'Age', 'Sex', 'Phone', 'Referred by', 'Test', 'Sample type',
                 'Charge', 'Paid', 'Status', 'Report due', 'Report ready on', 'Delivered on', 'Delivered to', 'Alert'])
    R2 = 1503
    grid(rg, 4, R2, 'ABCDEFGHIJKLMNOPQRS', {'K': 'calc', 'O': 'calc', 'S': 'calc'},
         {'B': DATE, 'C': 'hh:mm', 'L': INR0, 'M': INR0, 'O': DT, 'P': DATE, 'Q': DATE},
         {'K': '=IF($J{r}="","",IFERROR(VLOOKUP($J{r},Tests!$A$4:$B$103,2,FALSE),""))',
          'L': '=IF($J{r}="","",IFERROR(VLOOKUP($J{r},Tests!$A$4:$E$103,5,FALSE),""))',
          'O': '=IF(OR($B{r}="",$J{r}=""),"",$B{r}+N($C{r})+IFERROR(VLOOKUP($J{r},Tests!$A$4:$C$103,3,FALSE),0)/24)',
          'S': '=IF(OR($J{r}="",$O{r}=""),"",IF(OR($N{r}="Report ready",$N{r}="Delivered"),"",IF($N{r}="Recollect needed","Recollect",IF(NOW()>$O{r},"Overdue",IF($O{r}-NOW()<=N(Settings!$B$4)/24,"Due soon","")))))'})
    t = TODAY; tm = datetime.time
    rows_in(rg, 4, [
        ('S-%s-001' % t.strftime('%d%m'), t - D(days=2), tm(8, 15), 'Ramesh Kumar', 'P-0001', 54, 'M', '98140 11111', 'Dr. A. Sharma', 'Lipid profile', None, None, 600, 'Delivered', None, t - D(days=2), t - D(days=1), 'Patient'),
        ('S-%s-002' % t.strftime('%d%m'), t - D(days=2), tm(8, 15), 'Ramesh Kumar', 'P-0001', 54, 'M', '98140 11111', 'Dr. A. Sharma', 'HbA1c', None, None, 500, 'Report ready', None, t - D(days=1)),
        ('S-%s-003' % t.strftime('%d%m'), t - D(days=1), tm(9, 0), 'Sunita Devi', 'P-0002', 47, 'F', '98140 22222', 'Dr. R. Kaur', 'Thyroid profile (T3, T4, TSH)', None, None, 550, 'Processing'),
        ('S-%s-004' % t.strftime('%d%m'), t - D(days=1), tm(9, 20), 'Mohit Arora', 'P-0005', 26, 'M', '98140 55555', 'Self', 'Vitamin D (25-OH)', None, None, 1400, 'Sent to lab'),
        ('S-%s-005' % t.strftime('%d%m'), t, tm(8, 40), 'Kiran Bala', 'P-0006', 39, 'F', '98140 66666', 'Dr. A. Sharma', 'CBC (complete blood count)', None, None, 350, 'Collected'),
        ('S-%s-006' % t.strftime('%d%m'), t, tm(8, 40), 'Kiran Bala', 'P-0006', 39, 'F', '98140 66666', 'Dr. A. Sharma', 'Dengue NS1 antigen', None, None, 300, 'Collected'),
        ('S-%s-007' % t.strftime('%d%m'), t, tm(10, 5), 'Harpreet Singh', 'P-0003', 31, 'M', '98140 33333', 'Dr. R. Kaur', 'Urine culture', None, None, 700, 'Sent to lab')])
    dv(rg, 'G4:G%d' % R2, '"M,F,Other"')
    dv(rg, 'J4:J%d' % R2, '=Tests!$A$4:$A$103', 'Pick the test from the Tests sheet')
    dv(rg, 'N4:N%d' % R2, '"%s"' % ','.join(STATUS))
    rg.conditional_formatting.add('A4:S%d' % R2, FormulaRule(formula=['$S4="Overdue"'], fill=RED, font=RED_FONT))
    rg.conditional_formatting.add('A4:S%d' % R2, FormulaRule(formula=['OR($S4="Due soon",$S4="Recollect")'], fill=AMBER))
    rg.conditional_formatting.add('N4:N%d' % R2, CellIsRule(operator='equal', formula=['"Delivered"'], fill=GREEN))
    rg.freeze_panes = 'E4'
    rg.auto_filter.ref = 'A3:S%d' % R2
    rg.page_setup.orientation = 'landscape'
    # Summary
    title(sm, 'Summary', '="Now: "&TEXT(NOW(),"dd-mm-yyyy hh:mm")')
    sm['A2'].value = '="Now: "&TEXT(NOW(),"dd-mm-yyyy hh:mm")'
    head(sm, 4, ['Samples by status', 'Count', ''])
    for i, s_ in enumerate(STATUS, 5):
        cell(sm, 'A%d' % i, s_, 'box'); cell(sm, 'B%d' % i, '=COUNTIF(Register!$N:$N,$A%d)' % i, 'calc', align='center')
    cell(sm, 'A11', 'Reports overdue now', 'soft', bold=True); cell(sm, 'B11', '=COUNTIF(Register!$S:$S,"Overdue")', 'soft', bold=True, align='center')
    cell(sm, 'A12', 'Reports due soon', 'box'); cell(sm, 'B12', '=COUNTIF(Register!$S:$S,"Due soon")', 'calc', align='center')
    sm.conditional_formatting.add('B11', CellIsRule(operator='greaterThan', formula=['0'], fill=RED, font=RED_FONT))
    head(sm, 14, ['Today', 'Count', 'Amount'])
    cell(sm, 'A15', 'Samples collected today', 'box'); cell(sm, 'B15', '=COUNTIF(Register!$B:$B,TODAY())', 'calc', align='center')
    cell(sm, 'C15', '=SUMIFS(Register!$M:$M,Register!$B:$B,TODAY())', 'calc', INR0)
    inm = 'Register!$B:$B,">="&Settings!$B$5,Register!$B:$B,"<="&EOMONTH(Settings!$B$5,0)'
    head(sm, 17, ['Tests this month', 'Count', 'Collected'])
    for i in range(40):
        r, s = 18 + i, 4 + i
        cell(sm, 'A%d' % r, '=IF(Tests!A%d="","",Tests!A%d)' % (s, s), 'box')
        cell(sm, 'B%d' % r, '=IF($A{r}="","",COUNTIFS(Register!$J:$J,$A{r},{m}))'.format(r=r, m=inm), 'calc', align='center')
        cell(sm, 'C%d' % r, '=IF($A{r}="","",SUMIFS(Register!$M:$M,Register!$J:$J,$A{r},{m}))'.format(r=r, m=inm), 'calc', INR0)
    cell(sm, 'A58', 'Month total', 'soft', bold=True)
    cell(sm, 'B58', '=SUM(B18:B57)', 'soft', bold=True, align='center'); cell(sm, 'C58', '=SUM(C18:C57)', 'soft', INR0, bold=True)
    protect(sm)
    save(wb, 'Lab-Sample-Register.xlsx')

# ------------------------------------------------------------------ Duty roster
SHIFTS = [('M', 'Morning', '08:00 to 14:00', 6, 2), ('E', 'Evening', '14:00 to 20:00', 6, 2), ('N', 'Night', '20:00 to 08:00', 12, 1),
          ('G', 'General', '09:00 to 17:00', 8, 0), ('OFF', 'Weekly off', '', 0, None), ('L', 'Leave', '', 0, None)]
STAFF = [('Nisha Thomas', 'Nurse', 'Mon'), ('Priya Rani', 'Nurse', 'Tue'), ('Jaspreet Kaur', 'Nurse', 'Wed'), ('Anil Masih', 'Nurse', 'Thu'),
         ('Rekha Sharma', 'Receptionist', 'Sun'), ('Deepak Kumar', 'Receptionist', 'Fri'), ('Vikas Jain', 'Pharmacist', 'Sun'),
         ('Sonia Gill', 'Lab technician', 'Sat'), ('Ramu Prasad', 'Attendant', 'Mon'), ('Kamla Devi', 'Housekeeping', 'Sat')]

def duty_roster():
    wb = new_wb('clinic-duty-roster-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [12, 16, 18, 10, 14, 3, 50])
    sf = sheet(wb, 'Staff', [26, 18, 15, 12])
    ws = sheet(wb, 'Roster', [22, 13] + [4.6] * 31 + [5, 5, 5, 5, 5, 5, 7], tab=BLUE)
    ws.page_setup.orientation = 'landscape'
    title(st, 'Settings', 'Fill once. Other sheets read these values.')
    label(st, 'A3', 'Clinic'); st.merge_cells('B3:E3'); cell(st, 'B3', 'Sehat Nursing Home', 'in')
    label(st, 'A4', 'Month'); st.merge_cells('B4:E4'); cell(st, 'B4', MONTH1, 'in', 'mmmm yyyy', align='left')
    head(st, 6, ['Code', 'Shift', 'Timing', 'Hours', 'Minimum staff'])
    for i in range(8):
        r = 7 + i
        row = SHIFTS[i] if i < len(SHIFTS) else (None,) * 5
        for col, v, fmt in zip('ABCDE', row, [None, None, None, '0.0', None]):
            cell(st, '%s%d' % (col, r), v, 'in', fmt, align='center' if col in 'ADE' else None)
    st.merge_cells('G6:G12')
    cell(st, 'G6', 'Codes in the first three rows are counted as morning, evening and night on the Roster sheet; a day turns red when fewer people '
         'than the minimum are on that shift. Hours are added up for each person. You can rename codes or add two more in the empty rows.', wrap=True, size=9, color=MUTED)
    title(sf, 'Staff', 'Up to 40 staff. Phone numbers are personal: don\'t print this sheet for a public board.')
    head(sf, 3, ['Name', 'Role', 'Phone', 'Weekly off'])
    grid(sf, 4, 43, 'ABCD', {})
    rows_in(sf, 4, [(n, r, '98140 %05d' % (10000 + 1111 * i), o) for i, (n, r, o) in enumerate(STAFF)])
    dv(sf, 'B4:B43', '"Doctor,Nurse,Receptionist,Pharmacist,Lab technician,Attendant,Housekeeping,Security,Driver,Other"')
    dv(sf, 'D4:D43', '"Mon,Tue,Wed,Thu,Fri,Sat,Sun,Rotating"')
    # Roster
    title(ws, 'Staff duty roster', 'Shift codes and minimum staff are on the Settings sheet.')
    ws['A2'] = '=Settings!B3&"   ·   "&TEXT(Settings!B4,"mmmm yyyy")'; ws['A2'].font = Font(size=11, bold=True, color=INK)
    ws['A3'] = '=TRIM(Settings!A7&" "&Settings!C7&"   "&Settings!A8&" "&Settings!C8&"   "&Settings!A9&" "&Settings!C9&"   "&Settings!A10&" "&Settings!C10&"   "&Settings!A11&" = "&Settings!B11&"   "&Settings!A12&" = "&Settings!B12)'
    ws['A3'].font = Font(size=9, color=MUTED)
    for n in range(31):
        col = L(3 + n)
        cell(ws, '%s5' % col, '=IF(%d<=DAY(EOMONTH(Settings!$B$4,0)),Settings!$B$4+%d,"")' % (n + 1, n), 'calc', 'd', align='center', bold=True)
        cell(ws, '%s6' % col, '=IF(%s5="","",LEFT(TEXT(%s5,"ddd"),2))' % (col, col), 'calc', align='center', size=8)
    head(ws, 5, ['Name', 'Role']); ws.merge_cells('A5:A6'); ws.merge_cells('B5:B6')
    tails = ['AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
    for k, col in enumerate(tails):
        c = ws['%s5' % col]
        c.value = ('=Settings!A%d' % (7 + k)) if k < 6 else 'Hours'
        c.font, c.fill, c.alignment, c.border = Font(bold=True, color='FFFFFF', size=9), F_HEAD, Alignment(horizontal='center', vertical='center'), BOX
        ws.merge_cells('%s5:%s6' % (col, col))
    R1, R2 = 7, 46
    for i in range(40):
        r, s = R1 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Staff!A%d="","",Staff!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF(Staff!B%d="","",Staff!B%d)' % (s, s), 'box', size=9)
        for n in range(31): cell(ws, '%s%d' % (L(3 + n), r), None, 'in', align='center')
        for k, col in enumerate(tails[:6]):
            cell(ws, '%s%d' % (col, r), '=IF(OR($A{r}="",Settings!$A${s}=""),"",COUNTIF($C{r}:$AG{r},Settings!$A${s}))'.format(r=r, s=7 + k), 'calc', align='center', size=9)
        cell(ws, 'AN%d' % r, '=IF($A{r}="","",SUMPRODUCT(COUNTIF($C{r}:$AG{r},Settings!$A$7:$A$14)*Settings!$D$7:$D$14))'.format(r=r), 'calc', '0', align='center', bold=True)
    # sample month: a fair rotation that respects each person's weekly off
    days = (datetime.date(MONTH1.year + MONTH1.month // 12, MONTH1.month % 12 + 1, 1) - MONTH1).days
    WD = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    nurse_cycle = ['M', 'M', 'E', 'E', 'N', 'N', 'M']
    for i, (n, role, off) in enumerate(STAFF):
        for dd in range(days):
            dte = MONTH1 + D(days=dd)
            if WD[dte.weekday()] == off: code = 'OFF'
            elif role == 'Nurse': code = nurse_cycle[(dd + i * 2) % 7]
            elif role == 'Receptionist': code = 'M' if (dd // 7 + i) % 2 == 0 else 'E'
            else: code = 'G'
            ws.cell(R1 + i, 3 + dd, code)
    ws.cell(R1 + 1, 3 + 12, 'L'); ws.cell(R1 + 1, 3 + 13, 'L')
    for k, (r, lab_) in enumerate([(48, 'On morning'), (49, 'On evening'), (50, 'On night')]):
        cell(ws, 'A%d' % r, '="%s ("&Settings!A%d&", min "&N(Settings!E%d)&")"' % (lab_, 7 + k, 7 + k), 'soft', bold=True, size=9)
        ws.merge_cells('A%d:B%d' % (r, r))
        for n in range(31):
            col = L(3 + n)
            cell(ws, '%s%d' % (col, r), '=IF(OR(%s$5="",Settings!$A$%d=""),"",COUNTIF(%s$%d:%s$%d,Settings!$A$%d))' % (col, 7 + k, col, R1, col, R2, 7 + k), 'soft', align='center', size=8)
        ws.conditional_formatting.add('C%d:AG%d' % (r, r), FormulaRule(formula=['AND(C%d<>"",C%d<N(Settings!$E$%d))' % (r, r, 7 + k)], fill=RED, font=RED_FONT))
    dv(ws, 'C%d:AG%d' % (R1, R2), '=Settings!$A$7:$A$14')
    ws.conditional_formatting.add('C5:AG50', FormulaRule(formula=['C$5=""'], fill=PatternFill('solid', fgColor='CBD5E1')))
    ws.conditional_formatting.add('C5:AG46', FormulaRule(formula=['AND(C$5<>"",WEEKDAY(C$5)=1)'], fill=PatternFill('solid', fgColor='E2E8F0')))
    ws.conditional_formatting.add('C%d:AG%d' % (R1, R2), CellIsRule(operator='equal', formula=['Settings!$A$9'], fill=PatternFill('solid', fgColor='DBEAFE'), font=Font(bold=True, color='1E3A8A')))
    ws.conditional_formatting.add('C%d:AG%d' % (R1, R2), CellIsRule(operator='equal', formula=['Settings!$A$12'], fill=AMBER))
    ws.freeze_panes = 'C7'
    ws.print_title_rows = '5:6'
    ws.print_area = 'A1:AN50'
    protect(ws)
    save(wb, 'Clinic-Staff-Duty-Roster.xlsx')

# ------------------------------------------------------------------ Equipment maintenance and calibration
EQUIP_TYPES = ['Preventive maintenance', 'Breakdown repair', 'Calibration', 'Installation', 'Inspection']

def equipment_log():
    wb = new_wb('equipment-maintenance-log-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 22, 3, 60])
    eq = sheet(wb, 'Equipment', [9, 26, 22, 14, 14, 20, 13, 11, 11, 11, 11, 12, 12, 9, 11, 11, 22], tab=BLUE)
    lg = sheet(wb, 'Service log', [11, 9, 26, 22, 22, 18, 11, 10, 30], tab=BLUE)
    dl = sheet(wb, 'Due list', [34, 12, 14], tab=BLUE)
    title(st, 'Settings', 'Fill once. Other sheets read these values.')
    setting(st, 3, 'Clinic / lab name', 'Sehat Nursing Home')
    setting(st, 4, 'Position as on', '=TODAY()', DATE)
    setting(st, 5, 'Warn this many days before', 30)
    st.merge_cells('D3:D7')
    cell(st, 'D3', 'Calibration intervals come from the maker\'s manual or your accreditation body. This register only reminds you; '
         'keep the calibration certificates on file.', wrap=True, size=9, color=MUTED)
    title(eq, 'Equipment register', 'One row per machine. Grey columns calculate from the Service log.')
    head(eq, 3, ['Asset ID', 'Equipment', 'Make / model', 'Serial no.', 'Location', 'Vendor / service engineer', 'Vendor phone', 'Purchased on',
                 'Warranty ends', 'AMC ends', 'Calibrate every (months)', 'Last calibration', 'Next calibration', 'Days left', 'Last service', 'Cost to date', 'Status'])
    LG = "'Service log'"
    R2 = 203
    grid(eq, 4, R2, 'ABCDEFGHIJKLMNOPQ', {c: 'calc' for c in 'LMNOPQ'}, {c: DATE for c in 'HIJLMO'} | {'P': INR0},
         {'L': '=IF($A{r}="","",IFERROR(1/(1/MAX(IF((%s!$B$4:$B$2003=$A{r})*(%s!$D$4:$D$2003="Calibration"),%s!$A$4:$A$2003))),""))' % (LG, LG, LG),
          'M': '=IF(OR($A{r}="",N($K{r})=0),"",IF($L{r}<>"",EDATE($L{r},$K{r}),IF($H{r}<>"",EDATE($H{r},$K{r}),"")))',
          'N': '=IF($M{r}="","",$M{r}-Settings!$B$4)',
          'O': '=IF($A{r}="","",IFERROR(1/(1/MAX(IF(%s!$B$4:$B$2003=$A{r},%s!$A$4:$A$2003))),""))' % (LG, LG),
          'P': '=IF($A{r}="","",SUMIFS(%s!$G:$G,%s!$B:$B,$A{r}))' % (LG, LG),
          'Q': ('=IF($A{r}="","",IF(AND(N($K{r})>0,$M{r}=""),"Add a calibration",IF(AND($M{r}<>"",N($N{r})<0),"Calibration overdue",'
                'IF(AND($M{r}<>"",N($N{r})<=Settings!$B$5),"Calibration due soon",'
                'IF(IF($J{r}<>"",$J{r},IF($I{r}<>"",$I{r},Settings!$B$4+99999))-Settings!$B$4<=Settings!$B$5,"AMC / warranty ending","OK")))))')})
    arrayize(eq, 'L', 4, R2); arrayize(eq, 'O', 4, R2)
    t = TODAY
    rows_in(eq, 4, [('EQ-01', 'Autoclave (vertical, 40 L)', 'SteriPro VA-40', 'SP40-22187', 'CSSD room', 'MedServe Engineers', '98140 70001', t - D(days=900), t - D(days=170), t + D(days=20), 12),
                    ('EQ-02', 'Digital BP monitor', 'CareTrack BP-200', 'BP2-55102', 'OPD room 1', 'City Medical Equipment', '98140 70002', t - D(days=500), t + D(days=230), None, 12),
                    ('EQ-03', 'Weighing scale (adult)', 'WeighRight AW-150', 'WR-88231', 'Reception', 'City Medical Equipment', '98140 70002', t - D(days=700), None, None, 12),
                    ('EQ-04', 'ECG machine, 12 channel', 'CardioLine C12', 'CL12-4410', 'Procedure room', 'MedServe Engineers', '98140 70001', t - D(days=400), t + D(days=330), None, 12),
                    ('EQ-05', 'Semi-auto biochemistry analyser', 'LabChem SA-5', 'LC5-1092', 'Laboratory', 'LabChem Service Centre', '98140 70003', t - D(days=1200), None, t + D(days=160), 6),
                    ('EQ-06', 'Vaccine refrigerator (ILR)', 'ColdSafe ILR-90', 'CS90-7731', 'Pharmacy', 'CoolTech Services', '98140 70004', t - D(days=650), t - D(days=285), None, 12),
                    ('EQ-07', 'Pulse oximeter', 'OxyCheck P10', 'OC10-2209', 'Ward', 'City Medical Equipment', '98140 70002', t - D(days=150), t + D(days=215), None, None),
                    ('EQ-08', 'Diesel generator 15 kVA', 'PowerGen 15', 'PG15-3306', 'Back yard', 'Ludhiana Power Services', '98140 70005', t - D(days=1500), None, t + D(days=95), None)])
    eq.conditional_formatting.add('Q4:Q%d' % R2, CellIsRule(operator='equal', formula=['"Calibration overdue"'], fill=RED, font=RED_FONT))
    eq.conditional_formatting.add('Q4:Q%d' % R2, FormulaRule(formula=['OR(Q4="Calibration due soon",Q4="AMC / warranty ending",Q4="Add a calibration")'], fill=AMBER))
    eq.conditional_formatting.add('Q4:Q%d' % R2, CellIsRule(operator='equal', formula=['"OK"'], fill=GREEN))
    eq.freeze_panes = 'C4'
    eq.auto_filter.ref = 'A3:Q%d' % R2
    eq.page_setup.orientation = 'landscape'
    title(lg, 'Service log', 'One row per service, repair, calibration or inspection.')
    head(lg, 3, ['Date', 'Asset ID', 'Equipment', 'Type', 'Done by', 'Certificate / job no.', 'Cost', 'Downtime (days)', 'Remarks'])
    grid(lg, 4, 2003, 'ABCDEFGHI', {'C': 'calc'}, {'A': DATE, 'G': INR0},
         {'C': '=IF($B{r}="","",IFERROR(INDEX(Equipment!$B:$B,MATCH($B{r},Equipment!$A:$A,0)),"Not in register"))'})
    rows_in(lg, 4, [(t - D(days=350), 'EQ-01', None, 'Calibration', 'MedServe Engineers', 'CAL/23/0911', 2500, 0, 'Temperature and pressure gauges checked'),
                    (t - D(days=120), 'EQ-01', None, 'Preventive maintenance', 'MedServe Engineers', 'PM-4471', 0, 0, 'Under AMC: gasket replaced'),
                    (t - D(days=200), 'EQ-02', None, 'Calibration', 'City Medical Equipment', 'CAL/BP/220', 600, 0, ''),
                    (t - D(days=380), 'EQ-03', None, 'Calibration', 'Legal metrology / vendor', 'WM-7781', 400, 0, 'Stamped'),
                    (t - D(days=40), 'EQ-04', None, 'Breakdown repair', 'MedServe Engineers', 'BR-1209', 3200, 3, 'Patient cable replaced'),
                    (t - D(days=170), 'EQ-05', None, 'Calibration', 'LabChem Service Centre', 'LC/CAL/556', 4500, 1, 'Wavelength and temperature checks'),
                    (t - D(days=60), 'EQ-06', None, 'Inspection', 'In-house', '', 0, 0, 'Temperature log reviewed'),
                    (t - D(days=30), 'EQ-08', None, 'Preventive maintenance', 'Ludhiana Power Services', 'PM-889', 1800, 0, 'Oil and filter change')])
    dv(lg, 'B4:B2003', '=Equipment!$A$4:$A$%d' % R2, 'Pick the asset ID from the Equipment sheet')
    dv(lg, 'D4:D2003', '"%s"' % ','.join(EQUIP_TYPES))
    lg.freeze_panes = 'A4'
    title(dl, 'Due list', '="As on "&TEXT(Settings!B4,"dd-mm-yyyy")')
    dl['A2'].value = '="As on "&TEXT(Settings!B4,"dd-mm-yyyy")'
    head(dl, 4, ['Status', 'Machines', ''])
    for i, s_ in enumerate(['Calibration overdue', 'Calibration due soon', 'Add a calibration', 'AMC / warranty ending', 'OK'], 5):
        cell(dl, 'A%d' % i, s_, 'box'); cell(dl, 'B%d' % i, '=COUNTIF(Equipment!$Q:$Q,$A%d)' % i, 'calc', align='center')
    dl.conditional_formatting.add('B5', CellIsRule(operator='greaterThan', formula=['0'], fill=RED, font=RED_FONT))
    dl.conditional_formatting.add('B6:B8', CellIsRule(operator='greaterThan', formula=['0'], fill=AMBER))
    head(dl, 11, ['Service log totals', 'Entries', 'Amount'])
    cell(dl, 'A12', 'Maintenance cost, all time', 'box'); cell(dl, 'B12', "=COUNTA(%s!$B$4:$B$2003)" % LG, 'calc', align='center')
    cell(dl, 'C12', '=SUM(%s!$G$4:$G$2003)' % LG, 'calc', INR0)
    cell(dl, 'A13', 'Maintenance cost, last 12 months', 'box')
    cell(dl, 'B13', '=COUNTIFS(%s!$A:$A,">"&EDATE(Settings!$B$4,-12))' % LG, 'calc', align='center')
    cell(dl, 'C13', '=SUMIFS(%s!$G:$G,%s!$A:$A,">"&EDATE(Settings!$B$4,-12))' % (LG, LG), 'calc', INR0)
    cell(dl, 'A14', 'Breakdown downtime, last 12 months (days)', 'box')
    cell(dl, 'B14', '=SUMIFS(%s!$H:$H,%s!$A:$A,">"&EDATE(Settings!$B$4,-12),%s!$D:$D,"Breakdown repair")' % (LG, LG, LG), 'calc', align='center')
    dl.merge_cells('A16:C18')
    cell(dl, 'A16', 'Open the Equipment sheet and use the filter on the Status column to list the machines behind each count. '
         'Sort by Days left to see the next calibrations first.', wrap=True, size=9, color=MUTED)
    protect(dl)
    save(wb, 'Equipment-Maintenance-Calibration-Log.xlsx')

MAKERS = [opd_register, billing_register, pharmacy_stock, lab_register, duty_roster, equipment_log]
