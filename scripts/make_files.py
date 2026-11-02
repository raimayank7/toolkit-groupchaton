"""Makes every downloadable Excel / Word file into files/ (committed; the website links to them).
   python -X utf8 scripts/make_files.py          then   powershell -File scripts/check_office.ps1
   check_office.ps1 opens each workbook in real Excel, recalculates, fails on any #VALUE!/#REF!/… and saves the
   calculated values (so previews and viewers that don't calculate show real numbers).
   Text for the "How to use" sheets comes from content/catalog.py."""
import os, sys, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, 'content'))
from catalog import T, KIT
OUT = os.path.join(HERE, 'files')
SITE = 'https://toolkit.groupchaton.com'
ITEM = {t['id']: t for t in T}
CREDIT = 'Designed & Developed by: Nivesh Sahayak'
NS_URL = 'https://www.niveshsahayak.in/'

BLUE, INK, MUTED = '2563EB', '14213D', '55607A'
F_HEAD = PatternFill('solid', fgColor=BLUE)
F_IN = PatternFill('solid', fgColor='FFF7D6')      # yellow = type here
F_CALC = PatternFill('solid', fgColor='F1F5F9')    # grey = calculates itself
F_SOFT = PatternFill('solid', fgColor='E8F0FE')
THIN = Side(style='thin', color='CBD5E1')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INR = '[>=10000000]₹ ##\\,##\\,##\\,##0.00;[>=100000]₹ ##\\,##\\,##0.00;₹ ##,##0.00'
INR0 = '[>=10000000]₹ ##\\,##\\,##\\,##0;[>=100000]₹ ##\\,##\\,##0;₹ ##,##0'
INR_TXT = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'   # inside TEXT()
DATE = 'dd-mm-yyyy'
TODAY = datetime.date.today()
MONTH1 = datetime.date(TODAY.year, TODAY.month, 1)
STATES = ['Andaman and Nicobar Islands', 'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chandigarh', 'Chhattisgarh',
          'Dadra and Nagar Haveli and Daman and Diu', 'Delhi', 'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir',
          'Jharkhand', 'Karnataka', 'Kerala', 'Ladakh', 'Lakshadweep', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya',
          'Mizoram', 'Nagaland', 'Odisha', 'Puducherry', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura',
          'Uttar Pradesh', 'Uttarakhand', 'West Bengal', 'Other Territory']
ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen',
        'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
WORDS = [ONES[n] if n < 20 else TENS[n // 10] + ('-' + ONES[n % 10] if n % 10 else '') for n in range(100)]

# ------------------------------------------------------------------ workbook helpers
def new_wb(tid):
    wb = Workbook()
    t = ITEM[tid]
    p = wb.properties
    p.title, p.creator, p.lastModifiedBy = t['name'], 'Nivesh Sahayak', 'Nivesh Sahayak'
    p.subject, p.keywords = t['desc'], 'Toolkit by GroupChatOn; ' + CREDIT
    p.description = '%s. From Toolkit by GroupChatOn (%s). %s (%s)' % (t['name'], SITE, CREDIT, NS_URL)
    p.category = KIT[t['kit']]['name']
    wb._tid = tid
    return wb

def sheet(wb, title, widths, first=False, tab=None):
    ws = wb.active if first else wb.create_sheet(title)
    ws.title = title
    for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w
    ws.sheet_view.showGridLines = False
    if tab: ws.sheet_properties.tabColor = tab
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.oddFooter.center.text = 'Toolkit by GroupChatOn  ·  toolkit.groupchaton.com'
    ws.oddFooter.center.size = 8
    return ws

def title(ws, text, sub=None, span=8):
    ws['A1'] = text
    ws['A1'].font = Font(size=16, bold=True, color=BLUE)
    if sub:
        ws['A2'] = sub
        ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    ws.row_dimensions[1].height = 26

def head(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row, col + i, lab)
        c.font, c.fill, c.border = Font(bold=True, color='FFFFFF'), F_HEAD, BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = 32

def cell(ws, ref, value=None, kind=None, fmt=None, bold=False, align=None, wrap=False, size=None, color=None):
    c = ws[ref]
    if value is not None: c.value = value
    if kind == 'in': c.fill, c.protection = F_IN, Protection(locked=False)
    elif kind == 'calc': c.fill = F_CALC
    elif kind == 'soft': c.fill = F_SOFT
    if kind in ('in', 'calc', 'box', 'soft'): c.border = BOX
    if fmt: c.number_format = fmt
    if bold or size or color: c.font = Font(bold=bold, size=size or 11, color=color or INK)
    if align or wrap: c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    return c

def label(ws, ref, text, bold=True):
    return cell(ws, ref, text, bold=bold, color=INK)

def grid(ws, r1, r2, cols, kinds, fmts=None, formulas=None):
    """Data rows r1..r2. kinds[col] = 'in' | 'calc'; formulas[col] = template with {r} for the row number."""
    for r in range(r1, r2 + 1):
        for col in cols:
            ref = '%s%d' % (col, r)
            f = (formulas or {}).get(col)
            cell(ws, ref, f.format(r=r) if f else None, kinds.get(col, 'in'), (fmts or {}).get(col))

def dv(ws, rng, source, prompt=None):
    v = DataValidation(type='list', formula1=source, allow_blank=True, showErrorMessage=False)
    if prompt: v.promptTitle, v.prompt, v.showInputMessage = 'Pick from the list', prompt, True
    ws.add_data_validation(v)
    v.add(rng)

def protect(ws):
    p = ws.protection
    p.sheet = True
    p.formatCells = p.formatColumns = p.formatRows = p.sort = p.autoFilter = p.selectLockedCells = p.selectUnlockedCells = False

def words_formula(ref, lists='Lists'):
    """Rupees in words, Indian style (crore, lakh, thousand), from a helper list of 0-99 in Lists!A1:A100."""
    W = '%s!$A$1:$A$100' % lists
    n = 'INT(%s)' % ref
    part = lambda expr, word: 'IF(%s>0,INDEX(%s,%s+1)&"%s","")' % (expr, W, expr, word)
    cr, lk = 'INT(%s/10000000)' % n, 'INT(MOD(%s,10000000)/100000)' % n
    th, hu, re_ = 'INT(MOD(%s,100000)/1000)' % n, 'INT(MOD(%s,1000)/100)' % n, 'MOD(%s,100)' % n
    pa = 'MIN(99,ROUND(MOD(%s,1)*100,0))' % ref
    body = '&'.join([part(cr, ' Crore '), part(lk, ' Lakh '), part(th, ' Thousand '), part(hu, ' Hundred '), part(re_, '')])
    return ('=IF(N(%s)<=0,"",'
            '"Rupees "&IF(%s=0,"Zero",TRIM(%s))&IF(%s>0," and "&INDEX(%s,%s+1)&" Paise","")&" Only")') % (ref, n, body, pa, W, pa)

def lists_sheet(wb, states=False, extra=None):
    ws = wb.create_sheet('Lists')
    for i, w in enumerate(WORDS, 1): ws.cell(i, 1, w)
    if states:
        for i, s in enumerate(STATES, 1): ws.cell(i, 3, s)
    for col, vals in (extra or {}).items():
        for i, v in enumerate(vals, 1): ws.cell(i, col, v)
    ws.sheet_state = 'hidden'
    return ws

def howto(wb, first=False):
    t = ITEM[wb._tid]
    ws = sheet(wb, 'How to use', [4, 100], first=first, tab='16A34A')
    title(ws, t['name'], 'Toolkit by GroupChatOn · free template · ' + SITE)
    r = 4
    def para(text, bold=False, size=11, color=INK, h=None):
        nonlocal r
        c = ws.cell(r, 2, text); c.font = Font(bold=bold, size=size, color=color); c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[r].height = h or max(18, 15 * (1 + len(text) // 95)); r += 1
    para(t['lead'])
    r += 1
    para('Colours', bold=True, size=12, color=BLUE)
    ws.cell(r, 1).fill = F_IN; para('Yellow cells: type or pick here.')
    ws.cell(r, 1).fill = F_CALC; para('Grey and white cells: calculate by themselves. Don\'t type over them.')
    r += 1
    para('Steps', bold=True, size=12, color=BLUE)
    for i, s in enumerate(t['steps'], 1):
        ws.cell(r, 1, i).font = Font(bold=True, color=BLUE); para(s)
    r += 1
    para('Tips', bold=True, size=12, color=BLUE)
    for s in t['tips']:
        ws.cell(r, 1, '•'); para(s)
    r += 1
    para('Full guide, questions and answers, and more free templates:', bold=True)
    c = ws.cell(r, 2, '%s/%s/%s/' % (SITE, t['kit'], t['id'])); c.hyperlink = c.value; c.font = Font(color=BLUE, underline='single'); r += 2
    c = ws.cell(r, 2, CREDIT); c.hyperlink = NS_URL; c.font = Font(bold=True, color=INK); r += 1
    para('Free to use for your own work and your clients\' work. Please don\'t sell or re-publish this file. '
         'The calculations are a help, not professional advice: check figures, tax rates and rules before you rely on them.', color=MUTED, size=9)
    ws.sheet_view.showGridLines = False
    return ws

def save(wb, name):
    if 'Sheet' in wb.sheetnames: del wb['Sheet']
    wb.active = 0
    wb.save(os.path.join(OUT, name))
    print('  ', name)

# ------------------------------------------------------------------ Accountants & GST
def gst_invoice():
    wb = new_wb('gst-invoice-excel')
    inv = sheet(wb, 'Invoice', [5, 28, 10, 7, 7, 11, 7, 14, 7, 13], first=True, tab=BLUE)
    biz = sheet(wb, 'Business', [24, 50])
    cus = sheet(wb, 'Customers', [28, 18, 40, 24, 14, 26])
    itm = sheet(wb, 'Items', [30, 12, 8, 12, 8])
    lists_sheet(wb, states=True)
    # Business
    title(biz, 'Your business details', 'Fill once. The invoice picks these up.')
    rows = [('Business name', 'Sharma Traders'), ('Address', '12, Main Bazaar, Ludhiana, Punjab 141001'), ('GSTIN', '03ABCDE1234F1Z5'),
            ('State', 'Punjab'), ('Phone', '98140 00000'), ('Email', 'accounts@example.com'), ('Bank name', 'State Bank of India'),
            ('Account number', '00000000000'), ('IFSC', 'SBIN0000000')]
    for i, (k, v) in enumerate(rows, 3):
        label(biz, 'A%d' % i, k); cell(biz, 'B%d' % i, v, 'in')
    dv(biz, 'B6', '=Lists!$C$1:$C$37')
    # Customers
    title(cus, 'Customers', 'One row per customer. State decides CGST + SGST or IGST.')
    head(cus, 3, ['Customer name', 'GSTIN (if registered)', 'Address', 'State', 'Phone', 'Email'])
    grid(cus, 4, 203, 'ABCDEF', {})
    for i, row in enumerate([('Gupta Enterprises', '03AAACG1111A1Z2', 'Model Town, Ludhiana', 'Punjab', '98760 11111', ''),
                             ('Delhi Office Supplies', '07AAACD2222B1Z3', 'Karol Bagh, New Delhi', 'Delhi', '98110 22222', ''),
                             ('Walk-in customer', '', '', 'Punjab', '', '')], 4):
        for j, v in enumerate(row): cus.cell(i, j + 1, v)
    dv(cus, 'D4:D203', '=Lists!$C$1:$C$37')
    cus.freeze_panes = 'A4'
    # Items
    title(itm, 'Items and services', 'Rate before GST. GST % as a number, e.g. 18.')
    head(itm, 3, ['Item / service', 'HSN / SAC', 'Unit', 'Rate (₹)', 'GST %'])
    grid(itm, 4, 303, 'ABCDE', {}, {'D': INR})
    for i, row in enumerate([('A4 copier paper (500 sheets)', '4802', 'Ream', 260, 18), ('Ball pen (box of 10)', '9608', 'Box', 90, 18),
                             ('Printer toner cartridge', '8443', 'Pcs', 1850, 18), ('Annual maintenance service', '998713', 'Job', 4500, 18),
                             ('Notebooks (dozen)', '4820', 'Doz', 480, 5)], 4):
        for j, v in enumerate(row): itm.cell(i, j + 1, v)
    itm.freeze_panes = 'A4'
    # Invoice
    ws = inv
    for r, f, sz in ((1, '=Business!B3', 18), (2, '=Business!B4', 10), (3, '="GSTIN: "&Business!B5&"   ·   State: "&Business!B6&"   ·   Phone: "&Business!B7', 10)):
        ws.merge_cells('A%d:J%d' % (r, r)); cell(ws, 'A%d' % r, f, bold=(r == 1), size=sz, align='center', color=BLUE if r == 1 else INK)
    ws.merge_cells('A5:J5'); c = cell(ws, 'A5', 'TAX INVOICE', bold=True, size=13, align='center', color='FFFFFF'); c.fill = F_HEAD
    def lab(ref, text, span):
        ws.merge_cells(span); label(ws, ref, text)
    lab('A7', 'Invoice No.', 'A7:B7'); ws.merge_cells('C7:E7'); cell(ws, 'C7', 'INV-001', 'in')
    lab('F7', 'Date', 'F7:G7'); ws.merge_cells('H7:J7'); cell(ws, 'H7', TODAY, 'in', DATE, align='left')
    lab('A8', 'Bill to', 'A8:B8'); ws.merge_cells('C8:E8'); cell(ws, 'C8', 'Gupta Enterprises', 'in')
    look = lambda col: '=IFERROR(IF($C$8="","",INDEX(Customers!$%s:$%s,MATCH($C$8,Customers!$A:$A,0))&""),"")' % (col, col)
    lab('A9', 'Address', 'A9:B9'); ws.merge_cells('C9:E9'); cell(ws, 'C9', look('C'), wrap=True)
    lab('A10', 'Phone', 'A10:B10'); ws.merge_cells('C10:E10'); cell(ws, 'C10', look('E'))
    lab('F8', 'Place of supply', 'F8:G8'); ws.merge_cells('H8:J8'); cell(ws, 'H8', look('D'))
    lab('F9', 'Buyer GSTIN', 'F9:G9'); ws.merge_cells('H9:J9'); cell(ws, 'H9', look('B'))
    lab('F10', 'Supply type', 'F10:G10'); ws.merge_cells('H10:J10')
    cell(ws, 'H10', '=IF($H$8="","",IF($H$8=Business!$B$6,"Intra-state: CGST + SGST","Inter-state: IGST"))')
    ws.row_dimensions[9].height = 30
    dv(ws, 'C8', '=Customers!$A$4:$A$203', 'Pick a customer from the Customers sheet')
    head(ws, 12, ['#', 'Item / service', 'HSN/SAC', 'Qty', 'Unit', 'Rate', 'Disc %', 'Taxable value', 'GST %', 'GST amount'])
    il = lambda col, r: '=IF($B%d="","",IFERROR(INDEX(Items!$%s:$%s,MATCH($B%d,Items!$A:$A,0)),""))' % (r, col, col, r)
    for i, r in enumerate(range(13, 23), 1):
        cell(ws, 'A%d' % r, i, 'box', align='center')
        cell(ws, 'B%d' % r, None, 'in'); cell(ws, 'D%d' % r, None, 'in'); cell(ws, 'G%d' % r, None, 'in')
        cell(ws, 'C%d' % r, il('B', r), 'box'); cell(ws, 'E%d' % r, il('C', r), 'box'); cell(ws, 'F%d' % r, il('D', r), 'box', INR)
        cell(ws, 'I%d' % r, il('E', r), 'box')
        cell(ws, 'H%d' % r, '=IF(OR($B{r}="",$D{r}=""),"",ROUND($D{r}*N($F{r})*(1-N($G{r})/100),2))'.format(r=r), 'box', INR)
        cell(ws, 'J%d' % r, '=IF($H{r}="","",ROUND($H{r}*N($I{r})/100,2))'.format(r=r), 'box', INR)
    ws['B13'], ws['D13'], ws['B14'], ws['D14'], ws['G14'] = 'Printer toner cartridge', 2, 'A4 copier paper (500 sheets)', 10, 5
    dv(ws, 'B13:B22', '=Items!$A$4:$A$303', 'Pick an item from the Items sheet')
    intra = '($H$8=Business!$B$6)'
    tot = [('Taxable value', '=SUM(H13:H22)'), ('CGST', '=IF(%s,ROUND(SUM(J13:J22)/2,2),0)' % intra), ('SGST', '=IF(%s,ROUND(SUM(J13:J22)/2,2),0)' % intra),
           ('IGST', '=IF(%s,0,SUM(J13:J22))' % intra), ('Round off', '=ROUND(J29-SUM(J24:J27),2)'), ('Grand total', '=ROUND(SUM(J24:J27),0)')]
    for i, (k, f) in enumerate(tot, 24):
        ws.merge_cells('G%d:I%d' % (i, i)); label(ws, 'G%d' % i, k, bold=(k == 'Grand total')); cell(ws, 'J%d' % i, f, 'calc', INR, bold=(k == 'Grand total'))
    ws['J29'].fill = F_SOFT
    label(ws, 'A24', 'Amount in words')
    ws.merge_cells('A25:F26'); cell(ws, 'A25', words_formula('J29'), wrap=True, bold=True)
    ws.merge_cells('A28:F28'); cell(ws, 'A28', '="Bank: "&Business!B9&"   A/c: "&Business!B10&"   IFSC: "&Business!B11', size=9)
    ws.merge_cells('A30:F32')
    cell(ws, 'A30', 'Declaration: We declare that this invoice shows the actual price of the goods or services described and that all particulars are true and correct.', wrap=True, size=9, color=MUTED)
    ws.merge_cells('G31:J31'); cell(ws, 'G31', '="For "&Business!B3', bold=True, align='right')
    ws.merge_cells('G34:J34'); cell(ws, 'G34', 'Authorised signatory', align='right', size=9)
    ws.print_area = 'A1:J35'
    protect(ws)
    howto(wb)
    wb.move_sheet('How to use', -(len(wb.sheetnames) - 1))
    save(wb, 'GST-Invoice-Template.xlsx')

def gst_summary():
    wb = new_wb('gst-monthly-summary-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [22, 30])
    sal = sheet(wb, 'Sales', [12, 12, 24, 18, 18, 14, 7, 8, 12, 12, 12, 14, 10, 7], tab=BLUE)
    pur = sheet(wb, 'Purchases', [12, 12, 24, 18, 18, 14, 7, 8, 12, 12, 12, 14, 9, 10])
    sm = sheet(wb, 'Summary', [26, 16, 14, 14, 14, 16], tab=BLUE)
    lists_sheet(wb, states=True, extra={5: ['Yes', 'No']})
    title(st, 'Settings'); label(st, 'A3', 'Your state'); cell(st, 'B3', 'Punjab', 'in'); dv(st, 'B3', '=Lists!$C$1:$C$37')
    def ledger(ws, who, itc):
        title(ws, who[0], 'Grey columns calculate themselves. Delete the sample rows and enter your own.')
        cols = ['Date', 'Number', who[1], 'GSTIN', 'State', 'Taxable value', 'GST %', 'Type', 'IGST', 'CGST', 'SGST', 'Total']
        cols += (['ITC eligible', 'Month'] if itc else ['Month', 'B2B/B2C'])
        head(ws, 3, cols)
        f = {'H': '=IF($F{r}="","",IF($E{r}=Settings!$B$3,"Intra","Inter"))',
             'I': '=IF($H{r}="","",IF($H{r}="Inter",ROUND($F{r}*N($G{r})/100,2),0))',
             'J': '=IF($H{r}="","",IF($H{r}="Intra",ROUND($F{r}*N($G{r})/200,2),0))',
             'K': '=IF($H{r}="","",IF($H{r}="Intra",ROUND($F{r}*N($G{r})/200,2),0))',
             'L': '=IF($F{r}="","",$F{r}+$I{r}+$J{r}+$K{r})'}
        mcol = 'N' if itc else 'M'
        f[mcol] = '=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'
        if not itc: f['N'] = '=IF($F{r}="","",IF($D{r}<>"","B2B","B2C"))'
        kinds = {c: 'calc' for c in 'HIJKL' + mcol + ('' if itc else 'N')}
        grid(ws, 4, 503, 'ABCDEFGHIJKLMN', kinds, {'A': DATE, 'F': INR, 'I': INR, 'J': INR, 'K': INR, 'L': INR, mcol: 'mmm-yyyy'}, f)
        dv(ws, 'E4:E503', '=Lists!$C$1:$C$37'); dv(ws, 'G4:G503', '"0,5,12,18,28,40"')
        if itc: dv(ws, 'M4:M503', '=Lists!$E$1:$E$2')
        ws.freeze_panes = 'A4'
        return ws
    ledger(sal, ('Sales invoices', 'Customer'), False)
    ledger(pur, ('Purchase bills', 'Supplier'), True)
    d = MONTH1
    for i, row in enumerate([(d.replace(day=2), 'INV-101', 'Gupta Enterprises', '03AAACG1111A1Z2', 'Punjab', 42000, 18),
                             (d.replace(day=5), 'INV-102', 'Delhi Office Supplies', '07AAACD2222B1Z3', 'Delhi', 18500, 18),
                             (d.replace(day=9), 'INV-103', 'Walk-in customer', '', 'Punjab', 3200, 5),
                             (d.replace(day=14), 'INV-104', 'Gupta Enterprises', '03AAACG1111A1Z2', 'Punjab', 9600, 5)], 4):
        for j, v in enumerate(row): sal.cell(i, j + 1, v)
    for i, row in enumerate([(d.replace(day=3), 'P-7781', 'Paper Mills Ltd', '03AAAPM3333C1Z4', 'Punjab', 30000, 18, 'Yes'),
                             (d.replace(day=11), 'HR-209', 'Haryana Stationers', '06AAAHS4444D1Z5', 'Haryana', 7000, 5, 'Yes'),
                             (d.replace(day=20), 'R-55', 'Office rent', '', 'Punjab', 12000, 0, 'No')], 4):
        for j, v in enumerate(row[:7]): pur.cell(i, j + 1, v)
        pur.cell(i, 13, row[7])
    ws = sm
    title(ws, 'GST summary for the month', 'Type the first day of the month in the yellow cell.')
    label(ws, 'A3', 'Month'); cell(ws, 'B3', MONTH1, 'in', 'mmmm yyyy', align='left')
    cell(ws, 'C3', '=DATE(YEAR(B3),MONTH(B3),1)', fmt=DATE, color='FFFFFF')
    label(ws, 'A5', 'Outward supplies (sales) by GST rate')
    head(ws, 6, ['GST rate', 'Taxable value', 'IGST', 'CGST', 'SGST', 'Total tax'])
    S = lambda col, extra: 'SUMIFS(Sales!$%s:$%s,Sales!$M:$M,$C$3%s)' % (col, col, extra)
    for i, rate in enumerate([0, 5, 12, 18, 28, 40], 7):
        cell(ws, 'A%d' % i, '%d%%' % rate, 'box', align='center')
        for col, src in zip('BCDE', 'FIJK'):
            cell(ws, '%s%d' % (col, i), '=' + S(src, ',Sales!$G:$G,%d' % rate), 'calc', INR)
        cell(ws, 'F%d' % i, '=SUM(C%d:E%d)' % (i, i), 'calc', INR)
    cell(ws, 'A13', 'Total', 'soft', bold=True)
    for col in 'BCDEF': cell(ws, '%s13' % col, '=SUM(%s7:%s12)' % (col, col), 'soft', INR, bold=True)
    label(ws, 'A15', 'Sales by type of customer')
    head(ws, 16, ['Type', 'Taxable value', 'IGST', 'CGST', 'SGST', 'Total tax'])
    for i, k in enumerate(['B2B', 'B2C'], 17):
        cell(ws, 'A%d' % i, k + (' (buyer has GSTIN)' if k == 'B2B' else ' (no GSTIN)'), 'box')
        for col, src in zip('BCDE', 'FIJK'): cell(ws, '%s%d' % (col, i), '=' + S(src, ',Sales!$N:$N,"%s"' % k), 'calc', INR)
        cell(ws, 'F%d' % i, '=SUM(C%d:E%d)' % (i, i), 'calc', INR)
    label(ws, 'A20', 'Input tax credit (eligible purchases)')
    head(ws, 21, ['', 'Taxable value', 'IGST', 'CGST', 'SGST', 'Total ITC'])
    P = lambda col: '=SUMIFS(Purchases!$%s:$%s,Purchases!$N:$N,$C$3,Purchases!$M:$M,"Yes")' % (col, col)
    cell(ws, 'A22', 'Eligible ITC', 'box')
    for col, src in zip('BCDE', 'FIJK'): cell(ws, '%s22' % col, P(src), 'calc', INR)
    cell(ws, 'F22', '=SUM(C22:E22)', 'calc', INR)
    label(ws, 'A24', 'Tax position (estimate)')
    head(ws, 25, ['', '', 'IGST', 'CGST', 'SGST', 'Total'])
    cell(ws, 'A26', 'Output tax', 'box'); cell(ws, 'A27', 'Less: eligible ITC', 'box'); cell(ws, 'A28', 'Net (output − ITC)', 'soft', bold=True)
    for col in 'CDEF':
        cell(ws, '%s26' % col, '=%s13' % col, 'calc', INR); cell(ws, '%s27' % col, '=%s22' % col, 'calc', INR)
        cell(ws, '%s28' % col, '=%s26-%s27' % (col, col), 'soft', INR, bold=True)
    ws.merge_cells('A30:F32')
    cell(ws, 'A30', 'A negative net means more credit than tax this month. The GST portal sets off credit in a fixed order (IGST first), '
         'so the cash payable under each head can differ from this estimate: use the portal\'s figures for payment.', wrap=True, size=9, color=MUTED)
    protect(ws)
    save(wb, 'GST-Monthly-Summary.xlsx')

def client_dues():
    wb = new_wb('client-dues-ageing-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [24, 30])
    cl = sheet(wb, 'Clients', [30, 16])
    bl = sheet(wb, 'Bills', [26, 12, 12, 9, 14, 14, 14, 12, 10, 10], tab=BLUE)
    ag = sheet(wb, 'Ageing', [26, 14, 13, 13, 13, 13, 13, 70], tab=BLUE)
    title(st, 'Settings')
    label(st, 'A3', 'Position as on'); cell(st, 'B3', '=TODAY()', 'in', DATE, align='left')
    label(st, 'A4', 'Your business name'); cell(st, 'B4', 'Sharma Traders', 'in')
    cell(st, 'A6', 'Leave "As on" as =TODAY() for today\'s position, or type a date.', size=9, color=MUTED)
    title(cl, 'Clients'); head(cl, 3, ['Client name', 'Phone']); grid(cl, 4, 203, 'AB', {})
    for i, row in enumerate([('Gupta Enterprises', '98760 11111'), ('Delhi Office Supplies', '98110 22222'), ('Mehta & Sons', '98150 33333')], 4):
        for j, v in enumerate(row): cl.cell(i, j + 1, v)
    title(bl, 'Bills', 'One row per bill. Grey columns calculate themselves.')
    head(bl, 3, ['Client', 'Bill no.', 'Bill date', 'Credit days', 'Amount', 'Received', 'Balance', 'Due date', 'Days overdue', 'Age'])
    f = {'G': '=IF($E{r}="","",$E{r}-N($F{r}))', 'H': '=IF($C{r}="","",$C{r}+N($D{r}))',
         'I': '=IF(OR($G{r}="",N($G{r})<=0,$H{r}=""),"",MAX(0,Settings!$B$3-$H{r}))',
         'J': '=IF($I{r}="","",IF(Settings!$B$3<=$H{r},"Not due",IF($I{r}<=30,"1-30",IF($I{r}<=60,"31-60",IF($I{r}<=90,"61-90","90+")))))'}
    grid(bl, 4, 1003, 'ABCDEFGHIJ', {c: 'calc' for c in 'GHIJ'}, {'C': DATE, 'E': INR, 'F': INR, 'G': INR, 'H': DATE}, f)
    dv(bl, 'A4:A1003', '=Clients!$A$4:$A$203')
    t = TODAY
    for i, row in enumerate([('Gupta Enterprises', 'INV-071', t - datetime.timedelta(days=95), 30, 56000, 20000),
                             ('Gupta Enterprises', 'INV-088', t - datetime.timedelta(days=40), 30, 23500, 0),
                             ('Delhi Office Supplies', 'INV-090', t - datetime.timedelta(days=50), 15, 41300, 41300),
                             ('Delhi Office Supplies', 'INV-097', t - datetime.timedelta(days=10), 30, 18200, 0),
                             ('Mehta & Sons', 'INV-079', t - datetime.timedelta(days=70), 0, 9800, 5000)], 4):
        for j, v in enumerate(row): bl.cell(i, j + 1, v)
    bl.freeze_panes = 'A4'
    title(ag, 'Ageing of dues by client', '="As on "&TEXT(Settings!B3,"dd-mm-yyyy")')
    ag['A2'].value = '="As on "&TEXT(Settings!B3,"dd-mm-yyyy")'
    head(ag, 3, ['Client', 'Total due', 'Not due', '1-30', '31-60', '61-90', '90+', 'Reminder message (copy into WhatsApp or email)'])
    for i, r in enumerate(range(4, 204)):
        src = i + 4
        cell(ag, 'A%d' % r, '=IF(Clients!A%d="","",Clients!A%d)' % (src, src), 'box')
        cell(ag, 'B%d' % r, '=IF($A{r}="","",SUMIFS(Bills!$G:$G,Bills!$A:$A,$A{r}))'.format(r=r), 'calc', INR0)
        for col in 'CDEFG':
            cell(ag, '%s%d' % (col, r), '=IF($A{r}="","",SUMIFS(Bills!$G:$G,Bills!$A:$A,$A{r},Bills!$J:$J,{c}$3))'.format(r=r, c=col), 'calc', INR0)
        cell(ag, 'H%d' % r, ('=IF(OR($A{r}="",N($B{r})<=0),"","Dear "&$A{r}&", this is a gentle reminder that ₹ "&TEXT($B{r},"%s")&'
                             '" is pending against your bills with "&Settings!$B$4&". Kindly arrange the payment at the earliest. Thank you.")' % INR_TXT).format(r=r), wrap=False, size=9)
    ag.conditional_formatting.add('G4:G203', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='FECACA')))
    ag.freeze_panes = 'B4'
    protect(ag)
    save(wb, 'Client-Dues-Ageing.xlsx')

def bank_reco():
    wb = new_wb('bank-reconciliation-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [36, 26, 4, 56])
    it = sheet(wb, 'Items', [52, 12, 40, 14], tab=BLUE)
    sm = sheet(wb, 'Statement', [56, 18, 18], tab=BLUE)
    ADD = ['Add: cheques issued but not yet presented', 'Add: interest or direct deposits credited by bank, not in cash book', 'Add: error correction']
    LESS = ['Less: cheques deposited but not yet credited', 'Less: bank charges or direct debits, not in cash book', 'Less: error correction']
    title(st, 'Settings')
    for i, (k, v, fmt) in enumerate([('Bank name', 'State Bank of India', None), ('Account number', 'XXXXXX1234', None),
                                     ('Month ending', MONTH1 + datetime.timedelta(days=-1), DATE),
                                     ('Balance as per cash book (bank column)', 154230.50, INR), ('Balance as per bank statement', 163985.50, INR)], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    label(st, 'D2', 'Item types (used on the Items sheet)')
    for i, k in enumerate(ADD + LESS, 3): st.cell(i, 4, k)
    title(it, 'Unmatched items', 'Every entry that is in one record but not the other.')
    head(it, 3, ['Type', 'Date', 'Details', 'Amount'])
    grid(it, 4, 203, 'ABCD', {}, {'B': DATE, 'D': INR})
    dv(it, 'A4:A203', '=Settings!$D$3:$D$8')
    d = MONTH1 - datetime.timedelta(days=3)
    for i, row in enumerate([(ADD[0], d, 'Cheque 004512 to Paper Mills Ltd', 18500), (ADD[1], d, 'Savings interest', 455),
                             (LESS[0], d, 'Cheque from Gupta Enterprises deposited', 8600), (LESS[1], d, 'SMS and service charges', 600)], 4):
        for j, v in enumerate(row): it.cell(i, j + 1, v)
    ws = sm
    title(ws, 'Bank reconciliation statement')
    ws['A2'] = '=Settings!B3&"  ·  A/c "&Settings!B4&"  ·  as on "&TEXT(Settings!B5,"dd-mm-yyyy")'
    ws['A2'].font = Font(size=10, color=MUTED)
    head(ws, 4, ['Particulars', '₹', '₹'])
    cell(ws, 'A5', 'Balance as per cash book', 'box', bold=True); cell(ws, 'C5', '=Settings!B6', 'calc', INR)
    r = 6
    for k in ADD:
        cell(ws, 'A%d' % r, k, 'box'); cell(ws, 'B%d' % r, '=SUMIF(Items!$A:$A,A%d,Items!$D:$D)' % r, 'calc', INR); r += 1
    cell(ws, 'A%d' % r, 'Total additions', 'box', bold=True); cell(ws, 'C%d' % r, '=SUM(B6:B%d)' % (r - 1), 'calc', INR); add_r = r; r += 1
    l0 = r
    for k in LESS:
        cell(ws, 'A%d' % r, k, 'box'); cell(ws, 'B%d' % r, '=SUMIF(Items!$A:$A,A%d,Items!$D:$D)' % r, 'calc', INR); r += 1
    cell(ws, 'A%d' % r, 'Total deductions', 'box', bold=True); cell(ws, 'C%d' % r, '=SUM(B%d:B%d)' % (l0, r - 1), 'calc', INR); less_r = r; r += 1
    cell(ws, 'A%d' % r, 'Adjusted balance (should equal the bank statement)', 'soft', bold=True)
    cell(ws, 'C%d' % r, '=C5+C%d-C%d' % (add_r, less_r), 'soft', INR, bold=True); adj = r; r += 1
    cell(ws, 'A%d' % r, 'Balance as per bank statement', 'box', bold=True); cell(ws, 'C%d' % r, '=Settings!B7', 'calc', INR); bank = r; r += 1
    cell(ws, 'A%d' % r, 'Difference (must be 0.00)', 'box', bold=True); cell(ws, 'C%d' % r, '=ROUND(C%d-C%d,2)' % (adj, bank), 'box', INR, bold=True)
    ws.conditional_formatting.add('C%d' % r, FormulaRule(formula=['ABS(C%d)<0.005' % r], fill=PatternFill('solid', fgColor='BBF7D0')))
    ws.conditional_formatting.add('C%d' % r, FormulaRule(formula=['ABS(C%d)>=0.005' % r], fill=PatternFill('solid', fgColor='FECACA')))
    r += 3
    cell(ws, 'A%d' % r, 'Prepared by: ____________________'); cell(ws, 'B%d' % r, 'Checked by: ____________________'); ws.merge_cells('B%d:C%d' % (r, r))
    protect(ws)
    save(wb, 'Bank-Reconciliation-Statement.xlsx')

# ------------------------------------------------------------------ Shop
def stock_register():
    wb = new_wb('stock-register-excel')
    howto(wb, first=True)
    it = sheet(wb, 'Items', [30, 9, 11, 11, 13, 13])
    mv = sheet(wb, 'Movements', [12, 30, 10, 10, 30], tab=BLUE)
    sk = sheet(wb, 'Stock', [30, 9, 11, 10, 10, 11, 10, 15, 15, 10], tab=BLUE)
    title(it, 'Items', 'Opening stock = what you have on the day you start.')
    head(it, 3, ['Item', 'Unit', 'Opening stock', 'Reorder level', 'Purchase rate', 'Sale rate'])
    grid(it, 4, 303, 'ABCDEF', {}, {'E': INR, 'F': INR})
    for i, row in enumerate([('Basmati rice 5 kg', 'Bag', 40, 10, 520, 610), ('Sunflower oil 1 L', 'Pouch', 60, 24, 135, 155),
                             ('Toor dal 1 kg', 'Pack', 35, 15, 140, 165), ('Sugar 1 kg', 'Pack', 50, 20, 42, 48), ('Tea 250 g', 'Pack', 18, 12, 115, 140)], 4):
        for j, v in enumerate(row): it.cell(i, j + 1, v)
    title(mv, 'Stock in and out', 'One row per movement: purchase (In), sale (Out), return or adjustment.')
    head(mv, 3, ['Date', 'Item', 'In', 'Out', 'Note'])
    grid(mv, 4, 1003, 'ABCDE', {}, {'A': DATE})
    dv(mv, 'B4:B1003', '=Items!$A$4:$A$303')
    d = MONTH1
    for i, row in enumerate([(d, 'Sunflower oil 1 L', None, 18, 'Sales'), (d, 'Sugar 1 kg', None, 22, 'Sales'),
                             (d + datetime.timedelta(days=1), 'Sugar 1 kg', 40, None, 'Purchase bill 552'),
                             (d + datetime.timedelta(days=1), 'Tea 250 g', None, 9, 'Sales'),
                             (d + datetime.timedelta(days=2), 'Sunflower oil 1 L', None, 21, 'Sales')], 4):
        for j, v in enumerate(row): mv.cell(i, j + 1, v)
    mv.freeze_panes = 'A4'
    ws = sk
    title(ws, 'Stock position')
    label(ws, 'A2', 'Total stock value'); cell(ws, 'H2', '=SUM(H5:H304)', 'soft', INR, bold=True); cell(ws, 'I2', '=SUM(I5:I304)', 'soft', INR, bold=True)
    head(ws, 4, ['Item', 'Unit', 'Opening', 'In', 'Out', 'Closing', 'Reorder level', 'Value (purchase)', 'Value (sale)', 'Status'])
    for i in range(300):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Items!B%d&"")' % (r, s), 'box')
        cell(ws, 'C%d' % r, '=IF($A%d="","",N(Items!C%d))' % (r, s), 'calc')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIF(Movements!$B:$B,$A{r},Movements!$C:$C))'.format(r=r), 'calc')
        cell(ws, 'E%d' % r, '=IF($A{r}="","",SUMIF(Movements!$B:$B,$A{r},Movements!$D:$D))'.format(r=r), 'calc')
        cell(ws, 'F%d' % r, '=IF($A{r}="","",C{r}+D{r}-E{r})'.format(r=r), 'calc')
        cell(ws, 'G%d' % r, '=IF($A%d="","",N(Items!D%d))' % (r, s), 'calc')
        cell(ws, 'H%d' % r, '=IF($A{r}="","",F{r}*N(Items!E{s}))'.format(r=r, s=s), 'calc', INR)
        cell(ws, 'I%d' % r, '=IF($A{r}="","",F{r}*N(Items!F{s}))'.format(r=r, s=s), 'calc', INR)
        cell(ws, 'J%d' % r, '=IF($A{r}="","",IF(F{r}<=G{r},"Reorder","OK"))'.format(r=r), 'calc')
    ws.conditional_formatting.add('J5:J304', CellIsRule(operator='equal', formula=['"Reorder"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('F5:F304', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'B5'
    protect(ws)
    save(wb, 'Stock-Register.xlsx')

def cash_book():
    wb = new_wb('daily-cash-book-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [26, 18, 4, 26])
    en = sheet(wb, 'Entries', [12, 36, 18, 14, 14, 16], tab=BLUE)
    dw = sheet(wb, 'Day-wise', [14, 16, 16, 16, 16], tab=BLUE)
    hd = sheet(wb, 'Heads', [26, 16, 16])
    title(st, 'Settings')
    label(st, 'A3', 'Month (first day)'); cell(st, 'B3', MONTH1, 'in', DATE, align='left')
    label(st, 'A4', 'Opening cash balance'); cell(st, 'B4', 8500, 'in', INR, align='left')
    label(st, 'D2', 'Heads (edit freely)')
    heads = ['Sales', 'Purchase', 'Rent', 'Salary', 'Electricity', 'Transport', 'Owner drawings', 'Owner capital', 'Bank deposit', 'Bank withdrawal', 'Other income', 'Other expense']
    for i in range(3, 23): cell(st, 'D%d' % i, heads[i - 3] if i - 3 < len(heads) else None, 'in')
    title(en, 'Cash entries', 'Balance = cash that should be in hand after each entry.')
    head(en, 3, ['Date', 'Particulars', 'Head', 'Receipt', 'Payment', 'Balance'])
    grid(en, 4, 1003, 'ABCDEF', {'F': 'calc'}, {'A': DATE, 'D': INR, 'E': INR, 'F': INR},
         {'F': '=IF(AND($D{r}="",$E{r}=""),"",Settings!$B$4+SUM($D$4:$D{r})-SUM($E$4:$E{r}))'})
    dv(en, 'C4:C1003', '=Settings!$D$3:$D$22')
    d = MONTH1
    for i, row in enumerate([(d, 'Counter sales', 'Sales', 6420, None), (d, 'Milk supplier', 'Purchase', None, 2300),
                             (d, 'Delivery boy fuel', 'Transport', None, 150), (d + datetime.timedelta(days=1), 'Counter sales', 'Sales', 7180, None),
                             (d + datetime.timedelta(days=1), 'Deposited in bank', 'Bank deposit', None, 10000)], 4):
        for j, v in enumerate(row): en.cell(i, j + 1, v)
    en.freeze_panes = 'A4'
    title(dw, 'Day-wise summary'); head(dw, 3, ['Date', 'Opening', 'Receipts', 'Payments', 'Closing'])
    for n in range(31):
        r = 4 + n
        cell(dw, 'A%d' % r, '=IF(MONTH(Settings!$B$3+%d)=MONTH(Settings!$B$3),Settings!$B$3+%d,"")' % (n, n), 'box', DATE)
        cell(dw, 'B%d' % r, ('=Settings!$B$4' if n == 0 else '=IF($A{r}="","",E{p})'.format(r=r, p=r - 1)), 'calc', INR)
        cell(dw, 'C%d' % r, '=IF($A{r}="","",SUMIF(Entries!$A:$A,$A{r},Entries!$D:$D))'.format(r=r), 'calc', INR)
        cell(dw, 'D%d' % r, '=IF($A{r}="","",SUMIF(Entries!$A:$A,$A{r},Entries!$E:$E))'.format(r=r), 'calc', INR)
        cell(dw, 'E%d' % r, '=IF($A{r}="","",B{r}+C{r}-D{r})'.format(r=r), 'calc', INR)
    cell(dw, 'A36', 'Month total', 'soft', bold=True)
    for col in 'CD': cell(dw, '%s36' % col, '=SUM(%s4:%s34)' % (col, col), 'soft', INR, bold=True)
    protect(dw)
    title(hd, 'Totals by head'); head(hd, 3, ['Head', 'Receipts', 'Payments'])
    for i in range(20):
        r, s = 4 + i, 3 + i
        cell(hd, 'A%d' % r, '=IF(Settings!D%d="","",Settings!D%d)' % (s, s), 'box')
        cell(hd, 'B%d' % r, '=IF($A{r}="","",SUMIF(Entries!$C:$C,$A{r},Entries!$D:$D))'.format(r=r), 'calc', INR)
        cell(hd, 'C%d' % r, '=IF($A{r}="","",SUMIF(Entries!$C:$C,$A{r},Entries!$E:$E))'.format(r=r), 'calc', INR)
    protect(hd)
    save(wb, 'Daily-Cash-Book.xlsx')

# ------------------------------------------------------------------ Teachers
def attendance():
    wb = new_wb('attendance-register-excel')
    ws = sheet(wb, 'Register', [7, 24] + [4.2] * 31 + [6, 6, 6, 8], first=True, tab=BLUE)
    ws.page_setup.orientation = 'landscape'
    title(ws, 'Student attendance register')
    for i, (k, v, fmt) in enumerate([('School / centre', 'Govt. Senior Secondary School', None), ('Class & section', 'VIII-B', None),
                                     ('Month (first day)', MONTH1, 'mmmm yyyy'), ('Warn below', 0.75, '0%')], 2):
        ws.merge_cells('C%d:H%d' % (i, i)); label(ws, 'B%d' % i, k); cell(ws, 'C%d' % i, v, 'in', fmt, align='left')
    for n in range(31):
        col = L(3 + n)
        cell(ws, '%s7' % col, '=IF(%d<=DAY(EOMONTH($C$4,0)),$C$4+%d,"")' % (n + 1, n), 'calc', 'd', align='center', bold=True)
        cell(ws, '%s8' % col, '=IF(%s7="","",LEFT(TEXT(%s7,"ddd"),2))' % (col, col), 'calc', align='center', size=8)
    head(ws, 7, ['Roll', 'Name'])
    ws.merge_cells('A7:A8'); ws.merge_cells('B7:B8')
    for col, k in zip(['AH', 'AI', 'AJ', 'AK'], ['P', 'A', 'L', '%']):
        c = ws['%s7' % col]; c.value = k; c.font, c.fill, c.alignment = Font(bold=True, color='FFFFFF'), F_HEAD, Alignment(horizontal='center')
        ws.merge_cells('%s7:%s8' % (col, col))
    names = ['Aarav Sharma', 'Ananya Singh', 'Ishaan Verma', 'Diya Kaur', 'Kabir Gupta', 'Meera Joshi', 'Rohan Das', 'Saanvi Patel']
    for i in range(60):
        r = 9 + i
        cell(ws, 'A%d' % r, i + 1 if i < len(names) else None, 'in', align='center')
        cell(ws, 'B%d' % r, names[i] if i < len(names) else None, 'in')
        for n in range(31): cell(ws, '%s%d' % (L(3 + n), r), None, 'in', align='center')
        cell(ws, 'AH%d' % r, '=IF($B{r}="","",COUNTIF(C{r}:AG{r},"P"))'.format(r=r), 'calc', align='center')
        cell(ws, 'AI%d' % r, '=IF($B{r}="","",COUNTIF(C{r}:AG{r},"A"))'.format(r=r), 'calc', align='center')
        cell(ws, 'AJ%d' % r, '=IF($B{r}="","",COUNTIF(C{r}:AG{r},"L"))'.format(r=r), 'calc', align='center')
        cell(ws, 'AK%d' % r, '=IF(OR($B{r}="",N(AH{r})+N(AI{r})+N(AJ{r})=0),"",AH{r}/(AH{r}+AI{r}+AJ{r}))'.format(r=r), 'calc', '0%', align='center')
    # sample marks for the first week's working days
    import random
    rnd = random.Random(7)
    for i in range(len(names)):
        for n in range(10):
            dte = MONTH1 + datetime.timedelta(days=n)
            if dte.month == MONTH1.month and dte.weekday() != 6:
                ws.cell(9 + i, 3 + n, 'A' if rnd.random() < 0.08 else ('L' if rnd.random() < 0.04 else 'P'))
    cell(ws, 'B70', 'Present each day', 'soft', bold=True)
    for n in range(31):
        col = L(3 + n); cell(ws, '%s70' % col, '=IF(%s$7="","",COUNTIF(%s9:%s68,"P"))' % (col, col, col), 'soft', align='center', size=8)
    dv(ws, 'C9:AG68', '"P,A,L"')
    grey = PatternFill('solid', fgColor='E2E8F0')
    ws.conditional_formatting.add('C7:AG70', FormulaRule(formula=['C$7=""'], fill=PatternFill('solid', fgColor='CBD5E1')))
    ws.conditional_formatting.add('C7:AG70', FormulaRule(formula=['AND(C$7<>"",WEEKDAY(C$7)=1)'], fill=grey))
    ws.conditional_formatting.add('C9:AG68', CellIsRule(operator='equal', formula=['"A"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('AK9:AK68', FormulaRule(formula=['AND(AK9<>"",AK9<$C$5)'], fill=PatternFill('solid', fgColor='FECACA')))
    ws.freeze_panes = 'C9'
    ws.print_title_rows = '7:8'
    protect(ws)
    howto(wb)
    save(wb, 'Student-Attendance-Register.xlsx')

def marks():
    wb = new_wb('marks-report-card-excel')
    howto(wb, first=True)
    su = sheet(wb, 'Setup', [26, 14, 14])
    mk = sheet(wb, 'Marks', [7, 22, 10, 10, 10, 10, 10, 10, 9, 9, 9, 8, 7, 9], tab=BLUE)
    rc = sheet(wb, 'Report Card', [26, 14, 14, 14, 14], tab=BLUE)
    title(su, 'Setup')
    for i, (k, v) in enumerate([('School', 'Govt. Senior Secondary School'), ('Class & section', 'VIII-B'), ('Exam', 'Half-yearly examination')], 2):
        label(su, 'A%d' % i, k); cell(su, 'B%d' % i, v, 'in'); su.merge_cells('B%d:C%d' % (i, i))
    head(su, 6, ['Subject (leave empty if unused)', 'Maximum marks', 'Pass marks'])
    subs = [('English', 100, 33), ('Hindi', 100, 33), ('Mathematics', 100, 33), ('Science', 100, 33), ('Social Science', 100, 33), ('Punjabi', 100, 33)]
    for i, (s, mx, ps) in enumerate(subs, 7):
        cell(su, 'A%d' % i, s, 'in'); cell(su, 'B%d' % i, mx, 'in'); cell(su, 'C%d' % i, ps, 'in')
    label(su, 'A13', 'Maximum total'); cell(su, 'B13', '=SUMPRODUCT((A7:A12<>"")*B7:B12)', 'calc')
    head(su, 15, ['From percentage', 'Grade'])
    for i, (p, g) in enumerate([(0, 'E'), (33, 'D'), (41, 'C2'), (51, 'C1'), (61, 'B2'), (71, 'B1'), (81, 'A2'), (91, 'A1')], 16):
        cell(su, 'A%d' % i, p, 'in'); cell(su, 'B%d' % i, g, 'in')
    title(mk, 'Marks and result')
    mk['A2'] = '=Setup!B4&"  ·  "&Setup!B3&"  ·  "&Setup!B2'; mk['A2'].font = Font(size=10, color=MUTED)
    head(mk, 3, ['Roll', 'Name', '', '', '', '', '', '', 'Total', 'Max', '%', 'Grade', 'Rank', 'Result'])
    for j in range(6): mk.cell(3, 3 + j).value = '=IF(Setup!A%d="","—",Setup!A%d)' % (7 + j, 7 + j)
    import random
    rnd = random.Random(3)
    names = ['Aarav Sharma', 'Ananya Singh', 'Ishaan Verma', 'Diya Kaur', 'Kabir Gupta', 'Meera Joshi', 'Rohan Das', 'Saanvi Patel']
    for i in range(60):
        r = 4 + i
        cell(mk, 'A%d' % r, i + 1 if i < len(names) else None, 'in', align='center'); cell(mk, 'B%d' % r, names[i] if i < len(names) else None, 'in')
        for j in range(6):
            cell(mk, '%s%d' % (L(3 + j), r), (rnd.randint(42, 98) if i < len(names) else None), 'in', align='center')
        cell(mk, 'I%d' % r, '=IF($B{r}="","",SUM(C{r}:H{r}))'.format(r=r), 'calc')
        cell(mk, 'J%d' % r, '=IF($B{r}="","",Setup!$B$13)'.format(r=r), 'calc')
        cell(mk, 'K%d' % r, '=IF(OR($B{r}="",N(J{r})=0),"",I{r}/J{r})'.format(r=r), 'calc', '0.0%')
        cell(mk, 'L%d' % r, '=IF(K{r}="","",INDEX(Setup!$B$16:$B$23,MATCH(K{r}*100,Setup!$A$16:$A$23,1)))'.format(r=r), 'calc', align='center')
        cell(mk, 'M%d' % r, '=IF(I{r}="","",RANK(I{r},$I$4:$I$63))'.format(r=r), 'calc', align='center')
        fails = ','.join('AND(Setup!$A${s}<>"",OR({c}{r}="AB",N({c}{r})<Setup!$C${s}))'.format(s=7 + j, c=L(3 + j), r=r) for j in range(6))
        cell(mk, 'N%d' % r, '=IF($B{r}="","",IF(OR({f}),"Fail","Pass"))'.format(r=r, f=fails), 'calc', align='center')
    mk.cell(9, 5).value = 24   # one sample student below pass marks in a subject
    mk.conditional_formatting.add('N4:N63', CellIsRule(operator='equal', formula=['"Fail"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    mk.freeze_panes = 'C4'
    protect(mk)
    ws = rc
    ws.merge_cells('A1:E1'); cell(ws, 'A1', '=Setup!B2', bold=True, size=16, align='center', color=BLUE)
    ws.merge_cells('A2:E2'); cell(ws, 'A2', '="REPORT CARD  ·  "&Setup!B4', bold=True, align='center')
    label(ws, 'A4', 'Roll number'); cell(ws, 'B4', 1, 'in', align='left'); dv(ws, 'B4', '=Marks!$A$4:$A$63', 'Pick a roll number')
    m = 'MATCH($B$4,Marks!$A$4:$A$63,0)'
    get = lambda col: '=IFERROR(INDEX(Marks!$%s$4:$%s$63,%s),"")' % (col, col, m)
    for r, (k, f) in enumerate([('Name', get('B')), ('Class', '=Setup!B3')], 5):
        label(ws, 'A%d' % r, k); ws.merge_cells('B%d:E%d' % (r, r)); cell(ws, 'B%d' % r, f, 'box', bold=True)
    head(ws, 8, ['Subject', 'Maximum', 'Obtained', 'Percent', 'Grade'])
    for j in range(6):
        r = 9 + j
        cell(ws, 'A%d' % r, '=IF(Setup!A%d="","",Setup!A%d)' % (7 + j, 7 + j), 'box')
        cell(ws, 'B%d' % r, '=IF(A%d="","",Setup!B%d)' % (r, 7 + j), 'box', align='center')
        cell(ws, 'C%d' % r, '=IF(A{r}="","",IFERROR(IF(INDEX(Marks!${c}$4:${c}$63,{m})="","",INDEX(Marks!${c}$4:${c}$63,{m})),""))'.format(r=r, c=L(3 + j), m=m), 'box', align='center')
        cell(ws, 'D%d' % r, '=IF(OR(A{r}="",C{r}=""),"",N(C{r})/B{r})'.format(r=r), 'box', '0%', align='center')
        cell(ws, 'E%d' % r, '=IF(D{r}="","",INDEX(Setup!$B$16:$B$23,MATCH(D{r}*100,Setup!$A$16:$A$23,1)))'.format(r=r), 'box', align='center')
    for r, (k, f, fmt) in enumerate([('Total', get('I'), None), ('Percentage', get('K'), '0.0%'), ('Grade', get('L'), None),
                                     ('Rank in class', get('M'), None), ('Result', get('N'), None)], 16):
        label(ws, 'A%d' % r, k); cell(ws, 'B%d' % r, f, 'soft', fmt, bold=True, align='center')
    label(ws, 'A22', 'Teacher\'s remarks'); ws.merge_cells('A23:E25'); cell(ws, 'A23', 'Works sincerely. Keep practising mathematics every day.', 'in', wrap=True)
    for col, k in zip('ACE', ['Class teacher', 'Principal', 'Parent']): cell(ws, '%s29' % col, k, align='center', size=9); ws['%s28' % col] = '__________'
    ws.print_area = 'A1:E30'
    protect(ws)
    save(wb, 'Marks-and-Report-Card.xlsx')

def fee_register():
    wb = new_wb('tuition-fee-register-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [26, 30])
    sd = sheet(wb, 'Students', [24, 10, 16, 12, 12])
    py = sheet(wb, 'Payments', [12, 24, 12, 12, 10, 10], tab=BLUE)
    gr = sheet(wb, 'Fee Grid', [22, 10] + [9] * 12 + [12, 8, 12, 12, 70, 3], tab=BLUE)
    lists_sheet(wb, extra={5: ['Cash', 'UPI', 'Bank', 'Cheque']})
    title(st, 'Settings')
    sess = datetime.date(TODAY.year if TODAY.month >= 4 else TODAY.year - 1, 4, 1)
    for i, (k, v, fmt) in enumerate([('Centre name', 'Bright Minds Tuition Centre', None), ('Session starts (month)', sess, 'mmmm yyyy'),
                                     ('Dues up to', '=TODAY()', DATE)], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    title(sd, 'Students'); head(sd, 3, ['Student', 'Class', 'Parent phone', 'Monthly fee', 'Joined (month)'])
    grid(sd, 4, 203, 'ABCDE', {}, {'D': INR0, 'E': 'mmm-yyyy'})
    for i, row in enumerate([('Aarav Sharma', 'VIII', '98140 11111', 1500, sess), ('Diya Kaur', 'X', '98140 22222', 2000, sess),
                             ('Kabir Gupta', 'IX', '98140 33333', 1800, sess + datetime.timedelta(days=62))], 4):
        for j, v in enumerate(row): sd.cell(i, j + 1, v)
    title(py, 'Fee payments', 'One row per month paid. "For month" = any date in that month.')
    head(py, 3, ['Date paid', 'Student', 'For month', 'Amount', 'Mode', 'Month key'])
    grid(py, 4, 1503, 'ABCDEF', {'F': 'calc'}, {'A': DATE, 'C': 'mmm-yyyy', 'D': INR0, 'F': 'mmm-yy'},
         {'F': '=IF($C{r}="","",DATE(YEAR($C{r}),MONTH($C{r}),1))'})
    dv(py, 'B4:B1503', '=Students!$A$4:$A$203'); dv(py, 'E4:E1503', '=Lists!$E$1:$E$4')
    rows, mm = [], sess
    for k in range(3):
        rows += [(mm + datetime.timedelta(days=4), 'Aarav Sharma', mm, 1500, 'UPI'), (mm + datetime.timedelta(days=6), 'Diya Kaur', mm, 2000, 'Cash')]
        mm = (mm + datetime.timedelta(days=32)).replace(day=1)
    for i, row in enumerate(rows, 4):
        for j, v in enumerate(row): py.cell(i, j + 1, v)
    py.freeze_panes = 'A4'
    ws = gr
    title(ws, 'Fee grid', 'Paid amount per month. Red = month due and unpaid.')
    head(ws, 3, ['Student', 'Fee'] + [''] * 12 + ['Total paid', 'Months due', 'Fee due', 'Balance', 'Reminder message (copy and send)'])
    for n in range(12): cell(ws, '%s3' % L(3 + n), '=EDATE(Settings!$B$4,%d)' % n, fmt='mmm-yy', bold=True, color='FFFFFF').fill = F_HEAD
    for i in range(200):
        r, s = 4 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Students!A%d="","",Students!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",N(Students!D%d))' % (r, s), 'calc', INR0)
        cell(ws, 'T%d' % r, '=IF($A{r}="","",DATE(YEAR(MAX(N(Students!E{s}),Settings!$B$4)),MONTH(MAX(N(Students!E{s}),Settings!$B$4)),1))'.format(r=r, s=s), fmt='mmm-yy', color='FFFFFF')
        for n in range(12):
            c = L(3 + n)
            cell(ws, '%s%d' % (c, r), '=IF($A{r}="","",SUMIFS(Payments!$D:$D,Payments!$B:$B,$A{r},Payments!$F:$F,{c}$3))'.format(r=r, c=c), 'calc', '#,##0;;"–"', align='center')
        cell(ws, 'O%d' % r, '=IF($A{r}="","",SUM(C{r}:N{r}))'.format(r=r), 'calc', INR0)
        cell(ws, 'P%d' % r, '=IF($A{r}="","",MAX(0,(YEAR(Settings!$B$5)-YEAR($T{r}))*12+MONTH(Settings!$B$5)-MONTH($T{r})+1))'.format(r=r), 'calc', align='center')
        cell(ws, 'Q%d' % r, '=IF($A{r}="","",P{r}*B{r})'.format(r=r), 'calc', INR0)
        cell(ws, 'R%d' % r, '=IF($A{r}="","",Q{r}-O{r})'.format(r=r), 'calc', INR0)
        cell(ws, 'S%d' % r, ('=IF(OR($A{r}="",N(R{r})<=0),"","Dear parent, the tuition fee of ₹ "&TEXT(R{r},"%s")&" for "&$A{r}&" is pending at "&Settings!$B$3&". Kindly pay at your earliest convenience. Thank you.")' % INR_TXT).format(r=r), size=9)
    ws.conditional_formatting.add('C4:N203', FormulaRule(formula=['AND($A4<>"",N(C4)=0,C$3>=$T4,C$3<=Settings!$B$5)'], fill=PatternFill('solid', fgColor='FECACA')))
    ws.column_dimensions['T'].hidden = True
    ws.freeze_panes = 'C4'
    protect(ws)
    save(wb, 'Tuition-Fee-Register.xlsx')

# ------------------------------------------------------------------ Office & HR
def salary():
    wb = new_wb('salary-sheet-slip-excel')
    howto(wb, first=True)
    su = sheet(wb, 'Setup', [26, 40])
    em = sheet(wb, 'Employees', [8, 22, 18, 14, 12, 13, 14, 18, 12, 12, 12])
    sa = sheet(wb, 'Salary', [8, 22, 8, 12, 11, 11, 12, 10, 9, 10, 10, 10, 12, 13, 11, 11, 13], tab=BLUE)
    ps = sheet(wb, 'Payslip', [22, 16, 4, 22, 16], tab=BLUE)
    lists_sheet(wb, extra={5: ['Yes', 'No']})
    title(su, 'Setup')
    for i, (k, v, fmt) in enumerate([('Company name', 'Sharma Traders Pvt. Ltd.', None), ('Address', '12, Main Bazaar, Ludhiana, Punjab 141001', None),
                                     ('Salary month', MONTH1, 'mmmm yyyy'), ('Working days in month', 26, None), ('Deduct PF?', 'Yes', None), ('Deduct ESI?', 'Yes', None)], 3):
        label(su, 'A%d' % i, k); cell(su, 'B%d' % i, v, 'in', fmt, align='left')
    dv(su, 'B7:B8', '=Lists!$E$1:$E$2')
    cell(su, 'A10', 'PF: 12% of earned basic, wage ceiling ₹15,000. ESI: 0.75% employee + 3.25% employer when gross is ₹21,000 or less. Check current rules.', size=9, color=MUTED)
    title(em, 'Employees', 'Monthly salary at full attendance.')
    head(em, 3, ['Code', 'Name', 'Designation', 'Department', 'Joined', 'PAN', 'UAN', 'Bank account', 'Basic', 'HRA', 'Other allowances'])
    grid(em, 4, 103, 'ABCDEFGHIJK', {}, {'E': DATE, 'I': INR0, 'J': INR0, 'K': INR0})
    for i, row in enumerate([('E001', 'Rahul Mehra', 'Accountant', 'Accounts', datetime.date(2022, 6, 1), 'ABCPM1234K', '100200300400', 'SBI XXXX1234', 14000, 5600, 2400),
                             ('E002', 'Pooja Rani', 'Sales Executive', 'Sales', datetime.date(2023, 1, 16), 'BCDPR2345L', '100200300401', 'PNB XXXX5678', 11000, 4400, 1600),
                             ('E003', 'Arjun Nair', 'Manager', 'Operations', datetime.date(2020, 4, 1), 'CDEPN3456M', '100200300402', 'HDFC XXXX9012', 32000, 12800, 6000)], 4):
        for j, v in enumerate(row): em.cell(i, j + 1, v)
    ws = sa
    title(ws, 'Salary sheet'); ws['A2'] = '=Setup!B3&"  ·  "&TEXT(Setup!B5,"mmmm yyyy")'; ws['A2'].font = Font(size=10, color=MUTED)
    cols = ['Code', 'Name', 'Days paid', 'Basic', 'HRA', 'Other', 'Gross', 'PF', 'ESI', 'Prof. tax', 'TDS', 'Advance', 'Deductions', 'Net pay',
            'Employer PF', 'Employer ESI', 'Cost to company']
    head(ws, 5, cols)
    label(ws, 'B4', 'Totals')
    for col in 'DEFGHIJKLMNOPQ': cell(ws, '%s4' % col, '=SUM(%s6:%s105)' % (col, col), 'soft', INR0, bold=True)
    on_pf, on_esi = 'Setup!$B$7="Yes"', 'Setup!$B$8="Yes"'
    for i in range(100):
        r, s = 6 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Employees!A%d="","",Employees!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Employees!B%d)' % (r, s), 'box')
        cell(ws, 'C%d' % r, None, 'in', align='center')
        for col, src in zip('DEF', 'IJK'):
            cell(ws, '%s%d' % (col, r), '=IF($A{r}="","",ROUND(N(Employees!{src}{s})*N($C{r})/Setup!$B$6,0))'.format(r=r, s=s, src=src), 'calc', INR0)
        cell(ws, 'G%d' % r, '=IF($A{r}="","",D{r}+E{r}+F{r})'.format(r=r), 'calc', INR0)
        cell(ws, 'H%d' % r, '=IF($A{r}="","",IF({pf},ROUND(MIN(D{r},15000)*12%,0),0))'.format(r=r, pf=on_pf), 'calc', INR0)
        cell(ws, 'I%d' % r, '=IF($A{r}="","",IF(AND({esi},G{r}>0,G{r}<=21000),ROUNDUP(G{r}*0.75%,0),0))'.format(r=r, esi=on_esi), 'calc', INR0)
        for col in 'JKL': cell(ws, '%s%d' % (col, r), None, 'in', INR0)
        cell(ws, 'M%d' % r, '=IF($A{r}="","",H{r}+I{r}+N(J{r})+N(K{r})+N(L{r}))'.format(r=r), 'calc', INR0)
        cell(ws, 'N%d' % r, '=IF($A{r}="","",G{r}-M{r})'.format(r=r), 'calc', INR0, bold=True)
        cell(ws, 'O%d' % r, '=IF($A{r}="","",IF({pf},ROUND(MIN(D{r},15000)*12%,0),0))'.format(r=r, pf=on_pf), 'calc', INR0)
        cell(ws, 'P%d' % r, '=IF($A{r}="","",IF(AND({esi},G{r}>0,G{r}<=21000),ROUNDUP(G{r}*3.25%,0),0))'.format(r=r, esi=on_esi), 'calc', INR0)
        cell(ws, 'Q%d' % r, '=IF($A{r}="","",G{r}+O{r}+P{r})'.format(r=r), 'calc', INR0)
    for r, (days, pt, tds, adv) in zip((6, 7, 8), [(26, 200, None, None), (24.5, 150, None, 1000), (26, 200, 2500, None)]):
        ws['C%d' % r], ws['J%d' % r], ws['K%d' % r], ws['L%d' % r] = days, pt, tds, adv
    ws.freeze_panes = 'C6'
    protect(ws)
    ws = ps
    lists_ws = 'Lists'
    ws.merge_cells('A1:E1'); cell(ws, 'A1', '=Setup!B3', bold=True, size=16, align='center', color=BLUE)
    ws.merge_cells('A2:E2'); cell(ws, 'A2', '=Setup!B4', align='center', size=9)
    ws.merge_cells('A3:E3'); cell(ws, 'A3', '="Payslip for "&TEXT(Setup!B5,"mmmm yyyy")', bold=True, align='center')
    label(ws, 'A5', 'Employee code'); cell(ws, 'B5', 'E001', 'in', align='left'); dv(ws, 'B5', '=Employees!$A$4:$A$103', 'Pick an employee')
    em_m = 'MATCH($B$5,Employees!$A:$A,0)'
    sa_m = 'MATCH($B$5,Salary!$A:$A,0)'
    E = lambda col: '=IFERROR(INDEX(Employees!$%s:$%s,%s)&"","")' % (col, col, em_m)
    S = lambda col: '=IFERROR(N(INDEX(Salary!$%s:$%s,%s)),0)' % (col, col, sa_m)
    info = [('Name', E('B'), 'Designation', E('C')), ('Department', E('D'), 'PAN', E('F')), ('UAN', E('G'), 'Bank account', E('H')),
            ('Days paid', S('C'), 'Working days', '=Setup!B6')]
    for i, (a, fa, b, fb) in enumerate(info, 6):
        label(ws, 'A%d' % i, a, bold=False); cell(ws, 'B%d' % i, fa, 'box', bold=True, align='left')
        label(ws, 'D%d' % i, b, bold=False); cell(ws, 'E%d' % i, fb, 'box', bold=True, align='left')
    head(ws, 11, ['Earnings', 'Amount', '', 'Deductions', 'Amount'])
    ws['C11'].fill = PatternFill(None)
    earn = [('Basic', S('D')), ('HRA', S('E')), ('Other allowances', S('F'))]
    ded = [('Provident fund', S('H')), ('ESI', S('I')), ('Professional tax', S('J')), ('TDS', S('K')), ('Advance', S('L'))]
    for i in range(5):
        r = 12 + i
        if i < len(earn): cell(ws, 'A%d' % r, earn[i][0], 'box'); cell(ws, 'B%d' % r, earn[i][1], 'box', INR0)
        cell(ws, 'D%d' % r, ded[i][0], 'box'); cell(ws, 'E%d' % r, ded[i][1], 'box', INR0)
    cell(ws, 'A17', 'Gross earnings', 'soft', bold=True); cell(ws, 'B17', '=SUM(B12:B14)', 'soft', INR0, bold=True)
    cell(ws, 'D17', 'Total deductions', 'soft', bold=True); cell(ws, 'E17', '=SUM(E12:E16)', 'soft', INR0, bold=True)
    label(ws, 'A19', 'Net pay'); cell(ws, 'B19', '=B17-E17', 'soft', INR0, bold=True, size=13)
    ws.merge_cells('A20:E21'); cell(ws, 'A20', words_formula('B19'), wrap=True, bold=True)
    ws.merge_cells('A23:E23'); cell(ws, 'A23', 'This is a computer-generated payslip.', size=9, color=MUTED, align='center')
    cell(ws, 'A26', 'Employee signature', size=9); cell(ws, 'E26', 'Authorised signatory', size=9, align='right')
    ws.print_area = 'A1:E27'
    protect(ws)
    save(wb, 'Salary-Sheet-and-Payslip.xlsx')

def leave_tracker():
    wb = new_wb('leave-tracker-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [24, 16])
    ho = sheet(wb, 'Holidays', [14, 30])
    em = sheet(wb, 'Employees', [24, 18])
    lv = sheet(wb, 'Leaves', [22, 14, 12, 12, 9, 30, 8], tab=BLUE)
    bal = sheet(wb, 'Balance', [24] + [8] * 10, tab=BLUE)
    lists_sheet(wb, extra={5: ['Yes']})
    title(st, 'Settings')
    label(st, 'A3', 'Year starts'); cell(st, 'B3', datetime.date(TODAY.year, 1, 1), 'in', DATE, align='left')
    label(st, 'A4', 'Weekend code'); cell(st, 'B4', '0000001', 'in', '@', align='left')
    cell(st, 'A5', 'Seven digits Monday→Sunday, 1 = day off. 0000001 = Sunday off; 0000011 = Saturday and Sunday off.', size=9, color=MUTED)
    head(st, 7, ['Leave type', 'Days per year'])
    for i, (k, q) in enumerate([('Casual', 12), ('Sick', 12), ('Earned', 15), ('Unpaid', 0), (None, None)], 8):
        cell(st, 'A%d' % i, k, 'in'); cell(st, 'B%d' % i, q, 'in')
    title(ho, 'Holidays'); head(ho, 3, ['Date', 'Holiday']); grid(ho, 4, 63, 'AB', {}, {'A': DATE})
    y = TODAY.year
    for i, (d, n) in enumerate([(datetime.date(y, 1, 26), 'Republic Day'), (datetime.date(y, 8, 15), 'Independence Day'), (datetime.date(y, 10, 2), 'Gandhi Jayanti')], 4):
        ho.cell(i, 1, d); ho.cell(i, 2, n)
    title(em, 'Employees'); head(em, 3, ['Name', 'Department']); grid(em, 4, 103, 'AB', {})
    for i, row in enumerate([('Rahul Mehra', 'Accounts'), ('Pooja Rani', 'Sales'), ('Arjun Nair', 'Operations')], 4):
        em.cell(i, 1, row[0]); em.cell(i, 2, row[1])
    title(lv, 'Leaves taken'); head(lv, 3, ['Employee', 'Type', 'From', 'To', 'Half day', 'Reason', 'Days'])
    grid(lv, 4, 1003, 'ABCDEFG', {'G': 'calc'}, {'C': DATE, 'D': DATE},
         {'G': '=IF(OR($C{r}="",$D{r}=""),"",IF($E{r}="Yes",0.5,_xlfn.NETWORKDAYS.INTL($C{r},$D{r},Settings!$B$4,Holidays!$A$4:$A$63)))'})
    dv(lv, 'A4:A1003', '=Employees!$A$4:$A$103'); dv(lv, 'B4:B1003', '=Settings!$A$8:$A$12'); dv(lv, 'E4:E1003', '=Lists!$E$1:$E$1')
    base = datetime.date(y, 3, 2)
    while base.weekday() != 0: base += datetime.timedelta(days=1)
    for i, row in enumerate([('Rahul Mehra', 'Casual', base, base + datetime.timedelta(days=1), None, 'Family function'),
                             ('Pooja Rani', 'Sick', base + datetime.timedelta(days=7), base + datetime.timedelta(days=9), None, 'Fever'),
                             ('Rahul Mehra', 'Casual', base + datetime.timedelta(days=21), base + datetime.timedelta(days=21), 'Yes', 'Bank work')], 4):
        for j, v in enumerate(row): lv.cell(i, j + 1, v)
    lv.freeze_panes = 'A4'
    ws = bal
    title(ws, 'Leave balance')
    ws['A3'] = 'Employee'; ws['A3'].font, ws['A3'].fill = Font(bold=True, color='FFFFFF'), F_HEAD; ws.merge_cells('A3:A4')
    for k in range(5):
        c1, c2 = L(2 + 2 * k), L(3 + 2 * k)
        ws.merge_cells('%s3:%s3' % (c1, c2))
        cell(ws, '%s3' % c1, '=IF(Settings!A%d="","",Settings!A%d)' % (8 + k, 8 + k), bold=True, color='FFFFFF', align='center').fill = F_HEAD
        for c, t in ((c1, 'Taken'), (c2, 'Left')): cell(ws, '%s4' % c, t, 'soft', bold=True, align='center')
    for i in range(100):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Employees!A%d="","",Employees!A%d)' % (s, s), 'box')
        for k in range(5):
            c1, c2 = L(2 + 2 * k), L(3 + 2 * k)
            cell(ws, '%s%d' % (c1, r), '=IF(OR($A{r}="",{c}$3=""),"",SUMIFS(Leaves!$G:$G,Leaves!$A:$A,$A{r},Leaves!$B:$B,{c}$3,Leaves!$C:$C,">="&Settings!$B$3,Leaves!$C:$C,"<"&EDATE(Settings!$B$3,12)))'.format(r=r, c=c1), 'calc', '0.0', align='center')
            cell(ws, '%s%d' % (c2, r), '=IF({c}{r}="","",N(Settings!$B${q})-{c}{r})'.format(r=r, c=c1, q=8 + k), 'calc', '0.0', align='center')
            ws.conditional_formatting.add('%s%d' % (c2, r), CellIsRule(operator='lessThan', formula=['0'], fill=PatternFill('solid', fgColor='FECACA')))
    ws.freeze_panes = 'B5'
    protect(ws)
    save(wb, 'Employee-Leave-Tracker.xlsx')

# ------------------------------------------------------------------ General office
def task_tracker():
    wb = new_wb('task-tracker-excel')
    howto(wb, first=True)
    li = sheet(wb, 'Lists', [22, 16, 12])
    tk = sheet(wb, 'Tasks', [36, 16, 16, 10, 12, 12, 13, 9, 9, 10, 30], tab=BLUE)
    db = sheet(wb, 'Dashboard', [22, 10, 10, 10, 10], tab=BLUE)
    title(li, 'Lists'); head(li, 3, ['Team members', 'Status', 'Priority'])
    for i in range(4, 34): cell(li, 'A%d' % i, None, 'in')
    for i, n in enumerate(['Neha', 'Vikram', 'Sana', 'Amit'], 4): li['A%d' % i] = n
    STATUS = ['Not started', 'In progress', 'Blocked', 'Done']
    for i, s in enumerate(STATUS, 4): li['B%d' % i] = s
    for i, s in enumerate(['High', 'Medium', 'Low'], 4): li['C%d' % i] = s
    title(tk, 'Tasks', 'Grey columns calculate themselves. Use Data › Filter to see one owner or project.')
    head(tk, 3, ['Task', 'Project', 'Owner', 'Priority', 'Start', 'Due', 'Status', '% done', 'Days left', 'Overdue', 'Notes'])
    grid(tk, 4, 503, 'ABCDEFGHIJK', {'I': 'calc', 'J': 'calc'}, {'E': DATE, 'F': DATE, 'H': '0%'},
         {'I': '=IF(OR($F{r}="",$G{r}="Done"),"",$F{r}-TODAY())', 'J': '=IF(AND($F{r}<>"",$G{r}<>"Done",$F{r}<TODAY()),"Overdue","")'})
    dv(tk, 'C4:C503', '=Lists!$A$4:$A$33'); dv(tk, 'G4:G503', '=Lists!$B$4:$B$7'); dv(tk, 'D4:D503', '=Lists!$C$4:$C$6')
    t = TODAY
    D = lambda n: t + datetime.timedelta(days=n)
    for i, row in enumerate([('Prepare quarterly sales report', 'Reports', 'Neha', 'High', D(-10), D(-2), 'In progress', 0.7),
                             ('Renew office insurance', 'Admin', 'Vikram', 'High', D(-5), D(6), 'Not started', 0),
                             ('Update price list on website', 'Website', 'Sana', 'Medium', D(-8), D(-1), 'Done', 1),
                             ('Collect vendor quotations for laptops', 'Purchase', 'Amit', 'Medium', D(-3), D(10), 'In progress', 0.4),
                             ('Train new staff on billing software', 'HR', 'Neha', 'Low', D(2), D(15), 'Not started', 0),
                             ('Fix printer in accounts room', 'Admin', 'Vikram', 'Low', D(-6), D(-3), 'Blocked', 0.2)], 4):
        for j, v in enumerate(row): tk.cell(i, j + 1, v)
    tk.conditional_formatting.add('A4:K503', FormulaRule(formula=['$J4="Overdue"'], fill=PatternFill('solid', fgColor='FEE2E2')))
    tk.conditional_formatting.add('A4:K503', FormulaRule(formula=['$G4="Done"'], font=Font(color='94A3B8')))
    tk.freeze_panes = 'B4'
    tk.auto_filter.ref = 'A3:K503'
    ws = db
    title(ws, 'Dashboard')
    head(ws, 3, ['Status', 'Tasks'])
    for i, s in enumerate(STATUS, 4):
        cell(ws, 'A%d' % i, s, 'box'); cell(ws, 'B%d' % i, '=COUNTIF(Tasks!$G$4:$G$503,A%d)' % i, 'calc', align='center')
    cell(ws, 'A8', 'Overdue', 'box', bold=True); cell(ws, 'B8', '=COUNTIF(Tasks!$J$4:$J$503,"Overdue")', 'calc', bold=True, align='center')
    head(ws, 10, ['Priority', 'Open tasks'])
    for i, s in enumerate(['High', 'Medium', 'Low'], 11):
        cell(ws, 'A%d' % i, s, 'box'); cell(ws, 'B%d' % i, '=COUNTIFS(Tasks!$D$4:$D$503,A%d,Tasks!$G$4:$G$503,"<>Done")' % i, 'calc', align='center')
    head(ws, 15, ['Owner', 'Total', 'Open', 'Done', 'Overdue'])
    for i in range(30):
        r = 16 + i
        cell(ws, 'A%d' % r, '=IF(Lists!A%d="","",Lists!A%d)' % (4 + i, 4 + i), 'box')
        cell(ws, 'B%d' % r, '=IF($A{r}="","",COUNTIF(Tasks!$C$4:$C$503,$A{r}))'.format(r=r), 'calc', align='center')
        cell(ws, 'C%d' % r, '=IF($A{r}="","",COUNTIFS(Tasks!$C$4:$C$503,$A{r},Tasks!$G$4:$G$503,"<>Done"))'.format(r=r), 'calc', align='center')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",COUNTIFS(Tasks!$C$4:$C$503,$A{r},Tasks!$G$4:$G$503,"Done"))'.format(r=r), 'calc', align='center')
        cell(ws, 'E%d' % r, '=IF($A{r}="","",COUNTIFS(Tasks!$C$4:$C$503,$A{r},Tasks!$J$4:$J$503,"Overdue"))'.format(r=r), 'calc', align='center')
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Tasks by status'; ch.legend = None
    ch.add_data(Reference(ws, min_col=2, min_row=3, max_row=7), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=7))
    ch.series[0].graphicalProperties.solidFill = BLUE
    ch.height, ch.width = 7, 14
    ws.add_chart(ch, 'G3')
    protect(ws)
    save(wb, 'Task-Tracker.xlsx')

def expense_claim():
    wb = new_wb('expense-claim-excel')
    ws = sheet(wb, 'Claim', [5, 12, 18, 34, 12, 14, 3, 20, 14], first=True, tab=BLUE)
    CATS = ['Travel', 'Lodging', 'Food', 'Local conveyance', 'Phone & internet', 'Office supplies', 'Other']
    lists_sheet(wb, extra={5: CATS})
    ws.merge_cells('A1:I1'); cell(ws, 'A1', 'EXPENSE REIMBURSEMENT CLAIM', bold=True, size=16, color=BLUE, align='center')
    for i, (k, v, fmt) in enumerate([('Employee name', 'Rahul Mehra', None), ('Employee code', 'E001', None), ('Department', 'Accounts', None),
                                     ('Purpose', 'Client visit to Delhi for GST audit', None), ('Period', '="From "&TEXT(MIN(B10:B34),"dd-mm-yyyy")&" to "&TEXT(MAX(B10:B34),"dd-mm-yyyy")', None)], 3):
        ws.merge_cells('A%d:B%d' % (i, i)); label(ws, 'A%d' % i, k)
        ws.merge_cells('C%d:F%d' % (i, i)); cell(ws, 'C%d' % i, v, 'calc' if k == 'Period' else 'in')
    head(ws, 9, ['#', 'Date', 'Category', 'Description', 'Bill no.', 'Amount'])
    for i in range(25):
        r = 10 + i
        cell(ws, 'A%d' % r, i + 1, 'box', align='center')
        for col in 'BCDEF': cell(ws, '%s%d' % (col, r), None, 'in', DATE if col == 'B' else (INR if col == 'F' else None))
    d = MONTH1
    for i, row in enumerate([(d, 'Travel', 'Train Ludhiana → New Delhi (Shatabdi)', 'PNR 2451', 1265), (d, 'Local conveyance', 'Cab station → client office', 'OLA-881', 420),
                             (d, 'Lodging', 'Hotel, one night', 'H-1093', 3200), (d + datetime.timedelta(days=1), 'Food', 'Meals (2 days)', '—', 860),
                             (d + datetime.timedelta(days=1), 'Travel', 'Train New Delhi → Ludhiana', 'PNR 2477', 1265)], 10):
        for j, v in enumerate(row): ws.cell(i, 2 + j, v)
    dv(ws, 'C10:C34', '=Lists!$E$1:$E$7')
    cell(ws, 'E35', 'Total', 'soft', bold=True); cell(ws, 'F35', '=SUM(F10:F34)', 'soft', INR, bold=True)
    cell(ws, 'E36', 'Less advance', 'box'); cell(ws, 'F36', 5000, 'in', INR)
    cell(ws, 'E37', '=IF(F35-F36<0,"Refund to company","Payable to employee")', 'soft', bold=True); cell(ws, 'F37', '=ABS(F35-F36)', 'soft', INR, bold=True)
    head(ws, 9, ['By category', 'Amount'], col=8)
    for i, c in enumerate(CATS, 10):
        cell(ws, 'H%d' % i, c, 'box'); cell(ws, 'I%d' % i, '=SUMIF($C$10:$C$34,H%d,$F$10:$F$34)' % i, 'calc', INR)
    cell(ws, 'H17', 'Total', 'soft', bold=True); cell(ws, 'I17', '=SUM(I10:I16)', 'soft', INR, bold=True)
    ws.merge_cells('A39:F39'); cell(ws, 'A39', 'I confirm these expenses were incurred for official work and the bills are attached.', size=9, color=MUTED)
    for col, k in zip('BDF', ['Employee', 'Approved by (manager)', 'Accounts']): ws['%s43' % col] = '______________'; cell(ws, '%s44' % col, k, size=9)
    ws.print_area = 'A1:I45'
    protect(ws)
    howto(wb)
    save(wb, 'Expense-Claim-Form.xlsx')

# ------------------------------------------------------------------ Word documents
def doc_new(tid):
    d = Document()
    t = ITEM[tid]
    s = d.sections[0]
    s.page_height, s.page_width = Cm(29.7), Cm(21.0)
    s.left_margin = s.right_margin = Cm(2); s.top_margin = s.bottom_margin = Cm(1.8)
    st = d.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(11)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
    cp = d.core_properties
    cp.author, cp.last_modified_by, cp.title = 'Nivesh Sahayak', 'Nivesh Sahayak', t['name']
    cp.comments = 'From Toolkit by GroupChatOn (%s/%s/%s/). %s (%s)' % (SITE, t['kit'], t['id'], CREDIT, NS_URL)
    cp.keywords, cp.subject = 'Toolkit by GroupChatOn', t['desc']
    cp.created = cp.modified = datetime.datetime.now(); cp.revision = 1
    fp = s.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run('Template: Toolkit by GroupChatOn · toolkit.groupchaton.com · ' + CREDIT); r.font.size = Pt(7); r.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
    return d

def shade(cell_, color):
    tcPr = cell_._tc.get_or_add_tcPr()
    sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), color); tcPr.append(sh)

def para(d, text='', bold=False, size=None, align=None, color=None, space=4, italic=False):
    p = d.add_paragraph()
    if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'r': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space)
    if text:
        r = p.add_run(text); r.bold, r.italic = bold, italic
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = RGBColor.from_string(color)
    return p

def letterhead(d, name='[Your Business Name]', extra='[Address line] · [City, State PIN] · Phone [ ] · [email]', gst=True):
    para(d, name, bold=True, size=20, align='c', color=BLUE, space=0)
    para(d, extra, size=9, align='c', space=0)
    if gst: para(d, 'GSTIN: [ ]', size=9, align='c', space=2)
    p = para(d, space=8)
    pPr = p._p.get_or_add_pPr(); b = OxmlElement('w:pBdr'); bt = OxmlElement('w:bottom')
    for k, v in (('w:val', 'single'), ('w:sz', '12'), ('w:space', '1'), ('w:color', BLUE)): bt.set(qn(k), v)
    b.append(bt); pPr.append(b)

def table(d, headers, rows, widths=None, fill_rows=0):
    t = d.add_table(rows=1, cols=len(headers)); t.style = 'Table Grid'
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(10)
        shade(c, BLUE)
    for row in rows + [[''] * len(headers)] * fill_rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
            for p in cells[i].paragraphs:
                for r in p.runs: r.font.size = Pt(10)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Cm(w)
    para(d, space=2)
    return t

def kv(d, pairs, cols=2):
    t = d.add_table(rows=0, cols=cols * 2); t.style = 'Table Grid'
    for i in range(0, len(pairs), cols):
        cells = t.add_row().cells
        for j, (k, v) in enumerate(pairs[i:i + cols]):
            cells[2 * j].text = k; shade(cells[2 * j], 'E8F0FE'); cells[2 * j].paragraphs[0].runs[0].bold = True
            cells[2 * j + 1].text = v; shade(cells[2 * j + 1], 'FFF7D6')
            for c in (cells[2 * j], cells[2 * j + 1]):
                for r in c.paragraphs[0].runs: r.font.size = Pt(10)
    para(d, space=2)

def dsave(d, name):
    d.save(os.path.join(OUT, name)); print('  ', name)

def quotation():
    d = doc_new('quotation-delivery-challan-word')
    letterhead(d)
    para(d, 'QUOTATION', bold=True, size=15, align='c', space=6)
    kv(d, [('Quotation No.', 'Q-[001]'), ('Date', '[DD-MM-YYYY]'), ('To', '[Customer name]'), ('Valid until', '[DD-MM-YYYY]'),
           ('Address', '[Customer address]'), ('Contact', '[Name, phone]')])
    para(d, 'Dear Sir / Madam,'); para(d, 'Thank you for your enquiry. We are pleased to quote our best rates for the following:', space=6)
    table(d, ['#', 'Description', 'HSN/SAC', 'Qty', 'Unit', 'Rate (₹)', 'Amount (₹)'],
          [['1', '[Item or service description]', '[ ]', '[ ]', '[ ]', '[ ]', '[ ]'], ['2', '', '', '', '', '', '']], [1, 6.5, 2, 1.4, 1.4, 2.2, 2.5], fill_rows=4)
    kv(d, [('Sub-total', '₹ [ ]'), ('GST @ [ ]%', '₹ [ ]'), ('Freight / other', '₹ [ ]'), ('Total', '₹ [ ]')])
    para(d, 'Terms and conditions', bold=True, color=BLUE)
    for s in ['Prices are [inclusive / exclusive] of GST.', 'Payment: [100% advance / 50% advance, balance before dispatch / 30 days from invoice].',
              'Delivery: within [ ] days of confirmed order, [ex-works / delivered to your site].', 'This quotation is valid until the date mentioned above.',
              'Warranty: [as per manufacturer / [ ] months].']:
        d.add_paragraph(s, style='List Number')
    para(d, 'We look forward to your order.', space=18)
    para(d, 'For [Your Business Name]', bold=True, align='r', space=30); para(d, 'Authorised signatory', align='r', size=9)
    dsave(d, 'Quotation-Format.docx')
    d = doc_new('quotation-delivery-challan-word')
    letterhead(d)
    para(d, 'DELIVERY CHALLAN', bold=True, size=15, align='c', space=6)
    kv(d, [('Challan No.', 'DC-[001]'), ('Date', '[DD-MM-YYYY]'), ('Consignee', '[Name]'), ('Consignee GSTIN', '[ ]'),
           ('Delivery address', '[Address]'), ('Place of supply', '[State]'), ('Vehicle No.', '[ ]'), ('Transporter / LR No.', '[ ]'),
           ('Reason for movement', '[Job work / On approval / Branch transfer / Sale (invoice to follow) / Other]'), ('Reference (PO / invoice)', '[ ]')])
    table(d, ['#', 'Description of goods', 'HSN', 'Qty', 'Unit', 'Approx. value (₹)', 'Remarks'],
          [['1', '[Description]', '[ ]', '[ ]', '[ ]', '[ ]', '']], [1, 6.2, 1.8, 1.4, 1.4, 2.8, 2.4], fill_rows=7)
    para(d, 'Goods sent in good condition. Please check on receipt and sign below. This challan is not a tax invoice.', size=9, italic=True, space=24)
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = 'Received the above goods in good condition', 'For [Your Business Name]'
    t.rows[1].cells[0].text, t.rows[1].cells[1].text = '\n\nReceiver\'s signature, name and date', '\n\nAuthorised signatory'
    dsave(d, 'Delivery-Challan-Format.docx')

def letters():
    d = doc_new('offer-experience-letter-word')
    letterhead(d, '[Company Name]', '[Registered address] · [Phone] · [email] · [website]', gst=False)
    para(d, 'Ref: [HR/OL/2026/001]'); para(d, 'Date: [DD-MM-YYYY]', space=10)
    para(d, '[Candidate full name]', bold=True, space=0); para(d, '[Address]', space=10)
    para(d, 'Subject: Offer of employment for the position of [Designation]', bold=True, space=8)
    para(d, 'Dear [Candidate first name],')
    para(d, 'We are pleased to offer you the position of [Designation] in our [Department] department at [Company Name], on the terms below. '
            'We were impressed by your experience and believe you will be a valuable member of our team.', space=8)
    kv(d, [('Designation', '[ ]'), ('Department', '[ ]'), ('Reporting to', '[Manager name, designation]'), ('Place of work', '[City / office]'),
           ('Date of joining', '[DD-MM-YYYY]'), ('Working hours', '[10:00 am to 6:00 pm, Monday to Saturday]')], cols=1)
    para(d, 'Compensation', bold=True, color=BLUE)
    para(d, 'Your monthly gross salary will be ₹ [ ] (₹ [ ] per year), made up as follows:')
    table(d, ['Component', 'Monthly (₹)', 'Yearly (₹)'], [['Basic salary', '[ ]', '[ ]'], ['House rent allowance', '[ ]', '[ ]'], ['Other allowances', '[ ]', '[ ]'],
                                                          ['Gross salary', '[ ]', '[ ]'], ['Employer PF / ESI contribution', '[ ]', '[ ]'], ['Cost to company', '[ ]', '[ ]']], [8, 4, 4])
    para(d, 'Statutory deductions (PF, ESI, professional tax, income tax) will be made as per law.', size=10)
    para(d, 'Terms', bold=True, color=BLUE)
    for s in ['Probation: you will be on probation for [six] months from the date of joining. On satisfactory completion you will be confirmed in writing.',
              'Notice period: [one month] during probation and [two months] after confirmation, from either side, or salary in lieu of notice.',
              'Leave: as per company policy, currently [12 casual, 12 sick and 15 earned leave] days a year.',
              'Confidentiality: you will keep the company\'s business information confidential during and after your employment.',
              'This offer is subject to verification of your documents and references.']:
        d.add_paragraph(s, style='List Number')
    para(d, 'Documents to bring on the day of joining', bold=True, color=BLUE)
    for s in ['Proof of identity and address (Aadhaar card, PAN card)', 'Educational certificates and mark sheets', 'Relieving letter and last three payslips from your previous employer',
              'Two passport-size photographs', 'Bank account details (cancelled cheque)']:
        d.add_paragraph(s, style='List Bullet')
    para(d, 'Please sign and return a copy of this letter by [DD-MM-YYYY] to confirm your acceptance. We look forward to welcoming you.', space=18)
    para(d, 'For [Company Name]', bold=True, space=30); para(d, '[Name], [Designation]', space=18)
    para(d, 'Acceptance', bold=True, color=BLUE)
    para(d, 'I accept the offer on the terms above and will join on [DD-MM-YYYY].', space=24)
    para(d, 'Signature: ____________________     Name: ____________________     Date: ____________')
    dsave(d, 'Offer-Letter-Format.docx')
    d = doc_new('offer-experience-letter-word')
    letterhead(d, '[Company Name]', '[Registered address] · [Phone] · [email] · [website]', gst=False)
    para(d, 'Ref: [HR/EXP/2026/001]'); para(d, 'Date: [DD-MM-YYYY]', space=18)
    para(d, 'TO WHOM IT MAY CONCERN', bold=True, size=13, align='c', space=14)
    para(d, 'This is to certify that [Mr. / Ms.] [Employee full name] (Employee code [ ]) worked with [Company Name] from [DD-MM-YYYY] to [DD-MM-YYYY]. '
            'At the time of leaving, [he / she / they] held the position of [Designation] in the [Department] department.', space=8)
    para(d, 'During this period [his / her / their] main responsibilities included:')
    for s in ['[Responsibility 1, e.g. preparing GST returns and monthly accounts]', '[Responsibility 2]', '[Responsibility 3]']:
        d.add_paragraph(s, style='List Bullet')
    para(d, 'We found [him / her / them] sincere, hardworking and reliable. [His / Her / Their] conduct during the employment was [good / very good]. '
            '[He / She / They] has been relieved of all duties with effect from [DD-MM-YYYY] after completing the handover.', space=8)
    para(d, 'We wish [him / her / them] every success in the future.', space=30)
    para(d, 'For [Company Name]', bold=True, space=36); para(d, '[Authorised signatory name]', space=0); para(d, '[Designation]', space=0)
    dsave(d, 'Experience-Letter-Format.docx')

def minutes():
    d = doc_new('meeting-minutes-word')
    para(d, 'MINUTES OF MEETING', bold=True, size=18, color=BLUE, space=2)
    para(d, '[Meeting title]', bold=True, size=13, space=8)
    kv(d, [('Date', '[DD-MM-YYYY]'), ('Time', '[from – to]'), ('Venue / link', '[ ]'), ('Called by / chair', '[ ]'), ('Minutes by', '[ ]'), ('Next meeting', '[DD-MM-YYYY]')])
    para(d, 'Attendees', bold=True, color=BLUE); para(d, '[Name (role)], [Name (role)], [Name (role)]')
    para(d, 'Apologies (absent)', bold=True, color=BLUE); para(d, '[Name], [Name]', space=8)
    para(d, 'Agenda, discussion and decisions', bold=True, color=BLUE)
    table(d, ['#', 'Agenda item', 'Discussion (key points)', 'Decision'],
          [['1', 'Review of previous action items', '[ ]', '[ ]'], ['2', '[Agenda item]', '[ ]', '[ ]'], ['3', '[Agenda item]', '[ ]', '[ ]']], [1, 4, 7, 5], fill_rows=2)
    para(d, 'Action items', bold=True, color=BLUE)
    table(d, ['#', 'Action', 'Owner', 'Due date', 'Status'], [['1', '[What exactly will be done]', '[One person]', '[DD-MM-YYYY]', 'Open']], [1, 8, 3, 3, 2], fill_rows=5)
    para(d, 'Any other business', bold=True, color=BLUE); para(d, '[ ]', space=14)
    para(d, 'Minutes approved by: ____________________     Date: ____________')
    dsave(d, 'Meeting-Minutes-Template.docx')

def plugins():
    """more makers, one file per toolkit: scripts/makers/<kit-id>.py does `from make_files import *` and defines MAKERS = [functions]"""
    import glob, importlib.util
    out = []
    for f in sorted(glob.glob(os.path.join(HERE, 'scripts', 'makers', '*.py'))):
        s = importlib.util.spec_from_file_location('maker_' + os.path.basename(f)[:-3].replace('-', '_'), f); m = importlib.util.module_from_spec(s); s.loader.exec_module(m)
        out.append((os.path.basename(f)[:-3], m.MAKERS))
    return out

if __name__ == '__main__':
    # python -X utf8 scripts/make_files.py [--only <kit-id>]   (--only makes just that toolkit's files, so other files stay untouched)
    os.makedirs(OUT, exist_ok=True)
    only = sys.argv[sys.argv.index('--only') + 1] if '--only' in sys.argv else None
    if not only:
        for f in (gst_invoice, gst_summary, client_dues, bank_reco, stock_register, cash_book, attendance, marks, fee_register,
                  salary, leave_tracker, task_tracker, expense_claim, quotation, letters, minutes):
            f()
    for name, makers in plugins():
        if only and name != only: continue
        for f in makers: f()
    # every file named in the catalog exists
    for t in T:
        if only and t['kit'] != only: continue
        for fn, _ in t.get('files', []): assert os.path.exists(os.path.join(OUT, fn)), fn
    print('files made' + (' for ' + only if only else ''))
