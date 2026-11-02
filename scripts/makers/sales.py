"""Sales & CRM toolkit files: lead tracker, pipeline forecast, target vs achievement, customer list, price list, visit report + meeting notes.
   Global toolkit: amounts use plain number formats and the currency symbol from each workbook's Setup sheet (shown in the headings);
   dates use dd-mmm-yyyy (05-Nov-2026), which reads the same in every country."""
from make_files import *
import random
from openpyxl.workbook.defined_name import DefinedName

AMT, AMT2, DATE_G, MON = '#,##0', '#,##0.00', 'dd-mmm-yyyy', 'mmm-yyyy'
RED, AMBER, GREEN = PatternFill('solid', fgColor='FEE2E2'), PatternFill('solid', fgColor='FEF3C7'), PatternFill('solid', fgColor='DCFCE7')
D = lambda n: TODAY + datetime.timedelta(days=n)
PEOPLE = ['Anita', 'Karan', 'Divya', 'Rohit', 'Sara', 'Imran']

def note(ws, ref, text):
    cell(ws, ref, text, size=9, color=MUTED)

def lists(ws, row, col, heading, values, n):
    """an editable yellow list on a Setup sheet: heading in `row`, n input cells below, filled with values"""
    c = L(col)
    head(ws, row, [heading], col=col)
    for i in range(n):
        cell(ws, '%s%d' % (c, row + 1 + i), values[i] if i < len(values) else None, 'in')
    return '=Setup!$%s$%d:$%s$%d' % (c, row + 1, c, row + n)

def setting(ws, r, text, value, fmt=None, kind='in'):
    label(ws, 'A%d' % r, text, bold=False)
    cell(ws, 'B%d' % r, value, kind, fmt, align='left' if fmt in (DATE_G, MON) else None)

def put(ws, r0, rows, col=1):
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if v is not None: ws.cell(r0 + i, col + j, v)

def cf_row(ws, rng, formula, fill=None, font=None):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill, font=font))

# ------------------------------------------------------------------ 1. lead & follow-up tracker
def lead_tracker():
    wb = new_wb('sales-lead-tracker-excel')
    howto(wb, first=True)
    ld = sheet(wb, 'Leads', [12, 26, 18, 15, 26, 12, 20, 24, 13, 11, 14, 12, 13, 34, 13, 12, 10], tab=BLUE)
    sm = sheet(wb, 'Summary', [34, 11, 11, 11, 11, 15], tab=BLUE)
    su = sheet(wb, 'Setup', [44, 12, 3, 18, 24, 16, 30])
    # Setup
    title(su, 'Setup', 'Your currency and the lists used in the drop-downs. Yellow cells are yours to change.')
    setting(su, 4, 'Currency symbol (₹, $, €, £, AED …)', '₹')
    setting(su, 5, 'Show a follow-up as "Soon" within this many days', 2)
    note(su, 'A7', 'Stages: keep Won and Lost spelt exactly so; alerts and the win rate look for them.')
    people = lists(su, 9, 4, 'Salespeople', PEOPLE[:3], 30)
    sources = lists(su, 9, 5, 'Lead sources', ['Website', 'Phone enquiry', 'WhatsApp', 'Walk-in', 'Referral', 'Social media', 'Email campaign',
                                               'Trade show / exhibition', 'Partner / dealer', 'Online marketplace', 'Cold call'], 15)
    stages = lists(su, 9, 6, 'Stages', ['New', 'Contacted', 'Qualified', 'Proposal sent', 'Negotiation', 'Won', 'Lost'], 10)
    prods = lists(su, 9, 7, 'Products / services', ['Laptops', 'Printers', 'Projectors', 'POS billing systems', 'CCTV installation', 'Server and backup',
                                                    'Office furniture', 'Annual maintenance contract'], 50)
    CUR = 'Setup!$B$4'
    # Leads
    title(ld, 'Leads and follow-ups', 'One row per enquiry. Grey columns fill themselves. Data › Filter on "Follow-up status" shows today\'s calls.')
    head(ld, 3, ['Date added', 'Company / lead', 'Contact person', 'Phone', 'Email', 'City', 'Source', 'Interested in', '="Est. value ("&%s&")"' % CUR,
                 'Owner', 'Stage', 'Last contact', 'Next follow-up', 'Next action', 'Follow-up status', 'Days to follow-up (minus = late)', 'Lead age (days)'])
    grid(ld, 4, 503, 'ABCDEFGHIJKLMNOPQ', {'O': 'calc', 'P': 'calc', 'Q': 'calc'},
         {'A': DATE_G, 'L': DATE_G, 'M': DATE_G, 'I': AMT, 'P': '0;[Red]-0;0', 'Q': '0'},
         {'O': '=IF($B{r}="","",IF(OR($K{r}="Won",$K{r}="Lost"),"Closed",IF($M{r}="","No date set",IF($M{r}<TODAY(),"Overdue",'
               'IF($M{r}=TODAY(),"Today",IF($M{r}-TODAY()<=Setup!$B$5,"Soon","Upcoming"))))))',
          'P': '=IF(OR($B{r}="",$M{r}="",$O{r}="Closed"),"",$M{r}-TODAY())',
          'Q': '=IF(OR($A{r}="",$B{r}=""),"",TODAY()-$A{r})'})
    put(ld, 4, [
        (D(-21), 'Greenfield Hospital', 'Dr. Meera Nair', '98450 11223', 'purchase@greenfield.example', 'Kochi', 'Referral', 'Annual maintenance contract', 240000, 'Anita', 'Proposal sent', D(-4), D(-1), 'Call to discuss the revised AMC price'),
        (D(-14), 'Blue Ocean Logistics', 'Rahul Verma', '98200 33445', 'rahul@blueocean.example', 'Mumbai', 'Website', 'Laptops', 1150000, 'Karan', 'Negotiation', D(-2), D(0), 'Send final quote with 3-year warranty'),
        (D(-9), 'Sunrise Public School', 'Mrs. Harpreet Kaur', '98140 55667', 'office@sunrise.example', 'Ludhiana', 'Walk-in', 'Projectors', 270000, 'Anita', 'Qualified', D(-3), D(2), 'Demo in the school hall'),
        (D(-30), 'Metro Mart', 'Imran Shaikh', '99870 77889', None, 'Pune', 'Trade show / exhibition', 'POS billing systems', 180000, 'Divya', 'Won', D(-6), None, 'Delivered; ask for a referral'),
        (D(-25), 'Hilltop Resorts', 'Sanjay Thapa', '98160 99001', 'gm@hilltop.example', 'Shimla', 'Phone enquiry', 'CCTV installation', 95000, 'Karan', 'Lost', D(-10), None, 'Lost on price; try again next season'),
        (D(-5), 'Al Noor Trading LLC', 'Faisal Ahmed', '+971 50 000 0000', 'faisal@alnoor.example', 'Dubai', 'Email campaign', 'Office furniture', 420000, 'Divya', 'Contacted', D(-1), D(5), 'Share catalogue and price list'),
        (D(-2), 'Vista Interiors', 'Pooja Iyer', '98860 12121', None, 'Bengaluru', 'Social media', 'Printers', 66000, 'Anita', 'New', None, D(1), 'First call: understand the need'),
        (D(-40), 'Northstar Pharma', 'Vikas Rao', '98490 34343', 'vikas@northstar.example', 'Hyderabad', 'Partner / dealer', 'Server and backup', 850000, 'Karan', 'Proposal sent', D(-12), D(-6), 'Chase the proposal sent last month'),
        (D(-1), 'Green Leaf Cafe', 'Neha Joshi', '97690 45454', None, 'Mumbai', 'WhatsApp', 'POS billing systems', 35000, 'Divya', 'New', None, None, None)])
    dv(ld, 'G4:G503', sources); dv(ld, 'H4:H503', prods); dv(ld, 'J4:J503', people); dv(ld, 'K4:K503', stages)
    rng = 'A4:Q503'
    cf_row(ld, rng, '$O4="Overdue"', RED); cf_row(ld, rng, '$O4="Today"', AMBER)
    cf_row(ld, rng, '$K4="Won"', font=Font(color='15803D')); cf_row(ld, rng, '$K4="Lost"', font=Font(color='94A3B8'))
    ld.conditional_formatting.add('O4:O503', CellIsRule(operator='equal', formula=['"Overdue"'], font=Font(bold=True, color='B91C1C')))
    ld.freeze_panes = 'C4'; ld.auto_filter.ref = 'A3:Q503'
    # Summary
    title(sm, 'Lead summary', '="Position on "&TEXT(TODAY(),"dd-mmm-yyyy")')
    K, B, I, O, J, G = ('Leads!$%s$4:$%s$503' % (c, c) for c in 'KBIOJG')
    OPEN = '%s,"<>Won",%s,"<>Lost"' % (K, K)
    head(sm, 4, ['Key numbers', 'Value'])
    keys = [('Total leads', '=COUNTA(%s)' % B, None), ('Open leads', '=COUNTIFS(%s,"<>",%s)' % (B, OPEN), None),
            ('Overdue follow-ups', '=COUNTIF(%s,"Overdue")' % O, None), ('Follow-ups due today', '=COUNTIF(%s,"Today")' % O, None),
            ('Open leads with no follow-up date', '=COUNTIF(%s,"No date set")' % O, None),
            ('Won', '=COUNTIF(%s,"Won")' % K, None), ('Lost', '=COUNTIF(%s,"Lost")' % K, None),
            ('Win rate (won ÷ closed)', '=IF(B10+B11=0,"",B10/(B10+B11))', '0%'),
            ('="Open pipeline value ("&%s&")"' % CUR, '=SUMIFS(%s,%s,"<>",%s)' % (I, B, OPEN), AMT),
            ('="Won value ("&%s&")"' % CUR, '=SUMIF(%s,"Won",%s)' % (K, I), AMT)]
    for i, (k, f, fmt) in enumerate(keys, 5):
        cell(sm, 'A%d' % i, k, 'box'); cell(sm, 'B%d' % i, f, 'calc', fmt, bold=True, align='right')
    sm.column_dimensions['B'].width = 14
    head(sm, 16, ['Stage', 'Leads', '="Value ("&%s&")"' % CUR])
    for i in range(10):
        r = 17 + i
        cell(sm, 'A%d' % r, '=IF(Setup!F%d="","",Setup!F%d)' % (10 + i, 10 + i), 'box')
        cell(sm, 'B%d' % r, '=IF($A{r}="","",COUNTIF({K},$A{r}))'.format(r=r, K=K), 'calc', align='center')
        cell(sm, 'C%d' % r, '=IF($A{r}="","",SUMIF({K},$A{r},{I}))'.format(r=r, K=K, I=I), 'calc', AMT)
    head(sm, 28, ['Lead source', 'Leads', 'Won', 'Lost', 'Win rate', '="Won value ("&%s&")"' % CUR])
    for i in range(15):
        r = 29 + i
        cell(sm, 'A%d' % r, '=IF(Setup!E%d="","",Setup!E%d)' % (10 + i, 10 + i), 'box')
        cell(sm, 'B%d' % r, '=IF($A{r}="","",COUNTIF({G},$A{r}))'.format(r=r, G=G), 'calc', align='center')
        cell(sm, 'C%d' % r, '=IF($A{r}="","",COUNTIFS({G},$A{r},{K},"Won"))'.format(r=r, G=G, K=K), 'calc', align='center')
        cell(sm, 'D%d' % r, '=IF($A{r}="","",COUNTIFS({G},$A{r},{K},"Lost"))'.format(r=r, G=G, K=K), 'calc', align='center')
        cell(sm, 'E%d' % r, '=IF(N($C{r})+N($D{r})=0,"",$C{r}/($C{r}+$D{r}))'.format(r=r), 'calc', '0%', align='center')
        cell(sm, 'F%d' % r, '=IF($A{r}="","",SUMIFS({I},{G},$A{r},{K},"Won"))'.format(r=r, I=I, G=G, K=K), 'calc', AMT)
    head(sm, 45, ['Salesperson', 'Open leads', 'Overdue', 'Won', '="Won value ("&%s&")"' % CUR])
    for i in range(30):
        r = 46 + i
        cell(sm, 'A%d' % r, '=IF(Setup!D%d="","",Setup!D%d)' % (10 + i, 10 + i), 'box')
        cell(sm, 'B%d' % r, '=IF($A{r}="","",COUNTIFS({J},$A{r},{OPEN}))'.format(r=r, J=J, OPEN=OPEN), 'calc', align='center')
        cell(sm, 'C%d' % r, '=IF($A{r}="","",COUNTIFS({J},$A{r},{O},"Overdue"))'.format(r=r, J=J, O=O), 'calc', align='center')
        cell(sm, 'D%d' % r, '=IF($A{r}="","",COUNTIFS({J},$A{r},{K},"Won"))'.format(r=r, J=J, K=K), 'calc', align='center')
        cell(sm, 'E%d' % r, '=IF($A{r}="","",SUMIFS({I},{J},$A{r},{K},"Won"))'.format(r=r, I=I, J=J, K=K), 'calc', AMT)
    sm.conditional_formatting.add('C46:C75', CellIsRule(operator='greaterThan', formula=['0'], fill=RED))
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Leads by source'; ch.legend = None
    ch.add_data(Reference(sm, min_col=2, min_row=28, max_row=39), titles_from_data=True)
    ch.set_categories(Reference(sm, min_col=1, min_row=29, max_row=39))
    ch.series[0].graphicalProperties.solidFill = 'E11D48'
    ch.y_axis.majorGridlines = None; ch.x_axis.scaling.orientation = 'maxMin'
    ch.height, ch.width = 9, 15
    sm.add_chart(ch, 'H4')
    protect(sm)
    save(wb, 'Sales-Lead-Follow-up-Tracker.xlsx')

# ------------------------------------------------------------------ 2. pipeline & forecast
def pipeline():
    wb = new_wb('sales-pipeline-forecast-excel')
    howto(wb, first=True)
    dl = sheet(wb, 'Deals', [30, 24, 12, 14, 15, 11, 14, 12, 13, 32, 12, 20], tab=BLUE)
    fc = sheet(wb, 'Forecast', [40, 14, 16, 16, 16], tab=BLUE)
    su = sheet(wb, 'Setup', [34, 16, 3, 18])
    title(su, 'Setup', 'Currency, forecast window, stage probabilities and salespeople.')
    setting(su, 4, 'Currency symbol (₹, $, €, £, AED …)', '₹')
    setting(su, 5, 'Forecast starts in the month of', MONTH1, MON)
    head(su, 7, ['Stage', 'Win probability'])
    STG = [('Prospecting', .10), ('Qualified', .25), ('Proposal sent', .50), ('Negotiation', .75), ('Commit', .90), ('Won', 1), ('Lost', 0)]
    for i in range(10):
        r = 8 + i
        cell(su, 'A%d' % r, STG[i][0] if i < len(STG) else None, 'in'); cell(su, 'B%d' % r, STG[i][1] if i < len(STG) else None, 'in', '0%', align='center')
    note(su, 'A19', 'Keep Won and Lost spelt exactly so. Add new open stages above Won, in the order deals move.')
    people = lists(su, 7, 4, 'Salespeople', PEOPLE[:4], 30)
    CUR = 'Setup!$B$4'
    title(dl, 'Deals', 'One row per deal. Probability, weighted value and the warning fill themselves.')
    head(dl, 3, ['Deal', 'Customer', 'Owner', '="Value ("&%s&")"' % CUR, 'Stage', 'Win probability', '="Weighted value ("&%s&")"' % CUR,
                 'Created', 'Expected close', 'Next step', 'Next step date', 'Warning'])
    grid(dl, 4, 503, 'ABCDEFGHIJKL', {'F': 'calc', 'G': 'calc', 'L': 'calc'},
         {'D': AMT, 'F': '0%', 'G': AMT, 'H': DATE_G, 'I': DATE_G, 'K': DATE_G},
         {'F': '=IF(OR($A{r}="",$E{r}=""),"",IFERROR(INDEX(Setup!$B$8:$B$17,MATCH($E{r},Setup!$A$8:$A$17,0)),""))',
          'G': '=IF(OR($F{r}="",$D{r}=""),"",ROUND($D{r}*$F{r},0))',
          'L': '=IF(OR($A{r}="",$E{r}="Won",$E{r}="Lost"),"",IF(AND($I{r}<>"",$I{r}<TODAY()),"Close date passed",'
               'IF($K{r}="","No next step date",IF($K{r}<TODAY(),"Next step overdue",""))))'})
    put(dl, 4, [
        ('20 laptops for new branch', 'Blue Ocean Logistics', 'Karan', 1150000, 'Negotiation', None, None, D(-30), D(12), 'Final price call with the CFO', D(1)),
        ('Hospital AMC renewal', 'Greenfield Hospital', 'Anita', 240000, 'Proposal sent', None, None, D(-21), D(20), 'Revised AMC proposal', D(2)),
        ('Smart classroom projectors', 'Sunrise Public School', 'Anita', 270000, 'Qualified', None, None, D(-9), D(45), 'Demo in the school hall', D(4)),
        ('Server and backup upgrade', 'Northstar Pharma', 'Karan', 850000, 'Proposal sent', None, None, D(-40), D(-3), 'Chase the proposal', D(-2)),
        ('Office furniture fit-out', 'Al Noor Trading LLC', 'Divya', 420000, 'Prospecting', None, None, D(-5), D(75), 'Site visit and measurements', D(5)),
        ('POS for 3 outlets', 'Metro Mart', 'Divya', 180000, 'Won', None, None, D(-45), D(-6), 'Delivered', None),
        ('CCTV for the resort', 'Hilltop Resorts', 'Karan', 95000, 'Lost', None, None, D(-60), D(-20), 'Lost on price', None),
        ('Printers for design studio', 'Vista Interiors', 'Anita', 66000, 'Commit', None, None, D(-12), D(7), 'Collect the purchase order', None),
        ('Billing tablets', 'Green Leaf Cafe', 'Rohit', 35000, 'Prospecting', None, None, D(-1), D(30), 'First meeting', D(3)),
        ('Annual toner supply', 'City Diagnostics', 'Rohit', 310000, 'Negotiation', None, None, D(-18), D(38), 'Agree monthly delivery schedule', D(6)),
        ('Network cabling, 2nd floor', 'Orbit Software', 'Karan', 520000, 'Qualified', None, None, D(-10), D(100), 'Get the floor plan', D(8)),
        ('Laptops for interns', 'Orbit Software', 'Anita', 390000, 'Won', None, None, D(-35), D(-14), 'Delivered', None)])
    dv(dl, 'C4:C503', people); dv(dl, 'E4:E503', '=Setup!$A$8:$A$17')
    cf_row(dl, 'A4:L503', '$E4="Won"', font=Font(color='15803D')); cf_row(dl, 'A4:L503', '$E4="Lost"', font=Font(color='94A3B8'))
    dl.conditional_formatting.add('L4:L503', CellIsRule(operator='notEqual', formula=['""'], fill=RED, font=Font(bold=True, color='B91C1C')))
    dl.freeze_panes = 'B4'; dl.auto_filter.ref = 'A3:L503'
    # Forecast
    A, Dv, E, G, Ic = ('Deals!$%s$4:$%s$503' % (c, c) for c in 'ADEGI')
    OPEN = '%s,"<>",%s,"<>Won",%s,"<>Lost"' % (A, E, E)
    title(fc, 'Pipeline forecast', '="Six months from "&TEXT(Setup!$B$5,"mmm yyyy")&". Weighted = value × win probability of the stage."')
    head(fc, 3, ['Stage', 'Deals', '="Value ("&%s&")"' % CUR, 'Win probability', '="Weighted ("&%s&")"' % CUR])
    for i in range(10):
        r = 4 + i
        cell(fc, 'A%d' % r, '=IF(Setup!A%d="","",Setup!A%d)' % (8 + i, 8 + i), 'box')
        cell(fc, 'B%d' % r, '=IF($A{r}="","",COUNTIF({E},$A{r}))'.format(r=r, E=E), 'calc', align='center')
        cell(fc, 'C%d' % r, '=IF($A{r}="","",SUMIF({E},$A{r},{D}))'.format(r=r, E=E, D=Dv), 'calc', AMT)
        cell(fc, 'D%d' % r, '=IF($A{r}="","",Setup!B%d)'.format(r=r) % (8 + i), 'calc', '0%', align='center')
        cell(fc, 'E%d' % r, '=IF($A{r}="","",SUMIF({E},$A{r},{G}))'.format(r=r, E=E, G=G), 'calc', AMT)
    cell(fc, 'A14', 'Open pipeline (every stage except Won and Lost)', 'soft', bold=True)
    cell(fc, 'B14', '=COUNTIFS(%s)' % OPEN, 'soft', bold=True, align='center')
    cell(fc, 'C14', '=SUMIFS(%s,%s)' % (Dv, OPEN), 'soft', AMT, bold=True); cell(fc, 'D14', None, 'soft')
    cell(fc, 'E14', '=SUMIFS(%s,%s)' % (G, OPEN), 'soft', AMT, bold=True)
    head(fc, 17, ['Month (by expected close date)', 'Open deals', '="Pipeline value ("&%s&")"' % CUR, '="Weighted forecast ("&%s&")"' % CUR, '="Won ("&%s&")"' % CUR])
    for i in range(6):
        r = 18 + i
        win = '{I},">="&$A{r},{I},"<"&EDATE($A{r},1)'.format(I=Ic, r=r)
        cell(fc, 'A%d' % r, '=DATE(YEAR(Setup!$B$5),MONTH(Setup!$B$5)+%d,1)' % i, 'box', MON, align='left')
        cell(fc, 'B%d' % r, '=COUNTIFS(%s,%s)' % (win, OPEN), 'calc', align='center')
        cell(fc, 'C%d' % r, '=SUMIFS(%s,%s,%s)' % (Dv, win, OPEN), 'calc', AMT)
        cell(fc, 'D%d' % r, '=SUMIFS(%s,%s,%s)' % (G, win, OPEN), 'calc', AMT)
        cell(fc, 'E%d' % r, '=SUMIFS(%s,%s,%s,"Won")' % (Dv, win, E), 'calc', AMT)
    cell(fc, 'A24', 'Six-month total', 'soft', bold=True)
    for c in 'BCDE': cell(fc, '%s24' % c, '=SUM(%s18:%s23)' % (c, c), 'soft', None if c == 'B' else AMT, bold=True, align='center' if c == 'B' else None)
    cell(fc, 'A25', 'Open deals with a close date before this window', 'box')
    before = '%s,"<"&$A$18,%s' % (Ic, OPEN)
    cell(fc, 'B25', '=COUNTIFS(%s)' % before, 'calc', align='center'); cell(fc, 'C25', '=SUMIFS(%s,%s)' % (Dv, before), 'calc', AMT)
    cell(fc, 'D25', '=SUMIFS(%s,%s)' % (G, before), 'calc', AMT)
    head(fc, 28, ['Salesperson', 'Open deals', '="Pipeline value ("&%s&")"' % CUR, '="Weighted forecast ("&%s&")"' % CUR, '="Won ("&%s&")"' % CUR])
    C = 'Deals!$C$4:$C$503'
    for i in range(30):
        r = 29 + i
        cell(fc, 'A%d' % r, '=IF(Setup!D%d="","",Setup!D%d)' % (8 + i, 8 + i), 'box')
        own = '%s,$A%d' % (C, r)
        cell(fc, 'B%d' % r, '=IF($A{r}="","",COUNTIFS({own},{OPEN}))'.format(r=r, own=own, OPEN=OPEN), 'calc', align='center')
        cell(fc, 'C%d' % r, '=IF($A{r}="","",SUMIFS({D},{own},{OPEN}))'.format(r=r, D=Dv, own=own, OPEN=OPEN), 'calc', AMT)
        cell(fc, 'D%d' % r, '=IF($A{r}="","",SUMIFS({G},{own},{OPEN}))'.format(r=r, G=G, own=own, OPEN=OPEN), 'calc', AMT)
        cell(fc, 'E%d' % r, '=IF($A{r}="","",SUMIFS({D},{own},{E},"Won"))'.format(r=r, D=Dv, own=own, E=E), 'calc', AMT)
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Open pipeline by stage'
    ch.add_data(Reference(fc, min_col=3, min_row=3, max_row=8), titles_from_data=True)
    ch.add_data(Reference(fc, min_col=5, min_row=3, max_row=8), titles_from_data=True)
    ch.set_categories(Reference(fc, min_col=1, min_row=4, max_row=8))
    ch.series[0].graphicalProperties.solidFill = 'FDA4AF'; ch.series[1].graphicalProperties.solidFill = 'E11D48'
    ch.x_axis.scaling.orientation = 'maxMin'; ch.legend.position = 'b'
    ch.height, ch.width = 8, 15
    fc.add_chart(ch, 'G3')
    c2 = BarChart(); c2.type = 'col'; c2.style = 10; c2.title = 'Forecast by month'
    c2.add_data(Reference(fc, min_col=4, min_row=17, max_row=23), titles_from_data=True)
    c2.add_data(Reference(fc, min_col=5, min_row=17, max_row=23), titles_from_data=True)
    c2.set_categories(Reference(fc, min_col=1, min_row=18, max_row=23))
    c2.series[0].graphicalProperties.solidFill = 'E11D48'; c2.series[1].graphicalProperties.solidFill = '16A34A'
    c2.legend.position = 'b'; c2.x_axis.number_format = 'mmm-yy'
    c2.height, c2.width = 8, 15
    fc.add_chart(c2, 'G21')
    protect(fc)
    save(wb, 'Sales-Pipeline-Forecast.xlsx')

# ------------------------------------------------------------------ 3. target vs achievement
def targets():
    wb = new_wb('sales-target-achievement-excel')
    howto(wb, first=True)
    sc = sheet(wb, 'Scorecard', [16, 12, 14, 14, 9, 15, 15, 9, 15, 7, 11], tab=BLUE)
    tg = sheet(wb, 'Targets', [16, 12] + [11] * 12 + [13], tab=BLUE)
    sl = sheet(wb, 'Sales', [13, 16, 26, 14, 22, 14], tab=BLUE)
    ac = sheet(wb, 'Achievement', [16, 12] + [11] * 12 + [13])
    su = sheet(wb, 'Setup', [46, 18])
    fy = datetime.date(TODAY.year if TODAY.month >= 4 else TODAY.year - 1, 4, 1)
    title(su, 'Setup', 'Company, currency, sales year and report month.')
    setting(su, 3, 'Company name', 'Apex Office Solutions')
    setting(su, 4, 'Currency symbol (₹, $, €, £, AED …)', '₹')
    setting(su, 5, 'First month of the sales year (any date in it)', fy, DATE_G)
    setting(su, 6, 'Report month (any date in it)', MONTH1, MON)
    setting(su, 7, 'Report month is month number', '=MAX(1,MIN(12,(YEAR(B6)-YEAR(B5))*12+MONTH(B6)-MONTH(B5)+1))', kind='calc')
    setting(su, 8, '"Close" when year-to-date achievement is at least', 0.8, '0%')
    note(su, 'A10', 'Type 01-Jan for a calendar year or 01-Apr for an April to March year. Below 80% (or your level) shows as Behind.')
    CUR = 'Setup!$B$4'
    months = ['=DATE(YEAR(Setup!$B$5),MONTH(Setup!$B$5)+%d,1)' % i for i in range(12)]
    # Targets
    title(tg, 'Monthly sales targets', '="Targets in "&Setup!$B$4&" for each month of the sales year"')
    head(tg, 3, ['Salesperson', 'Region'] + months + ['Year total'])
    for c in range(3, 15): tg.cell(3, c).number_format = 'mmm-yy'
    grid(tg, 4, 33, 'ABCDEFGHIJKLMNO', {'O': 'calc'}, {c: AMT for c in 'CDEFGHIJKLMNO'}, {'O': '=IF($A{r}="","",SUM(C{r}:N{r}))'})
    TEAM = [('Anita', 'North', 400000), ('Karan', 'West', 500000), ('Divya', 'South', 350000), ('Rohit', 'East', 300000), ('Sara', 'Online', 250000), ('Imran', 'Middle East', 450000)]
    rnd = random.Random(11)
    for i, (n, reg, base) in enumerate(TEAM, 4):
        tg.cell(i, 1, n); tg.cell(i, 2, reg)
        for m in range(12): tg.cell(i, 3 + m, int(round(base * (1 + [0, .05, .05, .1, .15, .2, .25, .2, .1, .1, .15, .25][m]) / 10000.0)) * 10000)
    tg.freeze_panes = 'C4'
    # Sales
    title(sl, 'Sales', 'One row per invoice or order. Credit notes and returns as negative amounts.')
    head(sl, 3, ['Date', 'Salesperson', 'Customer', 'Invoice / order no.', 'Product', '="Amount ("&%s&")"' % CUR])
    grid(sl, 4, 2003, 'ABCDEF', {}, {'A': DATE_G, 'F': AMT})
    dv(sl, 'B4:B2003', '=Targets!$A$4:$A$33')
    CUSTS = ['Blue Ocean Logistics', 'Greenfield Hospital', 'Sunrise Public School', 'Metro Mart', 'Northstar Pharma', 'Orbit Software', 'City Diagnostics',
             'Al Noor Trading LLC', 'Vista Interiors', 'Kaveri Textiles', 'Prime Motors', 'Lotus Hotels']
    PRODS = ['Laptops', 'Printers', 'Projectors', 'POS billing systems', 'Toner and supplies', 'Annual maintenance contract', 'Networking']
    rows, no = [], 1001
    m = fy
    while m <= MONTH1:
        last = (datetime.date(m.year + m.month // 12, m.month % 12 + 1, 1) - datetime.timedelta(days=1)).day
        for n, reg, base in TEAM:
            perf = {'Anita': 1.05, 'Karan': .92, 'Divya': 1.1, 'Rohit': .7, 'Sara': .85, 'Imran': .98}[n]
            k = rnd.randint(3, 5)
            for _ in range(k):
                day = rnd.randint(1, last if m < MONTH1 else max(1, TODAY.day))
                amt = int(round(base * perf * rnd.uniform(.7, 1.3) / k / 100.0)) * 100
                rows.append((datetime.date(m.year, m.month, day), n, rnd.choice(CUSTS), 'INV-%d' % no, rnd.choice(PRODS), amt)); no += 1
        m = datetime.date(m.year + m.month // 12, m.month % 12 + 1, 1)
    rows.sort(key=lambda x: x[0])
    for i, r in enumerate(rows): r = list(r); r[3] = 'INV-%d' % (1001 + i); rows[i] = r
    put(sl, 4, rows[:1900])
    sl.freeze_panes = 'A4'; sl.auto_filter.ref = 'A3:F2003'
    # Achievement
    title(ac, 'Actual sales by month', 'Calculated from the Sales sheet. Nothing to type here.')
    head(ac, 3, ['Salesperson', 'Region'] + ['=Targets!%s3' % L(c) for c in range(3, 15)] + ['Year total'])
    for c in range(3, 15): ac.cell(3, c).number_format = 'mmm-yy'
    for i in range(30):
        r = 4 + i
        cell(ac, 'A%d' % r, '=IF(Targets!A%d="","",Targets!A%d)' % (r, r), 'box')
        cell(ac, 'B%d' % r, '=IF($A{r}="","",Targets!B{r}&"")'.format(r=r), 'box')
        for c in range(3, 15):
            col = L(c)
            cell(ac, '%s%d' % (col, r), '=IF($A{r}="","",SUMIFS(Sales!$F$4:$F$2003,Sales!$B$4:$B$2003,$A{r},Sales!$A$4:$A$2003,">="&{c}$3,Sales!$A$4:$A$2003,"<"&EDATE({c}$3,1)))'.format(r=r, c=col), 'calc', AMT)
        cell(ac, 'O%d' % r, '=IF($A{r}="","",SUM(C{r}:N{r}))'.format(r=r), 'calc', AMT)
    ac.freeze_panes = 'C4'
    protect(ac)
    # Scorecard
    sc['A1'] = '=Setup!B3&" · Sales target vs achievement"'; sc['A1'].font = Font(size=16, bold=True, color=BLUE)
    sc['A2'] = '="Report month: "&TEXT(Setup!B6,"mmmm yyyy")&"   ·   year to date from "&TEXT(Setup!B5,"mmmm yyyy")&"   ·   amounts in "&Setup!B4'
    sc['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(sc, 5, ['Salesperson', 'Region', 'Month target', 'Month sales', 'Month %', 'YTD target', 'YTD sales', 'YTD %', 'Gap to YTD target', 'Rank', 'Status'])
    for i in range(30):
        r, t = 6 + i, 4 + i
        f = {'A': '=IF(Targets!A%d="","",Targets!A%d)' % (t, t), 'B': '=IF($A{r}="","",Targets!B%d&"")' % t,
             'C': '=IF($A{r}="","",INDEX(Targets!$C%d:$N%d,Setup!$B$7))' % (t, t), 'D': '=IF($A{r}="","",INDEX(Achievement!$C%d:$N%d,Setup!$B$7))' % (t, t),
             'E': '=IF(OR($A{r}="",N($C{r})=0),"",$D{r}/$C{r})',
             'F': '=IF($A{r}="","",SUMPRODUCT((Targets!$C$3:$N$3<=Setup!$B$6)*Targets!$C%d:$N%d))' % (t, t),
             'G': '=IF($A{r}="","",SUMPRODUCT((Achievement!$C$3:$N$3<=Setup!$B$6)*Achievement!$C%d:$N%d))' % (t, t),
             'H': '=IF(OR($A{r}="",N($F{r})=0),"",$G{r}/$F{r})', 'I': '=IF($A{r}="","",MAX(0,$F{r}-$G{r}))',
             'J': '=IF($H{r}="","",RANK($H{r},$H$6:$H$35))', 'K': '=IF($H{r}="","",IF($H{r}>=1,"Achieved",IF($H{r}>=Setup!$B$8,"Close","Behind")))'}
        for col, fm in f.items():
            cell(sc, '%s%d' % (col, r), fm.format(r=r), 'box' if col in 'AB' else 'calc',
                 '0%' if col in 'EH' else (AMT if col in 'CDFGI' else None), align='center' if col in 'EHJK' else None)
    cell(sc, 'A36', 'Team total', 'soft', bold=True); cell(sc, 'B36', None, 'soft')
    for col in 'CDFGI': cell(sc, '%s36' % col, '=SUM(%s6:%s35)' % (col, col), 'soft', AMT, bold=True)
    cell(sc, 'E36', '=IF(N(C36)=0,"",D36/C36)', 'soft', '0%', bold=True, align='center')
    cell(sc, 'H36', '=IF(N(F36)=0,"",G36/F36)', 'soft', '0%', bold=True, align='center')
    for col in 'JK': cell(sc, '%s36' % col, None, 'soft')
    for txt, fill in (('Achieved', GREEN), ('Close', AMBER), ('Behind', RED)):
        sc.conditional_formatting.add('K6:K35', CellIsRule(operator='equal', formula=['"%s"' % txt], fill=fill))
    sc.conditional_formatting.add('H6:H35', CellIsRule(operator='greaterThanOrEqual', formula=['1'], font=Font(bold=True, color='15803D')))
    ch = BarChart(); ch.type = 'col'; ch.style = 10; ch.title = 'Year to date: target and sales'
    ch.add_data(Reference(sc, min_col=6, min_row=5, max_row=15), titles_from_data=True)
    ch.add_data(Reference(sc, min_col=7, min_row=5, max_row=15), titles_from_data=True)
    ch.set_categories(Reference(sc, min_col=1, min_row=6, max_row=15))
    ch.series[0].graphicalProperties.solidFill = 'CBD5E1'; ch.series[1].graphicalProperties.solidFill = 'E11D48'
    ch.legend.position = 'b'; ch.height, ch.width = 8, 22
    sc.add_chart(ch, 'A38')
    sc.freeze_panes = 'B6'
    protect(sc)
    save(wb, 'Sales-Target-vs-Achievement.xlsx')

# ------------------------------------------------------------------ 4. customer list
def customers():
    wb = new_wb('sales-customer-list-excel')
    howto(wb, first=True)
    cu = sheet(wb, 'Customers', [9, 26, 18, 15, 26, 13, 13, 12, 13, 13, 10, 8, 14, 13, 15, 28, 6], tab=BLUE)
    pu = sheet(wb, 'Purchases', [13, 26, 16, 28, 14], tab=BLUE)
    sm = sheet(wb, 'Summary', [8, 28, 15, 14, 15], tab=BLUE)
    su = sheet(wb, 'Setup', [42, 12, 3, 18, 18])
    title(su, 'Setup', 'Currency, status limits and the lists used in the drop-downs.')
    setting(su, 4, 'Currency symbol (₹, $, €, £, AED …)', '₹')
    setting(su, 5, 'Active if the last purchase was within (days)', 90)
    setting(su, 6, 'At risk up to (days); older is Lapsed', 180)
    segs = lists(su, 8, 4, 'Segments', ['Retail', 'Wholesale', 'Corporate', 'Online', 'Distributor', 'Government'], 15)
    owners = lists(su, 8, 5, 'Account owners', PEOPLE[:3], 30)
    CUR = 'Setup!$B$4'
    P = lambda c: 'Purchases!$%s$4:$%s$2003' % (c, c)
    title(cu, 'Customers', 'Grey columns come from the Purchases sheet. Sort or filter by Status to see who has gone quiet.')
    head(cu, 3, ['ID', 'Customer name', 'Contact person', 'Phone', 'Email', 'City / region', 'Segment', 'Account owner', 'Customer since', 'Last purchase',
                 'Days since last purchase', 'Orders', '="Total sales ("&%s&")"' % CUR, '="Average order ("&%s&")"' % CUR, 'Status', 'Notes', 'key'])
    f = {'L': '=IF($B{r}="","",COUNTIF(%s,$B{r}))' % P('B'),
         'J': '=IF(N($L{r})=0,"",SUMPRODUCT(MAX((%s=$B{r})*%s)))' % (P('B'), P('A')),
         'I': '=IF(N($L{r})=0,"",100000-SUMPRODUCT(MAX((%s=$B{r})*(%s<>"")*(100000-%s))))' % (P('B'), P('A'), P('A')),
         'K': '=IF($J{r}="","",TODAY()-$J{r})',
         'M': '=IF($B{r}="","",SUMIF(%s,$B{r},%s))' % (P('B'), P('E')),
         'N': '=IF(N($L{r})=0,"",$M{r}/$L{r})',
         'O': '=IF($B{r}="","",IF(N($L{r})=0,"No purchase yet",IF($K{r}<=Setup!$B$5,"Active",IF($K{r}<=Setup!$B$6,"At risk","Lapsed"))))',
         'Q': '=IF(N($M{r})=0,"",$M{r}+ROW()/1000000)'}
    grid(cu, 4, 303, 'ABCDEFGHIJKLMNOPQ', {c: 'calc' for c in 'IJKLMNOQ'}, {'I': DATE_G, 'J': DATE_G, 'K': '0', 'M': AMT, 'N': AMT, 'Q': '0.000000'}, f)
    cu.column_dimensions['Q'].hidden = True
    CL = [('C001', 'Blue Ocean Logistics', 'Rahul Verma', '98200 33445', 'rahul@blueocean.example', 'Mumbai', 'Corporate', 'Karan', 25, 6, 92000),
          ('C002', 'Greenfield Hospital', 'Dr. Meera Nair', '98450 11223', 'purchase@greenfield.example', 'Kochi', 'Corporate', 'Anita', 40, 9, 38000),
          ('C003', 'Sunrise Public School', 'Harpreet Kaur', '98140 55667', 'office@sunrise.example', 'Ludhiana', 'Government', 'Anita', 130, 3, 61000),
          ('C004', 'Metro Mart', 'Imran Shaikh', '99870 77889', None, 'Pune', 'Retail', 'Divya', 8, 11, 17500),
          ('C005', 'Northstar Pharma', 'Vikas Rao', '98490 34343', 'vikas@northstar.example', 'Hyderabad', 'Corporate', 'Karan', 210, 4, 125000),
          ('C006', 'Kaveri Textiles', 'S. Raghavan', '94430 56565', None, 'Coimbatore', 'Wholesale', 'Divya', 150, 5, 44000),
          ('C007', 'Al Noor Trading LLC', 'Faisal Ahmed', '+971 50 000 0000', 'faisal@alnoor.example', 'Dubai', 'Distributor', 'Divya', 60, 4, 210000),
          ('C008', 'Prime Motors', 'Ajay Malhotra', '98110 67676', None, 'Delhi', 'Corporate', 'Karan', 330, 2, 58000),
          ('C009', 'Lotus Hotels', 'Rina D\'Souza', '98220 78787', 'stores@lotus.example', 'Goa', 'Corporate', 'Anita', 95, 6, 33000),
          ('C010', 'ShopKart Online', 'Tanvi Shah', '98250 89898', None, 'Ahmedabad', 'Online', 'Divya', 3, 14, 9800),
          ('C011', 'City Diagnostics', 'Dr. Arun Bose', '98300 90909', None, 'Kolkata', 'Corporate', 'Karan', 400, 1, 27000),
          ('C012', 'Vista Interiors', 'Pooja Iyer', '98860 12121', None, 'Bengaluru', 'Retail', 'Anita', None, 0, 0)]
    rnd, purch = random.Random(5), []
    for c in CL:
        put(cu, 4 + CL.index(c), [c[:8]])
        last, n, avg = c[8], c[9], c[10]
        for k in range(n):
            purch.append((D(-(last + k * rnd.randint(18, 45))), c[1], None, rnd.choice(['Laptops', 'Printers', 'Toner and supplies', 'Projectors', 'Networking', 'Service visit']),
                          int(round(avg * rnd.uniform(.6, 1.4) / 100.0)) * 100))
    purch.sort(key=lambda x: x[0])
    put(pu, 4, [(p[0], p[1], 'INV-%d' % (2001 + i), p[3], p[4]) for i, p in enumerate(purch)])
    dv(cu, 'G4:G303', segs); dv(cu, 'H4:H303', owners)
    for txt, fill in (('Active', GREEN), ('At risk', AMBER), ('Lapsed', RED)):
        cu.conditional_formatting.add('O4:O303', CellIsRule(operator='equal', formula=['"%s"' % txt], fill=fill))
    cu.freeze_panes = 'C4'; cu.auto_filter.ref = 'A3:P303'
    title(pu, 'Purchases', 'One row per invoice or order. Pick the customer from the list so the name matches exactly.')
    head(pu, 3, ['Date', 'Customer', 'Invoice / order no.', 'Product / description', '="Amount ("&%s&")"' % CUR])
    grid(pu, 4 + len(purch), 2003, 'ABCDE', {}, {'A': DATE_G, 'E': AMT})
    for r in range(4, 4 + len(purch)):
        for col in 'ABCDE': cell(pu, '%s%d' % (col, r), None, 'in', DATE_G if col == 'A' else (AMT if col == 'E' else None))
    dv(pu, 'B4:B2003', '=Customers!$B$4:$B$303')
    pu.freeze_panes = 'A4'; pu.auto_filter.ref = 'A3:E2003'
    # Summary
    title(sm, 'Customer summary', '="Position on "&TEXT(TODAY(),"dd-mmm-yyyy")')
    O, M, G, Q, B = ('Customers!$%s$4:$%s$303' % (c, c) for c in 'OMGQB')
    head(sm, 4, ['', 'Status', 'Customers', '="Total sales ("&%s&")"' % CUR])
    for i, s in enumerate(['Active', 'At risk', 'Lapsed', 'No purchase yet'], 5):
        cell(sm, 'B%d' % i, s, 'box'); cell(sm, 'C%d' % i, '=COUNTIF(%s,B%d)' % (O, i), 'calc', align='center')
        cell(sm, 'D%d' % i, '=SUMIF(%s,B%d,%s)' % (O, i, M), 'calc', AMT)
    cell(sm, 'B9', 'All customers', 'soft', bold=True); cell(sm, 'C9', '=SUM(C5:C8)', 'soft', bold=True, align='center'); cell(sm, 'D9', '=SUM(D5:D8)', 'soft', AMT, bold=True)
    for txt, fill in (('Active', GREEN), ('At risk', AMBER), ('Lapsed', RED)):
        sm.conditional_formatting.add('B5:B8', CellIsRule(operator='equal', formula=['"%s"' % txt], fill=fill))
    head(sm, 11, ['', 'Segment', 'Customers', '="Total sales ("&%s&")"' % CUR, 'Share of sales'])
    for i in range(15):
        r = 12 + i
        cell(sm, 'B%d' % r, '=IF(Setup!D%d="","",Setup!D%d)' % (9 + i, 9 + i), 'box')
        cell(sm, 'C%d' % r, '=IF($B{r}="","",COUNTIF({G},$B{r}))'.format(r=r, G=G), 'calc', align='center')
        cell(sm, 'D%d' % r, '=IF($B{r}="","",SUMIF({G},$B{r},{M}))'.format(r=r, G=G, M=M), 'calc', AMT)
        cell(sm, 'E%d' % r, '=IF(OR($B{r}="",N($D$9)=0),"",$D{r}/$D$9)'.format(r=r), 'calc', '0%', align='center')
    head(sm, 28, ['Rank', 'Top customers', '="Total sales ("&%s&")"' % CUR, 'Last purchase', 'Status'])
    for k in range(1, 11):
        r = 28 + k
        mt = 'MATCH(LARGE(%s,%d),%s,0)' % (Q, k, Q)
        cell(sm, 'A%d' % r, k, 'box', align='center')
        cell(sm, 'B%d' % r, '=IFERROR(INDEX(%s,%s),"")' % (B, mt), 'calc')
        cell(sm, 'C%d' % r, '=IFERROR(INDEX(%s,%s),"")' % (M, mt), 'calc', AMT)
        cell(sm, 'D%d' % r, '=IFERROR(INDEX(Customers!$J$4:$J$303,%s),"")' % mt, 'calc', DATE_G, align='center')
        cell(sm, 'E%d' % r, '=IFERROR(INDEX(%s,%s),"")' % (O, mt), 'calc', align='center')
    protect(sm)
    save(wb, 'Sales-Customer-List.xlsx')

# ------------------------------------------------------------------ 5. price list
def price_list():
    wb = new_wb('sales-price-list-excel')
    howto(wb, first=True)
    pl = sheet(wb, 'Price List', [12, 34, 16, 16, 14, 9, 16, 10, 5], tab=BLUE)
    pr = sheet(wb, 'Products', [11, 32, 15, 8, 12, 12, 12, 10, 12, 13, 8, 14, 10, 13, 24, 5], tab=BLUE)
    su = sheet(wb, 'Setup', [44, 40, 3, 16])
    title(su, 'Setup', 'Company details for the printed list, currency, tax name and price levels.')
    for r, k, v, fmt in [(3, 'Company name', 'Apex Office Solutions', None), (4, 'Address', '2nd Floor, Trade Centre, MG Road, Pune 411001', None),
                         (5, 'Phone · email · website', '+91 98200 00000 · sales@apex.example', None), (6, 'Currency symbol (₹, $, €, £, AED …)', '₹', None),
                         (7, 'Prices effective from', TODAY, DATE_G), (8, 'Tax name (GST, VAT, sales tax …)', 'GST', None),
                         (9, 'Dealer discount off list price', 0.15, '0%'), (10, 'Distributor discount off list price', 0.25, '0%'),
                         (11, 'Round dealer / distributor prices to (0 = no rounding)', 1, None), (12, 'Warn if margin is below', 0.15, '0%')]:
        setting(su, r, k, v, fmt)
    head(su, 14, ['Status'], col=4)
    for i, s in enumerate(['Active', 'New', 'Discontinued'], 15): cell(su, 'D%d' % i, s, 'in')
    note(su, 'A14', 'Discontinued products stay on the Products sheet but are left out of the printed list.')
    CUR = 'Setup!$B$6'
    RND = lambda x: 'IF(N(Setup!$B$11)>0,ROUND(%s/Setup!$B$11,0)*Setup!$B$11,ROUND(%s,2))' % (x, x)
    title(pr, 'Products', 'Your master list with cost and margin (not printed). Grey columns calculate themselves.')
    head(pr, 3, ['Code', 'Product', 'Category', 'Unit', 'Pack / size', '="Cost price ("&%s&")"' % CUR, '="List price ("&%s&")"' % CUR, 'Margin %',
                 'Dealer price', 'Distributor price', '=Setup!$B$8&" %"', '="List price incl. "&Setup!$B$8', 'Min. order qty', 'Status', 'Notes', 'n'])
    f = {'H': '=IF(OR($B{r}="",N($G{r})=0),"",($G{r}-N($F{r}))/$G{r})',
         'I': '=IF(OR($B{r}="",$G{r}=""),"",%s)' % RND('$G{r}*(1-Setup!$B$9)'),
         'J': '=IF(OR($B{r}="",$G{r}=""),"",%s)' % RND('$G{r}*(1-Setup!$B$10)'),
         'L': '=IF(OR($B{r}="",$G{r}=""),"",ROUND($G{r}*(1+N($K{r})/100),2))',
         'P': '=IF(AND($B{r}<>"",$N{r}<>"Discontinued"),MAX(P$3:P{p})+1,"")'}
    for r in range(4, 404):
        for col in 'ABCDEFGHIJKLMNOP':
            fm = f.get(col)
            cell(pr, '%s%d' % (col, r), fm.format(r=r, p=r - 1) if fm else None, 'calc' if col in 'HIJLP' else 'in',
                 {'F': AMT2, 'G': AMT2, 'H': '0.0%', 'I': AMT2, 'J': AMT2, 'L': AMT2}.get(col))
    pr.column_dimensions['P'].hidden = True
    put(pr, 4, [('LP-101', 'Laptop, Core i5, 16 GB RAM, 512 GB SSD', 'Laptops', 'Nos', '1 unit', 52000, 61500, None, None, None, 18, None, 1, 'Active'),
                ('LP-102', 'Laptop, Core i7, 16 GB RAM, 1 TB SSD', 'Laptops', 'Nos', '1 unit', 68000, 79900, None, None, None, 18, None, 1, 'Active'),
                ('PR-201', 'Laser printer, mono, duplex, Wi-Fi', 'Printers', 'Nos', '1 unit', 14200, 17500, None, None, None, 18, None, 1, 'Active'),
                ('PR-202', 'Ink tank printer, colour, 3-in-1', 'Printers', 'Nos', '1 unit', 11800, 13200, None, None, None, 18, None, 1, 'Active'),
                ('TN-301', 'Toner cartridge for PR-201', 'Supplies', 'Pcs', 'Single', 1650, 2350, None, None, None, 18, None, 5, 'Active'),
                ('PJ-401', 'Projector, full HD, 4000 lumens', 'Projectors', 'Nos', '1 unit', 39000, 45000, None, None, None, 18, None, 1, 'New'),
                ('PJ-399', 'Projector, HD ready, 3000 lumens', 'Projectors', 'Nos', '1 unit', 27500, 29900, None, None, None, 18, None, 1, 'Discontinued'),
                ('NW-501', 'Network cable Cat6, 305 m box', 'Networking', 'Box', '305 m', 7400, 9800, None, None, None, 18, None, 2, 'Active'),
                ('SV-601', 'On-site installation and setup', 'Services', 'Job', 'Per visit', 1200, 2500, None, None, None, 18, None, 1, 'Active'),
                ('SV-602', 'Annual maintenance, per device', 'Services', 'Year', '12 months', 1800, 3000, None, None, None, 18, None, 10, 'Active')])
    dv(pr, 'N4:N403', '=Setup!$D$15:$D$17')
    pr.conditional_formatting.add('H4:H403', FormulaRule(formula=['AND($H4<>"",$H4<Setup!$B$12)'], fill=RED, font=Font(bold=True, color='B91C1C')))
    pr.conditional_formatting.add('A4:O403', FormulaRule(formula=['$N4="Discontinued"'], font=Font(color='94A3B8')))
    pr.freeze_panes = 'C4'; pr.auto_filter.ref = 'A3:O403'
    # printable list
    ws = pl
    for r, fm, sz in ((1, '=Setup!B3', 18), (2, '=Setup!B4', 10), (3, '=Setup!B5', 10)):
        ws.merge_cells('A%d:H%d' % (r, r)); cell(ws, 'A%d' % r, fm, bold=(r == 1), size=sz, align='center', color=BLUE if r == 1 else INK)
    label(ws, 'A5', 'Price level'); cell(ws, 'B5', 'List', 'in')
    dv(ws, 'B5', '"List,Dealer,Distributor"', 'List, Dealer or Distributor prices')
    label(ws, 'C5', 'Effective from'); cell(ws, 'D5', '=Setup!B7', fmt=DATE_G, align='left')
    label(ws, 'F5', 'Currency'); cell(ws, 'G5', '=Setup!B6', align='left')
    ws.merge_cells('A6:H6'); c = cell(ws, 'A6', '="PRICE LIST"&IF($B$5="List",""," · "&UPPER($B$5)&" PRICES")', bold=True, size=13, align='center', color='FFFFFF'); c.fill = F_HEAD
    head(ws, 8, ['Code', 'Product', 'Category', 'Unit / pack', '="Price ("&Setup!$B$6&")"', '=Setup!$B$8&" %"', '="Price incl. "&Setup!$B$8', 'Min. order'])
    for r in range(9, 409):
        k = 'Products!$%s$4:$%s$403'
        ix = lambda c: 'INDEX(%s,$I%d)' % (k % (c, c), r)
        cell(ws, 'I%d' % r, '=IFERROR(MATCH(ROW()-8,Products!$P$4:$P$403,0),"")', color='FFFFFF')
        fm = {'A': '=IF($I{r}="","",%s&"")' % ix('A'), 'B': '=IF($I{r}="","",%s&"")' % ix('B'), 'C': '=IF($I{r}="","",%s&"")' % ix('C'),
              'D': '=IF($I{r}="","",%s&IF(%s<>""," · "&%s,""))' % (ix('D'), ix('E'), ix('E')),
              'E': '=IF($I{r}="","",IF($B$5="Dealer",%s,IF($B$5="Distributor",%s,%s)))' % (ix('I'), ix('J'), ix('G')),
              'F': '=IF($I{r}="","",N(%s))' % ix('K'), 'G': '=IF(OR($I{r}="",$E{r}=""),"",ROUND($E{r}*(1+N($F{r})/100),2))',
              'H': '=IF($I{r}="","",IF(%s="","",%s))' % (ix('M'), ix('M'))}
        for col, f_ in fm.items():
            cell(ws, '%s%d' % (col, r), f_.format(r=r), fmt=AMT2 if col in 'EG' else None, align='center' if col in 'FH' else None)
    ws.conditional_formatting.add('A9:H408', FormulaRule(formula=['$I9<>""'], border=BOX))
    ws.conditional_formatting.add('A9:H408', FormulaRule(formula=['AND($I9<>"",MOD(ROW(),2)=0)'], fill=PatternFill('solid', fgColor='F8FAFC')))
    ws.column_dimensions['I'].hidden = True
    ws.print_title_rows = '8:8'
    ws.defined_names['_xlnm.Print_Area'] = DefinedName('_xlnm.Print_Area', attr_text="OFFSET('Price List'!$A$1,0,0,MAX(9,8+COUNT('Price List'!$I$9:$I$408)),8)")
    protect(ws)
    save(wb, 'Sales-Price-List.xlsx')

# ------------------------------------------------------------------ 6. Word: daily visit report + customer meeting notes
def visit_docs():
    d = doc_new('sales-visit-report-word')
    letterhead(d, '[Company Name]', '[Branch / office] · [Phone] · [email]', gst=False)
    para(d, 'DAILY SALES VISIT REPORT', bold=True, size=15, align='c', space=6)
    kv(d, [('Salesperson', '[Name]'), ('Date', '[DD-MM-YYYY]'), ('Territory / route', '[ ]'), ('Travel by', '[Bike / car / public transport]'),
           ('Start reading (km)', '[ ]'), ('End reading (km)', '[ ]'), ('Day started', '[time]'), ('Day ended', '[time]')])
    para(d, 'Visits and calls', bold=True, color=BLUE)
    table(d, ['#', 'Time', 'Customer and contact', 'Purpose', 'Outcome', 'Order value', 'Collected', 'Next action and date'],
          [['1', '[10:15]', '[Firm, contact person]', '[Demo / follow-up / collection]', '[What was agreed]', '[ ]', '[ ]', '[Action, DD-MM]']],
          [0.7, 1.3, 3.4, 2.0, 3.0, 1.8, 1.7, 3.1], fill_rows=9)
    para(d, 'Day summary', bold=True, color=BLUE)
    kv(d, [('Calls planned', '[ ]'), ('Calls made', '[ ]'), ('New customers met', '[ ]'), ('Orders booked (value)', '[ ]'),
           ('Collections (amount)', '[ ]'), ('Distance travelled (km)', '[ ]')])
    para(d, 'Market and competitor information', bold=True, color=BLUE)
    table(d, ['Competitor / brand', 'What we saw or heard (price, scheme, new product)', 'Where / which customer'], [['[ ]', '[ ]', '[ ]']], [4, 9, 4], fill_rows=2)
    para(d, 'Customer feedback and complaints', bold=True, color=BLUE)
    para(d, '[Product feedback, delivery problems, payment issues or requests to pass to the office]', space=8)
    para(d, 'Expenses for the day', bold=True, color=BLUE)
    table(d, ['Item', 'Details', 'Amount'], [['Fuel / fare', '[ ]', '[ ]'], ['Food', '[ ]', '[ ]'], ['Other', '[ ]', '[ ]'], ['Total', '', '[ ]']], [4, 9, 4])
    para(d, 'Plan for tomorrow', bold=True, color=BLUE)
    table(d, ['#', 'Customer / area', 'Purpose', 'Time'], [['1', '[ ]', '[ ]', '[ ]']], [1, 7, 6, 3], fill_rows=3)
    para(d, 'Support needed from the office', bold=True, color=BLUE)
    para(d, '[Samples, quotation approval, a technical visit, stock confirmation …]', space=18)
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = '\n\n____________________', '\n\n____________________'
    t.rows[1].cells[0].text, t.rows[1].cells[1].text = 'Salesperson\'s signature', 'Manager\'s remarks and signature'
    dsave(d, 'Sales-Daily-Visit-Report.docx')

    d = doc_new('sales-visit-report-word')
    letterhead(d, '[Company Name]', '[Branch / office] · [Phone] · [email]', gst=False)
    para(d, 'CUSTOMER MEETING NOTES', bold=True, size=15, align='c', space=6)
    kv(d, [('Customer', '[Company name]'), ('Date', '[DD-MM-YYYY]'), ('Place / online', '[ ]'), ('Meeting type', '[First meeting / Demo / Negotiation / Review]'),
           ('Our team', '[Names]'), ('Their team', '[Name, role]'), ('Deal / reference', '[ ]'), ('Notes by', '[Name]')])
    para(d, 'What we learned', bold=True, color=BLUE)
    table(d, ['Topic', 'Notes'], [
        ['Current situation: what do they use or do today?', ''],
        ['Need: what do they want to change, and why now?', ''],
        ['Budget: is money set aside? roughly how much?', ''],
        ['Decision: who decides, who else has a say, which approvals are needed?', ''],
        ['Timeline: by when do they need it? any deadline or event?', ''],
        ['Competition: who else are they talking to, and what do they like there?', ''],
        ['Success: how will they judge that the purchase worked?', '']], [6.5, 10.5])
    para(d, 'Products or services discussed', bold=True, color=BLUE)
    table(d, ['Product / service', 'Quantity / scope', 'Price discussed', 'Remarks'], [['[ ]', '[ ]', '[ ]', '[ ]']], [6, 3.5, 3, 4.5], fill_rows=2)
    para(d, 'Concerns raised and our answers', bold=True, color=BLUE)
    table(d, ['Concern', 'Our answer', 'Settled?'], [['[e.g. delivery time]', '[ ]', '[Yes / No]']], [6, 8.5, 2.5], fill_rows=2)
    para(d, 'Agreed next steps', bold=True, color=BLUE)
    table(d, ['#', 'Action', 'Owner (us / them)', 'Due date'], [['1', '[Send revised quotation]', '[ ]', '[DD-MM-YYYY]']], [1, 9, 4, 3], fill_rows=3)
    para(d, 'Our view', bold=True, color=BLUE)
    kv(d, [('Chance of winning', '[Low / Medium / High]'), ('Expected order value', '[ ]'), ('Expected decision date', '[DD-MM-YYYY]'), ('Next meeting', '[DD-MM-YYYY]')])
    para(d, 'Internal notes (do not send to the customer)', bold=True, color=BLUE)
    para(d, '[Pricing room, risks, people to involve from our side]', space=18)
    para(d, 'Shared with: ____________________     Date: ____________')
    dsave(d, 'Sales-Customer-Meeting-Notes.docx')

MAKERS = [lead_tracker, pipeline, targets, customers, price_list, visit_docs]
