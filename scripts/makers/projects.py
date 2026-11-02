"""Project Management toolkit (kit id `projects`): Gantt chart, RACI matrix, risk/issue/change log, budget vs actual (Excel),
   project charter, weekly status report, lessons learned and closure report (Word).
   Global toolkit: dates print as 03-Nov-2026, the currency is typed on each workbook's Settings sheet."""
from make_files import *
from openpyxl.styles.differential import DifferentialStyle

GD = 'dd-mmm-yyyy'                     # unambiguous in every country
NUM = '#,##0.00;[Red]-#,##0.00'
NUM0 = '#,##0;[Red]-#,##0'
PCT = '0%'
PURPLE = '7C3AED'
MON = 'CHOOSE(MONTH({x}),"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")'   # month label as text: works in every Excel language

def wdays(d, n):
    """d plus n working days, Monday to Friday (for the sample dates only; the workbook itself uses WORKDAY.INTL)."""
    while n > 0:
        d += datetime.timedelta(days=1)
        if d.weekday() < 5: n -= 1
    return d

def rule(ws, rng, formula, fill=None, font=None, border=None, stop=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=PatternFill('solid', fgColor=fill, bgColor=fill) if fill else None,
                                                   font=font, border=border, stopIfTrue=stop))

def note(ws, ref, text, span=None):
    if span: ws.merge_cells(span)
    c = cell(ws, ref, text, size=9, color=MUTED, wrap=True)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    return c

# ------------------------------------------------------------------ Gantt chart
WEEKENDS = [('Saturday and Sunday', '0000011'), ('Sunday only', '0000001'), ('Friday and Saturday', '0000110'),
            ('Friday only', '0000100'), ('No weekly off', '0000000')]

def gantt():
    wb = new_wb('pm-gantt-chart-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 26, 4, 60], tab=PURPLE)
    pl = sheet(wb, 'Plan', [6, 36, 11, 12, 13, 9, 13, 8, 11, 10] + [4.3] * 30, tab=PURPLE)
    hol = sheet(wb, 'Holidays', [16, 40])
    sm = sheet(wb, 'Summary', [30, 22, 4, 16, 10], tab=PURPLE)
    li = wb.create_sheet('Lists')
    for i, (k, v) in enumerate(WEEKENDS, 1): li.cell(i, 1, k); li.cell(i, 2, v)
    for i, k in enumerate(['Phase', 'Task', 'Milestone'], 1): li.cell(i, 3, k)
    for i, k in enumerate(['Weeks', 'Days'], 1): li.cell(i, 4, k)
    li.sheet_state = 'hidden'
    N1, N2 = 6, 155                                  # plan rows
    # Settings
    title(st, 'Project settings', 'Yellow cells: type or pick. The Plan and Summary sheets use these.')
    rows = [(3, 'Project name', 'Company website redesign', None), (4, 'Project manager', 'Priya Nair', None), (5, 'Weekly off days', 'Saturday and Sunday', None),
            (6, 'Timeline starts on', '=IF(COUNT(Plan!$E$6:$E$155)=0,TODAY()-WEEKDAY(TODAY(),3),MIN(Plan!$E$6:$E$155)-WEEKDAY(MIN(Plan!$E$6:$E$155),3))', GD),
            (7, 'Timeline scale', 'Weeks', None)]
    for r, k, v, f in rows:
        label(st, 'A%d' % r, k); cell(st, 'B%d' % r, v, 'in', f, align='left')
    dv(st, 'B5', '=Lists!$A$1:$A$5', 'Days nobody works. End dates skip them.'); dv(st, 'B7', '=Lists!$D$1:$D$2', 'Weeks for long projects, Days for detail')
    note(st, 'D5', 'End dates skip these days and every date on the Holidays sheet.')
    note(st, 'D6', 'Starts on the Monday before the first task. Type any date to move the chart (for example to see later months).')
    note(st, 'D7', 'Weeks: each timeline column is 7 days (30 weeks shown). Days: each column is one day (30 days shown).')
    label(st, 'A9', 'Weekend code (used by formulas)'); cell(st, 'B9', '=IFERROR(VLOOKUP(B5,Lists!$A$1:$B$5,2,FALSE),"0000011")', 'calc', align='left')
    label(st, 'A10', 'Days per timeline column'); cell(st, 'B10', '=IF(B7="Days",1,7)', 'calc', align='left')
    protect(st)
    # Holidays
    title(hol, 'Holidays and office closures', 'One date per row. End dates on the Plan skip these days.')
    head(hol, 3, ['Date', 'Holiday'])
    grid(hol, 4, 103, 'AB', {}, {'A': GD})
    y = TODAY.year
    for i, (d, n) in enumerate([(datetime.date(y, 12, 25), 'Christmas Day'), (datetime.date(y + 1, 1, 1), 'New Year\'s Day'),
                                (datetime.date(y + 1, 1, 26), 'Republic Day (India): replace with your own holidays')], 4):
        hol.cell(i, 1, d); hol.cell(i, 2, n)
    hol.freeze_panes = 'A4'
    # Plan
    ws = pl
    title(ws, 'Project plan and Gantt chart', 'Yellow: type. Grey: works itself out. Phase rows: leave Start, Duration and % done empty; they come from the tasks below.')
    cell(ws, 'A3', '="Project: "&Settings!B3&"   ·   Manager: "&Settings!B4&"   ·   Weekly off: "&Settings!B5', bold=True, color=INK)
    head(ws, 5, ['ID', 'Task, milestone or phase', 'Type', 'Owner', 'Start', 'Duration (work days)', 'End', '% done', 'Status', 'Depends on'])
    TL = [L(c) for c in range(11, 41)]                          # K..AN timeline
    for i, col in enumerate(TL):
        c = ws['%s5' % col]
        c.value = '=Settings!$B$6' if i == 0 else '=%s5+$AO$3' % TL[i - 1]
        c.number_format, c.font, c.fill, c.border = 'dd-mmm-yy', Font(bold=True, color='FFFFFF', size=8), F_HEAD, BOX
        c.alignment = Alignment(text_rotation=90, horizontal='center', vertical='center')
        m = ws['%s4' % col]
        mon = MON.format(x='%s5' % col)                         # text, so the label runs over the empty cells beside it
        m.value = '=' + mon if i == 0 else '=IF(MONTH(%s5)<>MONTH(%s5),%s,"")' % (col, TL[i - 1], mon)
        m.font = Font(bold=True, size=8, color=PURPLE)
    ws.row_dimensions[5].height = 56
    # helper cells and columns (hidden): AO3 step, AP3 weekend code; AO bar start, AP bar end, AQ done share, AR child rows, AS open milestone date
    ws['AO3'], ws['AP3'] = '=Settings!$B$10', '=Settings!$B$9'
    for col, lab in zip(['AO', 'AP', 'AQ', 'AR', 'AS'], ['Bar start', 'Bar end', 'Done share', 'Rows in phase', 'Open milestone']):
        ws['%s5' % col] = lab
        ws.column_dimensions[col].hidden = True
    WD = 'Settings!$B$9,Holidays!$A$4:$A$103'
    for r in range(N1, N2 + 1):
        for col in 'ABCDEFHJ':
            cell(ws, '%s%d' % (col, r), None, 'in', GD if col == 'E' else (PCT if col == 'H' else None))
        rng = lambda c: '%s%d:INDEX(%s$%d:%s$%d,%d+$AR%d)' % (c, r + 1, c, N1, c, N2 + 1, r - N1 + 1, r)
        cell(ws, 'G%d' % r, '=IF($C{r}="Phase",$AP{r},IF(OR($E{r}="",$F{r}=""),"",IF(N($F{r})<=0,$E{r},_xlfn.WORKDAY.INTL($E{r},$F{r}-1,{wd}))))'.format(r=r, wd=WD), 'calc', GD)
        cell(ws, 'I%d' % r, '=IF($B{r}="","",IF($AP{r}="","Not planned",IF(N($AQ{r})>=1,"Done",IF($AP{r}<TODAY(),"Late",IF($AO{r}<=TODAY(),"In progress","Not started")))))'.format(r=r), 'calc')
        ws['AR%d' % r] = '=IF($C{r}<>"Phase",0,IFERROR(MATCH("Phase",$C{r1}:$C${e},0)-1,{left}))'.format(r=r, r1=r + 1, e=N2 + 1, left=N2 - r)
        ws['AO%d' % r] = ('=IF($C{r}="Phase",IF($AR{r}=0,"",IF(COUNT({E})=0,"",MIN({E}))),IF($E{r}="","",$E{r}))').format(r=r, E=rng('E'))
        ws['AP%d' % r] = ('=IF($C{r}="Phase",IF($AO{r}="","",MAX({G})),IF($G{r}="","",$G{r}))').format(r=r, G=rng('G'))
        ws['AQ%d' % r] = ('=IF($C{r}="Phase",IF($AR{r}=0,0,IF(SUM({F})=0,0,SUMPRODUCT({F},{H})/SUM({F}))),N($H{r}))').format(r=r, F=rng('F'), H=rng('H'))
        ws['AS%d' % r] = '=IF(AND($C{r}="Milestone",$AP{r}<>"",N($H{r})<1),$AP{r},"")'.format(r=r)
        for col in TL:
            c = ws['%s%d' % (col, r)]
            c.value = '=IF(AND($C{r}="Milestone",$AP{r}<>""),IF(AND($AP{r}>={c}$5,$AP{r}<{c}$5+$AO$3),"◆",""),"")'.format(r=r, c=col)
            c.font, c.alignment = Font(bold=True, color='D97706', size=10), Alignment(horizontal='center', vertical='center')
            c.border = Border(left=Side(style='hair', color='E2E8F0'), right=Side(style='hair', color='E2E8F0'))
    dv(ws, 'C%d:C%d' % (N1, N2), '=Lists!$C$1:$C$3', 'Phase = heading row, Task = work, Milestone = checkpoint (duration 0)')
    v = DataValidation(type='decimal', operator='between', formula1='0', formula2='1', allow_blank=True)
    v.error, v.errorTitle, v.prompt, v.showInputMessage = 'Type a percent between 0% and 100%.', '% done', 'Type 50% for half done, 100% when finished.', True
    ws.add_data_validation(v); v.add('H%d:H%d' % (N1, N2))
    base = TODAY - datetime.timedelta(days=TODAY.weekday() + 21)
    plan = [('1', 'Plan', 'Phase'), ('1.1', 'Project kick-off held', 'Milestone', 'Priya', 0, 0, 1), ('1.2', 'Requirements and sitemap agreed', 'Task', 'Priya', 0, 5, 1, '1.1'),
            ('1.3', 'Content inventory of the old site', 'Task', 'Sana', 3, 4, 1),
            ('2', 'Design', 'Phase'), ('2.1', 'Home page and key page mockups', 'Task', 'Arjun', 5, 8, 1, '1.2'), ('2.2', 'Mobile layouts and style guide', 'Task', 'Arjun', 10, 5, 0.8, '2.1'),
            ('2.3', 'Design approved by sponsor', 'Milestone', 'Priya', 16, 0, 0, '2.2'),
            ('3', 'Build', 'Phase'), ('3.1', 'Hosting and staging site set up', 'Task', 'Vikram', 8, 3, 1), ('3.2', 'Page templates built', 'Task', 'Vikram', 16, 10, 0, '2.3'),
            ('3.3', 'Page content written and edited', 'Task', 'Sana', 12, 12, 0.4, '1.3'), ('3.4', 'Contact form and analytics working', 'Task', 'Vikram', 26, 4, 0, '3.2'),
            ('4', 'Test and launch', 'Phase'), ('4.1', 'Tested on phones and browsers', 'Task', 'Meera', 30, 5, 0, '3.4'), ('4.2', 'Defects fixed, final review', 'Task', 'Vikram', 35, 4, 0, '4.1'),
            ('4.3', 'Staff trained to edit pages', 'Task', 'Sana', 36, 2, 0), ('4.4', 'Website live', 'Milestone', 'Priya', 40, 0, 0, '4.2'),
            ('4.5', 'Post-launch check and handover', 'Task', 'Priya', 41, 3, 0, '4.4')]
    for i, p in enumerate(plan, N1):
        ws['A%d' % i], ws['B%d' % i], ws['C%d' % i] = p[0], p[1], p[2]
        if len(p) > 3:
            ws['D%d' % i], ws['E%d' % i], ws['F%d' % i], ws['H%d' % i] = p[3], wdays(base, p[4]), p[5], p[6]
            if len(p) > 7: ws['J%d' % i] = p[7]
    A = 'A%d:J%d' % (N1, N2)
    rule(ws, A, '$C6="Phase"', fill='EDE9FE', font=Font(bold=True, color='3B0764'))
    rule(ws, A, '$I6="Late"', fill='FEE2E2', font=Font(color='991B1B'))
    rule(ws, A, '$I6="Done"', font=Font(color='64748B'))
    T = 'K%d:AN%d' % (N1, N2)
    bar = 'AND($C6<>"Milestone",$AO6<>"",$AP6<>"",$AO6<K$5+$AO$3,$AP6>=K$5)'
    done = '$AO6+($AP6-$AO6+1)*N($AQ6)>MAX(K$5,$AO6)'                  # the done part of the bar overlaps this column
    red = Side(style='medium', color='DC2626')
    rule(ws, T, 'AND(%s,$C6="Phase",%s)' % (bar, done), fill='3B0764')
    rule(ws, T, 'AND(%s,$C6="Phase")' % bar, fill='A78BFA')
    rule(ws, T, 'AND(%s,%s)' % (bar, done), fill='6D28D9')
    rule(ws, T, 'AND(%s,$I6="Late")' % bar, fill='F87171')
    rule(ws, T, bar, fill='C4B5FD')
    rule(ws, T, 'AND(TODAY()>=K$5,TODAY()<K$5+$AO$3)', fill='FEE2E2', border=Border(left=red, right=red))
    rule(ws, T, 'AND($AO$3=1,MID($AP$3,WEEKDAY(K$5,2),1)="1")', fill='F1F5F9')
    rule(ws, 'K5:AN5', 'AND(TODAY()>=K$5,TODAY()<K$5+$AO$3)', fill='DC2626')
    ws.freeze_panes = 'C6'
    ws.print_title_rows = '5:5'
    ws.page_setup.orientation = 'landscape'
    ws.print_area = 'A1:AN%d' % N2
    # Summary
    ws = sm
    title(ws, 'Project summary', 'Everything here works itself out from the Plan sheet.')
    P = lambda c: 'Plan!$%s$%d:$%s$%d' % (c, N1, c, N2)
    items = [('Project', '=Settings!B3', None), ('Project manager', '=Settings!B4', None),
             ('Planned start', '=IF(COUNT(%s)=0,"",MIN(%s))' % (P('AO'), P('AO')), GD), ('Planned finish', '=IF(COUNT(%s)=0,"",MAX(%s))' % (P('AP'), P('AP')), GD),
             ('Working days in the plan', '=IF(B5="","",_xlfn.NETWORKDAYS.INTL(B5,B6,Settings!$B$9,Holidays!$A$4:$A$103))', '0'),
             ('Time gone (calendar)', '=IF(B5="","",MAX(0,MIN(1,(TODAY()-B5+1)/(B6-B5+1))))', PCT),
             ('Work done (weighted by duration)', '=IF(SUMIF(%s,"Task",%s)=0,"",SUMPRODUCT(--(%s="Task"),%s,%s)/SUMIF(%s,"Task",%s))' % (P('C'), P('F'), P('C'), P('F'), P('H'), P('C'), P('F')), PCT),
             ('Next open milestone', '=IF(COUNT(%s)=0,"None open",INDEX(%s,MATCH(MIN(%s),%s,0)))' % (P('AS'), P('B'), P('AS'), P('AS')), None),
             ('Milestone date', '=IF(COUNT(%s)=0,"",MIN(%s))' % (P('AS'), P('AS')), GD)]
    for i, (k, f, fmt) in enumerate(items, 3):
        label(ws, 'A%d' % i, k); cell(ws, 'B%d' % i, f, 'calc', fmt, bold=(k.startswith('Work done')), align='left')
    note(ws, 'A12', 'When work done is well behind time gone, the project is probably late even if no single task shows Late yet.', 'A12:B13')
    head(ws, 3, ['Tasks by status', 'Count'], col=4)
    for i, s in enumerate(['Done', 'In progress', 'Late', 'Not started', 'Not planned'], 4):
        cell(ws, 'D%d' % i, s, 'box'); cell(ws, 'E%d' % i, '=COUNTIFS(%s,"Task",%s,D%d)' % (P('C'), P('I'), i), 'calc', align='center')
    cell(ws, 'D9', 'All tasks', 'soft', bold=True); cell(ws, 'E9', '=COUNTIF(%s,"Task")' % P('C'), 'soft', bold=True, align='center')
    cell(ws, 'D10', 'Milestones done', 'box'); cell(ws, 'E10', '=COUNTIFS(%s,"Milestone",%s,"Done")&" of "&COUNTIF(%s,"Milestone")' % (P('C'), P('I'), P('C')), 'calc', align='center')
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Tasks by status'; ch.legend = None
    ch.add_data(Reference(ws, min_col=5, min_row=3, max_row=7), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=4, min_row=4, max_row=7))
    ch.series[0].graphicalProperties.solidFill = PURPLE
    ch.height, ch.width = 6.5, 13
    ws.add_chart(ch, 'A15')
    protect(ws)
    save(wb, 'PM-Gantt-Chart.xlsx')

# ------------------------------------------------------------------ RACI matrix
def raci():
    wb = new_wb('pm-raci-matrix-excel')
    howto(wb, first=True)
    R1, R2, NR = 6, 85, 12
    cols = [L(c) for c in range(3, 3 + NR)]                     # C..N roles
    ws = sheet(wb, 'RACI', [14, 40] + [10.5] * NR + [17], tab=PURPLE)
    gd = sheet(wb, 'Guide', [14, 90])
    title(ws, 'RACI matrix', 'Type the roles in the yellow cells of row 5, then pick R, A, C, I or A/R in each cell. One A per row.')
    cell(ws, 'A3', 'R = Responsible (does the work)   ·   A = Accountable (owns it, one per row)   ·   C = Consulted (asked before)   ·   I = Informed (told after)', size=9, color=MUTED)
    head(ws, 5, ['Phase', 'Activity, deliverable or decision'] + [''] * NR + ['Check'])
    roles = ['Sponsor', 'Project manager', 'Product owner', 'Designer', 'Developer', 'Content writer', 'QA tester', 'Marketing', 'IT support', 'Finance']
    for i, col in enumerate(cols):
        c = cell(ws, '%s5' % col, roles[i] if i < len(roles) else None, 'in', bold=True, wrap=True, align='center')
    ws.row_dimensions[5].height = 36
    chk = 'O'
    for r in range(R1, R2 + 1):
        cell(ws, 'A%d' % r, None, 'in'); cell(ws, 'B%d' % r, None, 'in', wrap=True)
        for col in cols: cell(ws, '%s%d' % (col, r), None, 'in', align='center', bold=True)
        rg = 'C{r}:N{r}'.format(r=r)
        a = '(COUNTIF({g},"A")+COUNTIF({g},"A/R"))'.format(g=rg); rr = '(COUNTIF({g},"R")+COUNTIF({g},"A/R"))'.format(g=rg)
        cell(ws, '%s%d' % (chk, r), '=IF(COUNTA({g})=0,IF($B{r}="","","Not assigned"),IF({a}=0,"No A",IF({a}>1,"More than one A",IF({rr}=0,"No R","OK"))))'.format(g=rg, r=r, a=a, rr=rr), 'calc', align='center')
    dv(ws, 'C%d:N%d' % (R1, R2), '"R,A,C,I,A/R"', 'R does the work, A owns it (one per row), C is asked, I is told')
    S = lambda *a: dict(zip(cols, a))
    rows = [('Start', 'Project charter approved', 'A', 'R', 'C', '', '', '', '', 'I', '', 'C'),
            ('Start', 'Budget approved', 'A', 'R', '', '', '', '', '', '', '', 'C'),
            ('Start', 'Requirements and sitemap', 'I', 'A', 'R', 'C', 'C', 'C', '', 'C', '', ''),
            ('Design', 'Page mockups', 'I', 'C', 'A', 'R', 'C', 'C', '', 'C', '', ''),
            ('Design', 'Design sign-off', 'A', 'R', 'C', 'C', 'I', '', '', 'I', '', ''),
            ('Build', 'Hosting and domain set-up', '', 'C', '', '', 'R', '', '', '', 'A', 'I'),
            ('Build', 'Page templates', '', 'C', 'A', 'C', 'R', '', 'I', '', '', ''),
            ('Build', 'Page content', '', 'C', 'A', '', '', 'R', '', 'C', '', ''),
            ('Build', 'Choose hosting provider', 'A', 'A', '', '', 'R', '', '', '', 'C', 'C'),
            ('Test', 'Testing on phones and browsers', '', 'I', 'C', 'C', 'C', '', 'A/R', '', '', ''),
            ('Launch', 'Go-live decision', 'A', 'R', 'C', '', 'C', '', 'C', 'I', 'I', ''),
            ('Launch', 'Staff training', '', 'A', 'C', '', 'C', 'R', '', '', 'I', ''),
            ('Launch', 'Launch announcement', 'I', 'C', 'C', 'C', '', 'C', '', 'A/R', '', ''),
            ('Close', 'Handover to support', 'I', 'A', '', '', 'R', '', '', '', 'R', ''),
            ('Close', 'Final cost report', 'I', 'A', '', '', '', '', '', '', '', 'R'),
            ('Close', 'Lessons learned meeting', 'C', 'A/R', 'C', 'C', 'C', 'C', 'C', 'C', 'C', 'I')]
    for i, row in enumerate(rows, R1):
        for j, v in enumerate(row):
            if v: ws.cell(i, j + 1, v)
    G = 'C%d:N%d' % (R1, R2)
    for val, fill, fc in (('A/R', 'DDD6FE', '4C1D95'), ('A', 'DDD6FE', '4C1D95'), ('R', 'DBEAFE', '1E40AF'), ('C', 'FEF3C7', '92400E'), ('I', 'F1F5F9', '475569')):
        ws.conditional_formatting.add(G, CellIsRule(operator='equal', formula=['"%s"' % val], fill=PatternFill('solid', fgColor=fill, bgColor=fill), font=Font(bold=True, color=fc)))
    rule(ws, 'O%d:O%d' % (R1, R2), 'O6="OK"', font=Font(bold=True, color='15803D'))
    rule(ws, 'O%d:O%d' % (R1, R2), 'AND(O6<>"",O6<>"OK")', fill='FEE2E2', font=Font(bold=True, color='B91C1C'))
    ws.freeze_panes = 'C6'
    # load per role
    L0 = R2 + 3
    cell(ws, 'A%d' % (L0 - 1), 'Load per role', bold=True, size=12, color=PURPLE)
    head(ws, L0, ['Role', '', 'R', 'A', 'C', 'I', 'Total', 'Share of all A', 'Note'])
    ws.merge_cells('A%d:B%d' % (L0, L0))
    ws.merge_cells('I%d:N%d' % (L0, L0))
    for i, col in enumerate(cols):
        r = L0 + 1 + i
        rng = '%s$%d:%s$%d' % (col, R1, col, R2)
        ws.merge_cells('A%d:B%d' % (r, r)); cell(ws, 'A%d' % r, '=IF(%s$5="","",%s$5)' % (col, col), 'box', bold=True)
        for c2, f in zip('CDEF', ['COUNTIF({g},"R")+COUNTIF({g},"A/R")', 'COUNTIF({g},"A")+COUNTIF({g},"A/R")', 'COUNTIF({g},"C")', 'COUNTIF({g},"I")']):
            cell(ws, '%s%d' % (c2, r), ('=IF($A{r}="","",' + f + ')').format(g=rng, r=r), 'calc', align='center')
        cell(ws, 'G%d' % r, '=IF($A{r}="","",COUNTA({g}))'.format(r=r, g=rng), 'calc', align='center')
        totA = '(COUNTIF($C${a}:$N${b},"A")+COUNTIF($C${a}:$N${b},"A/R"))'.format(a=R1, b=R2)
        cell(ws, 'H%d' % r, '=IF(OR($A{r}="",{t}=0),"",D{r}/{t})'.format(r=r, t=totA), 'calc', PCT, align='center')
        ws.merge_cells('I%d:N%d' % (r, r))
        cell(ws, 'I%d' % r, '=IF($A{r}="","",IF(G{r}=0,"Not used yet",IF(C{r}+D{r}=0,"Only consulted or informed: is this role needed?",IF(N(H{r})>0.4,"Accountable for a large share: delegate some A",IF(C{r}>12,"Many R: check the workload","")))))'.format(r=r), 'calc')
    ws.print_area = 'A1:O%d' % (L0 + NR)
    ws.page_setup.orientation = 'landscape'
    # guide
    title(gd, 'How RACI works', 'Share this sheet with the team at the kick-off meeting.')
    head(gd, 3, ['Letter', 'Meaning'])
    for i, (k, v) in enumerate([('R', 'Responsible: does the work. At least one per row; several people can share it.'),
                                ('A', 'Accountable: owns the result, approves it and answers for it. Exactly one person per row.'),
                                ('C', 'Consulted: asked for input before the work is done or the decision is made. Two-way conversation.'),
                                ('I', 'Informed: told about progress or the decision after it happens. One-way.'),
                                ('A/R', 'The same person both does the work and owns it. Common for small tasks.')], 4):
        cell(gd, 'A%d' % i, k, 'soft', bold=True, align='center'); cell(gd, 'B%d' % i, v, 'box', wrap=True)
    r = 10
    cell(gd, 'A%d' % r, 'Rules the Check column applies', bold=True, color=PURPLE, size=12); r += 1
    for s in ['Every row needs exactly one A. "No A" means nobody owns the result; "More than one A" means nobody really does.',
              'Every row needs at least one R, otherwise nobody does the work.',
              'An activity with no letters at all shows "Not assigned".']:
        cell(gd, 'B%d' % r, '• ' + s, wrap=True); gd.row_dimensions[r].height = 30; r += 1
    r += 1
    cell(gd, 'A%d' % r, 'Good practice', bold=True, color=PURPLE, size=12); r += 1
    for s in ['Push accountability down to the person closest to the work, not always to the manager.',
              'Too many Cs slow the project: every C is someone you must wait for.',
              'A role with only Is may not need to be in the project team; keep them on the report list instead.',
              'Go through the matrix with every person named in it and agree it, then keep it with the project charter.']:
        cell(gd, 'B%d' % r, '• ' + s, wrap=True); gd.row_dimensions[r].height = 30; r += 1
    protect(gd)
    save(wb, 'PM-RACI-Matrix.xlsx')

# ------------------------------------------------------------------ Risk, issue and change log
def raid():
    wb = new_wb('pm-risk-issue-log-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 14, 4, 22, 18, 14, 16, 12, 14, 16], tab=PURPLE)
    rk = sheet(wb, 'Risks', [7, 12, 42, 16, 12, 11, 11, 8, 10, 11, 34, 12, 14, 10, 10, 9, 12, 10], tab=PURPLE)
    hm = sheet(wb, 'Heat Map', [16, 14, 13, 13, 13, 13, 13])
    iss = sheet(wb, 'Issues', [7, 12, 42, 10, 12, 36, 12, 13, 12, 9, 10, 10], tab=PURPLE)
    chg = sheet(wb, 'Changes', [8, 12, 16, 34, 32, 12, 14, 11, 10, 12, 14, 12, 9, 26], tab=PURPLE)
    db = sheet(wb, 'Dashboard', [9, 46, 12, 9, 10, 4, 22, 12, 12], tab=PURPLE)
    R1, R2 = 6, 205
    # Settings and lists
    title(st, 'Settings and lists', 'Yellow cells can be changed. The lists feed the drop-downs on the other sheets.')
    label(st, 'A3', 'Currency (code or symbol)'); cell(st, 'B3', 'USD', 'in', align='left')
    label(st, 'A5', 'Medium from score'); cell(st, 'B5', 5, 'in', align='left')
    label(st, 'A6', 'High from score'); cell(st, 'B6', 10, 'in', align='left')
    label(st, 'A7', 'Critical from score'); cell(st, 'B7', 20, 'in', align='left')
    note(st, 'A8', 'Score = probability × impact (1 to 25). Below the Medium limit a risk is Low.', 'A8:B9')
    lists = {'D': ('Risk categories', ['Schedule', 'Cost', 'Scope', 'Quality', 'People and skills', 'Technical', 'Supplier', 'Legal and compliance', 'Stakeholders', 'External'], True),
             'E': ('Risk status', ['Open', 'Monitoring', 'Closed', 'Became an issue'], False),
             'F': ('Response', ['Avoid', 'Reduce', 'Transfer', 'Accept', 'Exploit (opportunity)'], False),
             'G': ('Issue status', ['Open', 'In progress', 'Resolved', 'Closed'], False),
             'H': ('Priority', ['High', 'Medium', 'Low'], False),
             'I': ('CR decision', ['Pending', 'Approved', 'Rejected', 'Deferred'], False),
             'J': ('Change type', ['Scope', 'Schedule', 'Cost', 'Quality', 'Resources', 'Other'], False)}
    for col, (h_, vals, edit) in lists.items():
        c = cell(st, '%s3' % col, h_, bold=True, color='FFFFFF'); c.fill = F_HEAD
        for i in range(20):
            cell(st, '%s%d' % (col, 4 + i), vals[i] if i < len(vals) else None, 'in' if edit else 'box')
    LST = lambda col, n: '=Settings!$%s$4:$%s$%d' % (col, col, 3 + n)
    # Risks
    ws = rk
    title(ws, 'Risk register', 'Write each risk as: because of [cause], [event] may happen, which would [effect].')
    head(ws, 5, ['ID', 'Raised on', 'Risk (cause → event → effect)', 'Category', 'Owner', 'Probability (1-5)', 'Impact (1-5)', 'Score', 'Rating', 'Response',
                 'Action to take', 'Action due', 'Status', 'Prob. after', 'Impact after', 'Score after', 'Last reviewed', 'Due check'])
    grid(ws, R1, R2, [L(c) for c in range(1, 19)], {'H': 'calc', 'I': 'calc', 'P': 'calc', 'R': 'calc'}, {'B': GD, 'L': GD, 'Q': GD},
         {'H': '=IF(OR($F{r}="",$G{r}=""),"",$F{r}*$G{r})',
          'I': '=IF($H{r}="","",IF($H{r}>=Settings!$B$7,"Critical",IF($H{r}>=Settings!$B$6,"High",IF($H{r}>=Settings!$B$5,"Medium","Low"))))',
          'P': '=IF(OR($N{r}="",$O{r}=""),"",$N{r}*$O{r})',
          'R': '=IF(OR($L{r}="",$M{r}="Closed",$M{r}="Became an issue"),"",IF($L{r}<TODAY(),"Overdue",""))'})
    for r in range(R1, R2 + 1):
        for col in 'ABDEFGHIJLMNOPQR': ws['%s%d' % (col, r)].alignment = Alignment(vertical='top', horizontal='center' if col in 'FGHNOP' else None)
        ws['C%d' % r].alignment = ws['K%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
        ws['S%d' % r] = '=IF(AND($H{r}<>"",OR($M{r}="Open",$M{r}="Monitoring")),$H{r}+(1000-ROW())/10000,"")'.format(r=r)
    ws.column_dimensions['S'].hidden = True
    one5 = '"1,2,3,4,5"'
    dv(ws, 'D%d:D%d' % (R1, R2), LST('D', 20)); dv(ws, 'J%d:J%d' % (R1, R2), LST('F', 5)); dv(ws, 'M%d:M%d' % (R1, R2), LST('E', 4))
    dv(ws, 'F%d:F%d' % (R1, R2), one5, '1 = very unlikely, 3 = possible, 5 = almost certain')
    dv(ws, 'G%d:G%d' % (R1, R2), one5, '1 = minor, 3 = noticeable delay or cost, 5 = project fails')
    dv(ws, 'N%d:O%d' % (R1, R2), one5)
    d = TODAY
    D = lambda n: d + datetime.timedelta(days=n)
    sample = [('R-01', D(-30), 'Because the sponsor is travelling in the design phase, design approval may be delayed, which would push back the build.', 'Stakeholders', 'Priya', 3, 4, 'Reduce', 'Book two review slots now; agree a deputy who can approve.', D(5), 'Open', 2, 4),
              ('R-02', D(-28), 'Because only one developer knows the hosting set-up, illness or leave may stop the build, which would delay go-live by weeks.', 'People and skills', 'Vikram', 2, 5, 'Reduce', 'Write a set-up note; pair a second developer for one day a week.', D(-2), 'Open', 2, 3),
              ('R-03', D(-25), 'Because content comes from five departments, pages may arrive late or incomplete, which would delay testing.', 'Schedule', 'Sana', 4, 3, 'Reduce', 'Weekly content deadline per department; send a page template with word limits.', D(9), 'Monitoring', 3, 3),
              ('R-04', D(-20), 'Because the hosting price is in a foreign currency, a currency change may raise the yearly cost, which would exceed the budget line.', 'Cost', 'Finance', 2, 2, 'Accept', 'Keep 5% contingency on the hosting line.', D(30), 'Open', 2, 2),
              ('R-05', D(-18), 'Because the old site has no redirects planned, search engines may drop old pages, which would lose enquiries after launch.', 'Technical', 'Vikram', 3, 4, 'Avoid', 'List every old address and map it to a new page before launch.', D(20), 'Open', 1, 4),
              ('R-06', D(-14), 'Because the design agency may be late, mockups may arrive after the review date, which would delay sign-off.', 'Supplier', 'Arjun', 3, 3, 'Transfer', 'Penalty clause for late delivery agreed in the purchase order.', D(-10), 'Became an issue', None, None)]
    for i, s in enumerate(sample, R1):
        for col, v in zip('ABCDEFGJKLMNO', s):
            if v is not None: ws['%s%d' % (col, i)] = v
        ws['Q%d' % i] = D(-3)
    A = 'A%d:R%d' % (R1, R2)
    rule(ws, 'H%d:I%d' % (R1, R2), '$I6="Critical"', fill='DC2626', font=Font(bold=True, color='FFFFFF'))
    rule(ws, 'H%d:I%d' % (R1, R2), '$I6="High"', fill='FB923C', font=Font(bold=True, color='431407'))
    rule(ws, 'H%d:I%d' % (R1, R2), '$I6="Medium"', fill='FDE047', font=Font(bold=True, color='422006'))
    rule(ws, 'H%d:I%d' % (R1, R2), '$I6="Low"', fill='86EFAC', font=Font(bold=True, color='14532D'))
    rule(ws, 'R%d:R%d' % (R1, R2), '$R6="Overdue"', fill='FEE2E2', font=Font(bold=True, color='B91C1C'))
    rule(ws, A, 'OR($M6="Closed",$M6="Became an issue")', font=Font(color='94A3B8'))
    ws.freeze_panes = 'D6'; ws.auto_filter.ref = 'A5:R%d' % R2
    ws.page_setup.orientation = 'landscape'
    # Heat map
    ws = hm
    title(ws, 'Risk heat map', 'Number of Open and Monitoring risks at each probability and impact.')
    ws['J1'], ws['J2'], ws['J3'] = '=Settings!B5', '=Settings!B6', '=Settings!B7'
    ws.column_dimensions['J'].hidden = True
    ws.merge_cells('C4:G4'); c = cell(ws, 'C4', 'Impact →', bold=True, align='center', color=PURPLE)
    for j, (n, t) in enumerate([(1, 'Minor'), (2, 'Small'), (3, 'Noticeable'), (4, 'Serious'), (5, 'Severe')]):
        cell(ws, '%s5' % L(3 + j), n, 'soft', bold=True, align='center'); cell(ws, '%s11' % L(3 + j), t, size=9, color=MUTED, align='center')
    cell(ws, 'A5', 'Probability ↓', bold=True, color=PURPLE)
    for i, (n, t) in enumerate([(5, 'Almost certain'), (4, 'Likely'), (3, 'Possible'), (2, 'Unlikely'), (1, 'Rare')]):
        r = 6 + i
        cell(ws, 'A%d' % r, t, size=9, color=MUTED); cell(ws, 'B%d' % r, n, 'soft', bold=True, align='center')
        ws.row_dimensions[r].height = 34
        for j in range(5):
            col = L(3 + j)
            f = '=COUNTIFS(Risks!$F$6:$F$205,$B{r},Risks!$G$6:$G$205,{c}$5,Risks!$M$6:$M$205,"Open")+COUNTIFS(Risks!$F$6:$F$205,$B{r},Risks!$G$6:$G$205,{c}$5,Risks!$M$6:$M$205,"Monitoring")'.format(r=r, c=col)
            c = cell(ws, '%s%d' % (col, r), f, 'box', '0;-0;""', bold=True, size=14, align='center')
    H = 'C6:G10'
    rule(ws, H, '$B6*C$5>=$J$3', fill='DC2626', font=Font(bold=True, color='FFFFFF'))
    rule(ws, H, '$B6*C$5>=$J$2', fill='FB923C')
    rule(ws, H, '$B6*C$5>=$J$1', fill='FDE047')
    rule(ws, H, '$B6*C$5<$J$1', fill='86EFAC')
    note(ws, 'A13', 'Colours follow the score limits on the Settings sheet: green Low, yellow Medium, orange High, red Critical. Work on the top-right corner first.', 'A13:G14')
    protect(ws)
    # Issues
    ws = iss
    title(ws, 'Issue log', 'Problems happening now. Every issue needs an owner, an action and a target date.')
    head(ws, 5, ['ID', 'Raised on', 'Issue', 'Priority', 'Owner', 'Action / resolution', 'Target date', 'Status', 'Resolved on', 'Days open', 'Due check', 'From risk'])
    grid(ws, R1, R2, [L(c) for c in range(1, 13)], {'J': 'calc', 'K': 'calc'}, {'B': GD, 'G': GD, 'I': GD},
         {'J': '=IF($B{r}="","",IF($I{r}<>"",$I{r},TODAY())-$B{r})',
          'K': '=IF(AND($G{r}<>"",OR($H{r}="Open",$H{r}="In progress"),$G{r}<TODAY()),"Overdue","")'})
    for r in range(R1, R2 + 1): ws['C%d' % r].alignment = ws['F%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
    dv(ws, 'D%d:D%d' % (R1, R2), LST('H', 3)); dv(ws, 'H%d:H%d' % (R1, R2), LST('G', 4))
    for i, s in enumerate([('I-01', D(-10), 'Design agency delivered mockups 6 days late; sign-off meeting moved.', 'High', 'Arjun', 'Agency working Saturday; new review date agreed with sponsor.', D(3), 'In progress', None, 'R-06'),
                           ('I-02', D(-6), 'Staging server password shared on chat.', 'Medium', 'Vikram', 'Password changed; access list reviewed.', D(-4), 'Resolved', D(-5), ''),
                           ('I-03', D(-4), 'Finance department has not sent its page content.', 'Medium', 'Sana', 'Escalate to finance manager at the weekly meeting.', D(-1), 'Open', None, 'R-03')], R1):
        for col, v in zip('ABCDEFGHIL', s):
            if v is not None: ws['%s%d' % (col, i)] = v
    rule(ws, 'A%d:L%d' % (R1, R2), '$K6="Overdue"', fill='FEE2E2')
    rule(ws, 'D%d:D%d' % (R1, R2), '$D6="High"', font=Font(bold=True, color='B91C1C'))
    rule(ws, 'A%d:L%d' % (R1, R2), 'OR($H6="Resolved",$H6="Closed")', font=Font(color='94A3B8'))
    ws.freeze_panes = 'D6'; ws.auto_filter.ref = 'A5:L%d' % R2
    ws.page_setup.orientation = 'landscape'
    # Changes
    ws = chg
    title(ws, 'Change request log', 'Every request to change scope, time, cost or quality, and what was decided.')
    head(ws, 5, ['CR no.', 'Raised on', 'Requested by', 'Change requested', 'Reason / benefit', 'Type', 'Cost impact', 'Days impact', 'Priority', 'Decision', 'Decided by', 'Decision date', 'Days waiting', 'Notes'])
    ws['G5'] = '="Cost impact ("&Settings!$B$3&")"'
    grid(ws, R1, R2, [L(c) for c in range(1, 15)], {'M': 'calc'}, {'B': GD, 'L': GD, 'G': NUM},
         {'M': '=IF($B{r}="","",IF($L{r}<>"",$L{r},TODAY())-$B{r})'})
    for r in range(R1, R2 + 1):
        ws['D%d' % r].alignment = ws['E%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
        ws['O%d' % r] = '=IF($J{r}="Pending",$M{r},"")'.format(r=r)
    ws.column_dimensions['O'].hidden = True
    dv(ws, 'F%d:F%d' % (R1, R2), LST('J', 6)); dv(ws, 'I%d:I%d' % (R1, R2), LST('H', 3)); dv(ws, 'J%d:J%d' % (R1, R2), LST('I', 4))
    for i, s in enumerate([('CR-01', D(-20), 'Marketing', 'Add a blog section with categories', 'Regular articles to bring search traffic', 'Scope', 1800, 6, 'Medium', 'Approved', 'Sponsor', D(-15), 'Templates in build task 3.2'),
                           ('CR-02', D(-12), 'Sales', 'Online chat widget on every page', 'Faster replies to enquiries', 'Scope', 600, 2, 'Low', 'Deferred', 'Sponsor', D(-9), 'After launch, phase 2'),
                           ('CR-03', D(-5), 'Sponsor', 'Hindi version of the five main pages', 'Customers in north India', 'Scope', 2400, 8, 'High', 'Pending', None, None, 'Translation quote awaited')], R1):
        for col, v in zip('ABCDEFGHIJKLN', s):
            if v is not None: ws['%s%d' % (col, i)] = v
    rule(ws, 'J%d:J%d' % (R1, R2), '$J6="Approved"', fill='DCFCE7', font=Font(bold=True, color='166534'))
    rule(ws, 'J%d:J%d' % (R1, R2), '$J6="Rejected"', fill='FEE2E2', font=Font(bold=True, color='991B1B'))
    rule(ws, 'J%d:J%d' % (R1, R2), '$J6="Pending"', fill='FEF3C7', font=Font(bold=True, color='92400E'))
    ws.freeze_panes = 'E6'; ws.auto_filter.ref = 'A5:N%d' % R2
    ws.page_setup.orientation = 'landscape'
    # Dashboard
    ws = db
    title(ws, 'Risks, issues and changes at a glance', 'Updates by itself from the three logs.')
    head(ws, 3, ['Open risks by rating', '', 'Count'])
    ws.merge_cells('A3:B3')
    RK = lambda c: 'Risks!$%s$%d:$%s$%d' % (c, R1, c, R2)
    for i, k in enumerate(['Critical', 'High', 'Medium', 'Low'], 4):
        ws.merge_cells('A%d:B%d' % (i, i)); cell(ws, 'A%d' % i, k, 'box')
        cell(ws, 'C%d' % i, '=COUNTIFS(%s,A%d,%s,"Open")+COUNTIFS(%s,A%d,%s,"Monitoring")' % (RK('I'), i, RK('M'), RK('I'), i, RK('M')), 'calc', align='center')
    ws.merge_cells('A8:B8'); cell(ws, 'A8', 'Risk actions overdue', 'box', bold=True); cell(ws, 'C8', '=COUNTIF(%s,"Overdue")' % RK('R'), 'calc', bold=True, align='center')
    rule(ws, 'A4:C4', '$C$4>0', fill='FEE2E2')
    IS = lambda c: 'Issues!$%s$%d:$%s$%d' % (c, R1, c, R2)
    head(ws, 3, ['Open issues', 'Count', 'Overdue'], col=7)
    for i, k in enumerate(['High', 'Medium', 'Low'], 4):
        cell(ws, 'G%d' % i, k, 'box')
        cell(ws, 'H%d' % i, '=COUNTIFS(%s,G%d,%s,"Open")+COUNTIFS(%s,G%d,%s,"In progress")' % (IS('D'), i, IS('H'), IS('D'), i, IS('H')), 'calc', align='center')
        cell(ws, 'I%d' % i, '=COUNTIFS(%s,G%d,%s,"Overdue")' % (IS('D'), i, IS('K')), 'calc', align='center')
    CH = lambda c: 'Changes!$%s$%d:$%s$%d' % (c, R1, c, R2)
    head(ws, 9, ['Change requests', 'Count', 'Days'], col=7)
    for i, k in enumerate(['Pending', 'Approved', 'Rejected', 'Deferred'], 10):
        cell(ws, 'G%d' % i, k, 'box'); cell(ws, 'H%d' % i, '=COUNTIF(%s,G%d)' % (CH('J'), i), 'calc', align='center')
        cell(ws, 'I%d' % i, '=SUMIF(%s,G%d,%s)' % (CH('J'), i, CH('H')), 'calc', align='center')
    cell(ws, 'G15', '="Cost of approved changes ("&Settings!$B$3&")"', 'soft', bold=True, wrap=True); ws.row_dimensions[15].height = 30
    ws.merge_cells('H15:I15'); cell(ws, 'H15', '=SUMIF(%s,"Approved",%s)' % (CH('J'), CH('G')), 'soft', NUM, bold=True)
    cell(ws, 'G16', 'Oldest pending (days)', 'box'); ws.merge_cells('H16:I16')
    cell(ws, 'H16', '=IF(COUNTIF(%s,"Pending")=0,"",MAX(%s))' % (CH('J'), CH('O')), 'calc', align='center')
    head(ws, 11, ['ID', 'Top open risks', 'Owner', 'Score', 'Rating'])
    for k in range(1, 6):
        r = 11 + k
        key = 'LARGE(Risks!$S$%d:$S$%d,%d)' % (R1, R2, k)
        pos = 'MATCH(%s,Risks!$S$%d:$S$%d,0)' % (key, R1, R2)
        for col, src in zip('ABCDE', 'ACEHI'):
            cell(ws, '%s%d' % (col, r), '=IFERROR(INDEX(Risks!$%s$%d:$%s$%d,%s),"")' % (src, R1, src, R2, pos), 'calc', wrap=(col == 'B'), align='center' if col in 'DE' else None)
        ws.row_dimensions[r].height = 44
        ws['B%d' % r].alignment = Alignment(wrap_text=True, vertical='center')
    rule(ws, 'E12:E16', '$E12="Critical"', fill='DC2626', font=Font(bold=True, color='FFFFFF'))
    rule(ws, 'E12:E16', '$E12="High"', fill='FB923C')
    rule(ws, 'E12:E16', '$E12="Medium"', fill='FDE047')
    rule(ws, 'E12:E16', '$E12="Low"', fill='86EFAC')
    protect(ws)
    save(wb, 'PM-Risk-Issue-Change-Log.xlsx')

# ------------------------------------------------------------------ Budget vs actual
COST_TYPES = ['Labour', 'Materials', 'Equipment', 'Software and licences', 'Services and contractors', 'Travel', 'Training', 'Contingency', 'Other']

def budget():
    wb = new_wb('pm-budget-vs-actual-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 30, 4, 26], tab=PURPLE)
    bd = sheet(wb, 'Budget', [34, 20, 13, 13, 13, 13, 14, 14, 13, 10, 12], tab=PURPLE)
    sp = sheet(wb, 'Spending', [13, 34, 34, 20, 14, 13, 12, 11], tab=PURPLE)
    mo = sheet(wb, 'Monthly', [14, 14, 14, 16, 12], tab=PURPLE)
    sm = sheet(wb, 'Summary', [24, 14, 14, 14, 14, 14, 10], tab=PURPLE)
    B1, B2, S1, S2 = 6, 45, 6, 405
    title(st, 'Settings', 'Type once. Every heading shows your currency.')
    for r, k, v, f in [(3, 'Project name', 'Company website redesign', None), (4, 'Currency (code or symbol)', 'USD', None),
                       (5, 'First month of the project', datetime.date(TODAY.year, TODAY.month, 1) - datetime.timedelta(days=28), 'mmm yyyy')]:
        label(st, 'A%d' % r, k); cell(st, 'B%d' % r, v, 'in', f, align='left')
    st['B5'].value = datetime.date(st['B5'].value.year, st['B5'].value.month, 1)
    label(st, 'A7', 'Approved budget (total of lines)'); cell(st, 'B7', '=Budget!C%d' % (B2 + 1), 'calc', NUM, align='left')
    label(st, 'A8', 'Forecast at completion'); cell(st, 'B8', '=Budget!H%d' % (B2 + 1), 'calc', NUM, align='left')
    label(st, 'A9', 'Variance (+ saving, − overspend)'); cell(st, 'B9', '=B7-B8', 'calc', NUM, bold=True, align='left')
    c = cell(st, 'D3', 'Cost types', bold=True, color='FFFFFF'); c.fill = F_HEAD
    for i in range(15): cell(st, 'D%d' % (4 + i), COST_TYPES[i] if i < len(COST_TYPES) else None, 'in')
    note(st, 'A11', 'Change or add cost types in the yellow list on the right. Keep "Contingency" for the reserve so its use is visible.', 'A11:B12')
    cur = lambda t: '="%s ("&Settings!$B$4&")"' % t
    # Budget
    ws = bd
    title(ws, 'Budget vs actual by line', 'Yellow: type. Committed and actual come from the Spending sheet.')
    cell(ws, 'A3', '="Project: "&Settings!B3&"   ·   Currency: "&Settings!B4', bold=True)
    head(ws, 5, ['Budget line / work package', 'Cost type', 'Budget', 'Committed', 'Actual (paid)', 'Remaining', 'Estimate to complete', 'Forecast at completion', 'Variance', 'Variance %', 'Status'])
    for col, t in zip('CDEFGHI', ['Budget', 'Committed', 'Actual (paid)', 'Remaining', 'Estimate to complete', 'Forecast at completion', 'Variance']): ws['%s5' % col] = cur(t)
    SP = lambda c: 'Spending!$%s$%d:$%s$%d' % (c, S1, c, S2)
    grid(ws, B1, B2, 'ABCDEFGHIJK', {c: 'calc' for c in 'DEFHIJK'}, {c: NUM for c in 'CDEFGHI'} | {'J': '0.0%'},
         {'D': '=IF($A{r}="","",SUMIFS(%s,%s,$A{r},%s,"Committed"))' % (SP('F'), SP('B'), SP('G')),
          'E': '=IF($A{r}="","",SUMIFS(%s,%s,$A{r},%s,"Paid"))' % (SP('F'), SP('B'), SP('G')),
          'F': '=IF($A{r}="","",N($C{r})-$D{r}-$E{r})',
          'H': '=IF($A{r}="","",$E{r}+$D{r}+IF($G{r}="",MAX(0,$F{r}),N($G{r})))',
          'I': '=IF($A{r}="","",N($C{r})-$H{r})',
          'J': '=IF(OR($A{r}="",N($C{r})=0),"",$I{r}/$C{r})',
          'K': '=IF($A{r}="","",IF($H{r}>N($C{r}),"Over budget",IF(OR($D{r}+$E{r}>=0.9*N($C{r}),AND($G{r}<>"",$H{r}>=0.95*N($C{r}))),"Watch","OK")))'})
    dv(ws, 'B%d:B%d' % (B1, B2), '=Settings!$D$4:$D$18')
    lines = [('Project manager (part time)', 'Labour', 4000, None), ('Web developer', 'Labour', 6500, 4800), ('Design agency', 'Services and contractors', 3000, 0),
             ('Content writing and translation', 'Services and contractors', 1500, None), ('Hosting and domain (first year)', 'Software and licences', 600, None),
             ('Website builder and plugins', 'Software and licences', 450, None), ('Photography', 'Services and contractors', 800, None),
             ('Testing devices', 'Equipment', 350, None), ('Staff training', 'Training', 300, None), ('Contingency (10%)', 'Contingency', 1750, None)]
    for i, (a, b, c_, g) in enumerate(lines, B1):
        ws['A%d' % i], ws['B%d' % i], ws['C%d' % i] = a, b, c_
        if g is not None: ws['G%d' % i] = g
    T = B2 + 1
    cell(ws, 'A%d' % T, 'Total', 'soft', bold=True); cell(ws, 'B%d' % T, None, 'soft')
    for col in 'CDEFGHI': cell(ws, '%s%d' % (col, T), '=SUM(%s%d:%s%d)' % (col, B1, col, B2), 'soft', NUM, bold=True)
    cell(ws, 'J%d' % T, '=IF(C{t}=0,"",I{t}/C{t})'.format(t=T), 'soft', '0.0%', bold=True)
    cell(ws, 'K%d' % T, '=IF(H{t}>C{t},"Over budget",IF(D{t}+E{t}>=0.9*C{t},"Watch","OK"))'.format(t=T), 'soft', bold=True)
    K = 'K%d:K%d' % (B1, T)
    rule(ws, K, '$K6="Over budget"', fill='FEE2E2', font=Font(bold=True, color='B91C1C'))
    rule(ws, K, '$K6="Watch"', fill='FEF3C7', font=Font(bold=True, color='92400E'))
    rule(ws, K, '$K6="OK"', fill='DCFCE7', font=Font(bold=True, color='166534'))
    ws.freeze_panes = 'B6'
    ws.page_setup.orientation = 'landscape'
    ws.print_area = 'A1:K%d' % T
    protect(ws)
    # Spending
    ws = sp
    title(ws, 'Spending log', 'One row per purchase order or bill. Change Committed to Paid when you pay it; do not add a second row.')
    head(ws, 5, ['Date', 'Budget line', 'Description', 'Supplier / person', 'Reference', 'Amount', 'Status', 'Month'])
    ws['F5'] = cur('Amount')
    grid(ws, S1, S2, 'ABCDEFGH', {'H': 'calc'}, {'A': GD, 'F': NUM, 'H': 'mmm yyyy'}, {'H': '=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'})
    dv(ws, 'B%d:B%d' % (S1, S2), '=Budget!$A$%d:$A$%d' % (B1, B2), 'Pick the budget line this cost belongs to')
    dv(ws, 'G%d:G%d' % (S1, S2), '"Committed,Paid"', 'Committed = ordered, not yet paid. Paid = money gone.')
    m0 = st['B5'].value
    M = lambda k, day: datetime.date(m0.year + (m0.month - 1 + k) // 12, (m0.month - 1 + k) % 12 + 1, day)
    for i, s in enumerate([(M(0, 5), 'Project manager (part time)', 'PM time, month 1', 'Priya Nair', 'TS-01', 1000, 'Paid'),
                           (M(0, 8), 'Hosting and domain (first year)', 'Hosting plan, 12 months', 'HostCo', 'INV-5521', 540, 'Paid'),
                           (M(0, 12), 'Design agency', 'Design: 50% advance', 'Pixel Studio', 'PO-104', 1500, 'Paid'),
                           (M(0, 12), 'Design agency', 'Design: balance on approval', 'Pixel Studio', 'PO-104', 1700, 'Committed'),
                           (M(0, 20), 'Web developer', 'Development, sprint 1', 'Vikram Rao', 'INV-031', 1600, 'Paid'),
                           (M(1, 3), 'Website builder and plugins', 'Theme and form plugin licences', 'WebKit Ltd', 'INV-889', 470, 'Paid'),
                           (M(1, 5), 'Project manager (part time)', 'PM time, month 2', 'Priya Nair', 'TS-02', 1000, 'Paid'),
                           (M(1, 10), 'Photography', 'Office and team photo shoot', 'Lens & Light', 'PO-107', 800, 'Committed'),
                           (M(1, 18), 'Contingency (10%)', 'Extra design round (risk R-06)', 'Pixel Studio', 'PO-110', 400, 'Committed')], S1):
        for j, v in enumerate(s): ws.cell(i, j + 1, v)
    rule(ws, 'A%d:H%d' % (S1, S2), '$G6="Committed"', font=Font(italic=True, color='92400E'))
    ws.freeze_panes = 'A6'; ws.auto_filter.ref = 'A5:H%d' % S2
    # Monthly
    ws = mo
    title(ws, 'Spending by month', 'Money paid each month and the running total against the budget.')
    head(ws, 5, ['Month', 'Paid', 'Committed (ordered)', 'Paid to date', '% of budget'])
    for col, t in zip('BCD', ['Paid', 'Committed (ordered)', 'Paid to date']): ws['%s5' % col] = cur(t)
    for k in range(18):
        r = 6 + k
        cell(ws, 'A%d' % r, '=Settings!$B$5' if k == 0 else '=EDATE(A%d,1)' % (r - 1), 'calc', 'mmm yyyy')
        cell(ws, 'B%d' % r, '=SUMIFS(%s,%s,$A%d,%s,"Paid")' % (SP('F'), SP('H'), r, SP('G')), 'calc', NUM)
        cell(ws, 'C%d' % r, '=SUMIFS(%s,%s,$A%d,%s,"Committed")' % (SP('F'), SP('H'), r, SP('G')), 'calc', NUM)
        cell(ws, 'D%d' % r, '=B%d' % r if k == 0 else '=D%d+B%d' % (r - 1, r), 'calc', NUM)
        cell(ws, 'E%d' % r, '=IF(Settings!$B$7=0,"",D%d/Settings!$B$7)' % r, 'calc', '0.0%')
    ch = BarChart(); ch.style = 10; ch.title = 'Paid each month'; ch.legend = None
    ch.add_data(Reference(ws, min_col=2, min_row=5, max_row=23), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=23))
    ch.series[0].graphicalProperties.solidFill = PURPLE
    ch.height, ch.width = 7, 16
    ws.add_chart(ch, 'G5')
    protect(ws)
    # Summary
    ws = sm
    title(ws, 'Summary by cost type', 'Budget, spending and forecast per type of cost.')
    head(ws, 4, ['Cost type', 'Budget', 'Committed', 'Actual (paid)', 'Forecast', 'Variance', 'Status'])
    for col, t in zip('BCDEF', ['Budget', 'Committed', 'Actual (paid)', 'Forecast', 'Variance']): ws['%s4' % col] = cur(t)
    BD = lambda c: 'Budget!$%s$%d:$%s$%d' % (c, B1, c, B2)
    for i in range(15):
        r = 5 + i
        cell(ws, 'A%d' % r, '=IF(Settings!$D${s}="","",Settings!$D${s})'.format(s=4 + i), 'box')
        for col, src in zip('BCDE', 'CDEH'):
            cell(ws, '%s%d' % (col, r), '=IF($A{r}="","",SUMIF({t},$A{r},{v}))'.format(r=r, t=BD('B'), v=BD(src)), 'calc', NUM)
        cell(ws, 'F%d' % r, '=IF($A{r}="","",B{r}-E{r})'.format(r=r), 'calc', NUM)
        cell(ws, 'G%d' % r, '=IF(OR($A{r}="",N(B{r})+N(E{r})=0),"",IF(E{r}>B{r},"Over",IF(C{r}+D{r}>=0.9*B{r},"Watch","OK")))'.format(r=r), 'calc', align='center')
    cell(ws, 'A20', 'Total', 'soft', bold=True)
    for col in 'BCDEF': cell(ws, '%s20' % col, '=SUM(%s5:%s19)' % (col, col), 'soft', NUM, bold=True)
    rule(ws, 'G5:G19', '$G5="Over"', fill='FEE2E2', font=Font(bold=True, color='B91C1C'))
    rule(ws, 'G5:G19', '$G5="Watch"', fill='FEF3C7', font=Font(bold=True, color='92400E'))
    label(ws, 'A22', 'Contingency used')
    cell(ws, 'B22', '=IFERROR((SUMIF(%s,"Contingency",%s)+SUMIF(%s,"Contingency",%s))/SUMIF(%s,"Contingency",%s),"")' % (BD('B'), BD('D'), BD('B'), BD('E'), BD('B'), BD('C')), 'calc', '0.0%')
    note(ws, 'C22', 'Share of the contingency line already committed or paid.', 'C22:F22')
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Budget vs forecast'
    ch.add_data(Reference(ws, min_col=2, min_row=4, max_row=13), titles_from_data=True)
    ch.add_data(Reference(ws, min_col=5, min_row=4, max_row=13), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=13))
    ch.series[0].graphicalProperties.solidFill = 'C4B5FD'; ch.series[1].graphicalProperties.solidFill = PURPLE
    ch.height, ch.width = 9, 16
    ws.add_chart(ch, 'A25')
    protect(ws)
    save(wb, 'PM-Project-Budget-vs-Actual.xlsx')

# ------------------------------------------------------------------ Word documents
def heading(d, text):
    para(d, text, bold=True, size=13, color=PURPLE, space=4)

def bullets(d, items, style='List Bullet'):
    for s in items: d.add_paragraph(s, style=style)

def hint(d, text):
    para(d, text, italic=True, size=9, color=MUTED, space=6)

def charter():
    d = doc_new('pm-project-charter-word')
    para(d, 'PROJECT CHARTER', bold=True, size=20, color=PURPLE, space=0)
    para(d, '[Project name]', bold=True, size=14, space=8)
    kv(d, [('Project name', '[ ]'), ('Version / date', '[1.0] · [DD-MMM-YYYY]'), ('Sponsor', '[Name, role]'), ('Project manager', '[Name, role]'),
           ('Start date', '[DD-MMM-YYYY]'), ('Target finish', '[DD-MMM-YYYY]'), ('Department / client', '[ ]'), ('Budget', '[Currency] [ ]')])
    heading(d, '1. Background and purpose')
    hint(d, 'Two or three sentences: what is wrong or missing today, and what will be better when the project is done.')
    para(d, '[Example: Our website is eight years old, does not work well on phones and brings in about 40 enquiries a month. A new, mobile-friendly website will make it easier for customers to find us and contact us.]', space=8)
    heading(d, '2. Objectives')
    hint(d, 'Three to five objectives. Each needs a measure and a target so you can check it at the end.')
    table(d, ['#', 'Objective', 'Measure', 'Target and date'],
          [['1', '[Example: More enquiries from the website]', '[Enquiries per month]', '[From 40 to 70 by DD-MMM-YYYY]'],
           ['2', '[Example: Site works well on phones]', '[Mobile speed score]', '[90 or more at launch]'], ['3', '[ ]', '[ ]', '[ ]']], [1, 6.5, 4, 5.5], fill_rows=1)
    heading(d, '3. Scope')
    t = d.add_table(rows=2, cols=2); t.style = 'Table Grid'
    for i, h_ in enumerate(['In scope (we will do)', 'Out of scope (we will not do)']):
        c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(h_); r.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); shade(c, PURPLE)
    t.rows[1].cells[0].text = '• [Example: Design and build of 25 pages]\n• [Content for the main pages]\n• [Contact form and analytics]\n• [ ]'
    t.rows[1].cells[1].text = '• [Example: Online shop or payments]\n• [Mobile app]\n• [Translation into other languages]\n• [ ]'
    para(d, space=4)
    heading(d, '4. Deliverables')
    hint(d, 'The things that will exist at the end, which someone can see, test or sign for.')
    table(d, ['Deliverable', 'Description', 'Accepted by'], [['[Example: Live website]', '[25 pages on the new design, on our domain]', '[Sponsor]'], ['[ ]', '[ ]', '[ ]']], [5, 8, 4], fill_rows=1)
    heading(d, '5. Key milestones')
    table(d, ['Milestone', 'Target date'], [['Charter approved', '[DD-MMM-YYYY]'], ['[Design approved]', '[DD-MMM-YYYY]'], ['[Testing complete]', '[DD-MMM-YYYY]'],
                                           ['[Go-live]', '[DD-MMM-YYYY]'], ['Project closed (lessons learned done)', '[DD-MMM-YYYY]']], [11, 6])
    heading(d, '6. Budget')
    table(d, ['Item', 'Amount ([currency])', 'Notes'], [['[Labour / internal time]', '[ ]', ''], ['[Suppliers and services]', '[ ]', ''], ['[Software, equipment, other]', '[ ]', ''],
                                                        ['Contingency ([10]%)', '[ ]', 'Released by the sponsor only'], ['Total', '[ ]', '']], [7, 4, 6])
    para(d, 'Funding source: [department budget / client / grant]. Spending limit the project manager can approve alone: [currency] [ ].', size=10)
    heading(d, '7. Team and roles')
    table(d, ['Role', 'Name', 'Responsibility', 'Time on project'],
          [['Sponsor', '[ ]', 'Provides money and authority; approves scope, changes and go-live', '[ ]'],
           ['Project manager', '[ ]', 'Plans, coordinates and reports; manages risks and issues', '[ ]'],
           ['[Team member]', '[ ]', '[ ]', '[ ]']], [3.5, 3.5, 7.5, 2.5], fill_rows=2)
    hint(d, 'The RACI matrix in this toolkit shows who does, owns, is consulted on and is told about each deliverable.')
    heading(d, '8. Stakeholders')
    table(d, ['Stakeholder', 'Interest in the project', 'How we keep them involved'], [['[Example: Sales team]', '[More enquiries]', '[Monthly demo, status report]'], ['[ ]', '[ ]', '[ ]']], [4.5, 6, 6.5])
    heading(d, '9. Top risks')
    table(d, ['Risk', 'Probability (1-5)', 'Impact (1-5)', 'Response'],
          [['[Example: Content from departments arrives late]', '[4]', '[3]', '[Weekly deadlines; page templates with word limits]'], ['[ ]', '', '', '']], [6.5, 2.5, 2.5, 5.5], fill_rows=2)
    heading(d, '10. Assumptions and constraints')
    bullets(d, ['Assumption: [Example: Each department names one person to provide content.]', 'Assumption: [ ]',
                'Constraint: [Example: Go-live must be before the trade fair on DD-MMM-YYYY.]', 'Constraint: [ ]'])
    heading(d, '11. How the project will be run')
    bullets(d, ['Status report: [weekly, every Friday] to the sponsor and stakeholders.', 'Team meeting: [every Monday, 30 minutes].',
                'Changes to scope, budget or dates: by change request, approved by [the sponsor / change board].',
                'Escalation: the project manager raises issues the team cannot solve to the sponsor within [two] working days.'])
    heading(d, '12. Approval')
    para(d, 'By signing, the sponsor authorises the project to start on the terms above and the project manager accepts responsibility for delivering it.', size=10, space=8)
    table(d, ['Role', 'Name', 'Signature', 'Date'], [['Sponsor', '', '', ''], ['Project manager', '', '', ''], ['[Other approver]', '', '', '']], [4, 5, 5, 3])
    dsave(d, 'PM-Project-Charter.docx')

def status_doc():
    d = doc_new('pm-weekly-status-report-word')
    para(d, 'WEEKLY PROJECT STATUS REPORT', bold=True, size=18, color=PURPLE, space=2)
    para(d, '[Project name]', bold=True, size=13, space=8)
    kv(d, [('Report no.', '[12]'), ('Period', '[DD-MMM] to [DD-MMM-YYYY]'), ('Project manager', '[ ]'), ('Sponsor', '[ ]'),
           ('Overall status', '[GREEN / AMBER / RED]'), ('Target go-live', '[DD-MMM-YYYY] (forecast [DD-MMM-YYYY])')])
    heading(d, 'Summary')
    hint(d, 'Three lines: progress, the main problem, and what you need from the reader.')
    para(d, '[Example: Design is 90% done and the build has started. Approval of the design is two days late because the sponsor is travelling. Decision needed: who can approve the design by Wednesday?]', space=8)
    heading(d, 'Status by area')
    table(d, ['Area', 'Status', 'Comment'], [['Schedule', '[G / A / R]', '[ ]'], ['Cost', '[G / A / R]', '[ ]'], ['Scope', '[G / A / R]', '[ ]'],
                                            ['Quality', '[G / A / R]', '[ ]'], ['Resources / team', '[G / A / R]', '[ ]']], [3.5, 2.5, 11])
    para(d, 'Key: GREEN = on track. AMBER = at risk; the team has a plan to recover. RED = off track; needs a decision or help from the sponsor.', size=9, italic=True, color=MUTED, space=8)
    heading(d, 'Done this period')
    bullets(d, ['[Example: Home page and contact page mockups delivered]', '[ ]', '[ ]'])
    heading(d, 'Planned for next period')
    bullets(d, ['[Example: Design sign-off meeting on Wednesday]', '[ ]', '[ ]'])
    heading(d, 'Milestones')
    table(d, ['Milestone', 'Planned', 'Forecast', 'Status'], [['[Design approved]', '[DD-MMM]', '[DD-MMM]', '[On track / At risk / Late / Done]'], ['[ ]', '', '', '']], [7, 3, 3, 4], fill_rows=2)
    heading(d, 'Top risks and issues')
    table(d, ['Risk or issue', 'Impact', 'Owner', 'Action and date'], [['[ ]', '[High / Medium / Low]', '[ ]', '[ ]'], ['[ ]', '', '', '']], [6, 2.7, 2.6, 5.7], fill_rows=1)
    heading(d, 'Budget snapshot ([currency])')
    table(d, ['Budget', 'Spent to date', 'Forecast at completion', 'Variance'], [['[ ]', '[ ]', '[ ]', '[ ]']], [4.25, 4.25, 4.25, 4.25])
    heading(d, 'Decisions or help needed')
    table(d, ['Decision or help needed', 'From whom', 'Needed by'], [['[ ]', '[ ]', '[DD-MMM-YYYY]']], [10, 4, 3], fill_rows=1)
    dsave(d, 'PM-Weekly-Status-Report.docx')

def closure():
    d = doc_new('pm-lessons-learned-word')
    para(d, 'PROJECT CLOSURE REPORT AND LESSONS LEARNED', bold=True, size=17, color=PURPLE, space=2)
    para(d, '[Project name]', bold=True, size=13, space=8)
    kv(d, [('Sponsor', '[ ]'), ('Project manager', '[ ]'), ('Planned finish', '[DD-MMM-YYYY]'), ('Actual finish', '[DD-MMM-YYYY]'),
           ('Report date', '[DD-MMM-YYYY]'), ('Prepared by', '[ ]')])
    heading(d, '1. Summary')
    hint(d, 'Three or four sentences: what the project delivered, whether it met its objectives, and the most important lesson.')
    para(d, '[ ]', space=8)
    heading(d, '2. Objectives achieved')
    table(d, ['Objective (from the charter)', 'Target', 'Result', 'Met?'], [['[ ]', '[ ]', '[ ]', '[Yes / Partly / No]'], ['[ ]', '', '', '']], [6.5, 3.5, 4.5, 2.5], fill_rows=1)
    heading(d, '3. Schedule and cost')
    table(d, ['', 'Baseline (plan)', 'Actual', 'Difference', 'Main reason'], [['Finish date', '[ ]', '[ ]', '[ ] days', '[ ]'], ['Total cost ([currency])', '[ ]', '[ ]', '[ ]', '[ ]'],
                                                                             ['Approved changes', '—', '[number]', '[cost, days]', '[ ]']], [3.5, 3, 3, 2.5, 5])
    heading(d, '4. Handover and closure checklist')
    table(d, ['✓', 'Item', 'Owner / notes'],
          [['☐', 'Deliverables accepted and signed off by the sponsor', ''], ['☐', 'Users trained; user guide handed over', ''], ['☐', 'Support and maintenance owner agreed', ''],
           ['☐', 'Open issues and defects handed over with owners', ''], ['☐', 'Supplier contracts and purchase orders closed', ''], ['☐', 'Final invoices paid; project budget closed', ''],
           ['☐', 'Access, accounts and passwords transferred or removed', ''], ['☐', 'Project files archived in [location]', ''], ['☐', 'Team members released and thanked', '']], [1, 10, 6])
    heading(d, '5. Lessons learned')
    hint(d, 'One row per lesson. Write a recommendation someone can act on next time, not a complaint.')
    table(d, ['Area', 'What went well', 'What did not go well', 'Effect', 'Recommendation'],
          [['Planning', '[ ]', '[ ]', '[ ]', '[ ]'], ['Scope and changes', '', '', '', ''], ['People and roles', '', '', '', ''], ['Suppliers', '', '', '', ''],
           ['Communication', '', '', '', ''], ['Technology', '', '', '', ''], ['Quality and testing', '', '', '', ''], ['Budget', '', '', '', '']], [2.8, 3.4, 3.4, 2.6, 4.8])
    heading(d, '6. Top three recommendations for the next project')
    bullets(d, ['[ ]', '[ ]', '[ ]'], 'List Number')
    heading(d, '7. Sign-off')
    table(d, ['Role', 'Name', 'Signature', 'Date'], [['Sponsor', '', '', ''], ['Project manager', '', '', '']], [4, 5, 5, 3])
    d.add_page_break()
    heading(d, 'Guide: a 60-minute lessons learned meeting')
    hint(d, 'Delete this page before you send the report.')
    table(d, ['Minutes', 'Step', 'How'],
          [['0–5', 'Set the rules', 'We look for lessons, not blame. Everything said stays in the room unless agreed.'],
           ['5–15', 'Look back at the timeline', 'Walk through the plan and the main events: milestones, changes, problems.'],
           ['15–30', 'What went well', 'Everyone writes notes silently, then reads them out. Group similar notes.'],
           ['30–45', 'What did not go well', 'The same way. Ask "why?" until you reach a cause the team could change.'],
           ['45–55', 'Recommendations', 'For each group of notes agree one recommendation with an owner.'],
           ['55–60', 'Close', 'Agree the top three recommendations and who will share them.']], [2.2, 4.3, 10.5])
    para(d, 'Questions that help', bold=True, color=PURPLE)
    bullets(d, ['If we started this project again tomorrow, what would we do differently?', 'What surprised us? Could we have seen it coming?',
                'Which decision took too long, and why?', 'What should the next project manager be told on day one?'])
    dsave(d, 'PM-Lessons-Learned-Closure-Report.docx')

MAKERS = [gantt, raci, raid, budget, charter, status_doc, closure]
