"""Builds toolkit.groupchaton.com into _site/ (GitHub Actions runs this on every push; also fine to run locally).
   python -X utf8 scripts/build.py            → _site/
   Reads content/catalog.py (all text), files/ (the Excel/Word downloads, already checked in Excel by check_office.ps1),
   static/ (brand pictures), assets/ (CSS/JS). Writes every page with its own title, description, canonical, Open Graph,
   one H1 and JSON-LD; plus search data, sitemap.xml, robots.txt, ads.txt, manifest, service worker, 404 and CNAME."""
import os, sys, json, html, shutil, hashlib, datetime, re
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, 'content'))
from catalog import TOOLKITS, T, KIT
from openpyxl import load_workbook
from docx import Document

OUT = os.environ.get('TK_OUT') or os.path.join(HERE, '_site')   # TK_OUT = another folder for a private preview
SITE = 'https://toolkit.groupchaton.com'
NAME = 'Toolkit by GroupChatOn'
TAG = 'Office templates and add-ins for Excel and Word'
TODAY = datetime.date.today().isoformat()
YEAR = datetime.date.today().year
AD_CLIENT = 'ca-pub-5935207738848708'
esc = lambda s: html.escape(str(s or ''), quote=True)

def wr(rel, text, binary=False):
    p = os.path.join(OUT, rel.replace('/', os.sep))
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'wb' if binary else 'w', **({} if binary else {'encoding': 'utf-8', 'newline': '\n'})) as f: f.write(text)

def ver(rel):   # cache-busting stamp from the file's contents
    with open(os.path.join(HERE, rel.replace('/', os.sep)), 'rb') as f: return hashlib.md5(f.read()).hexdigest()[:8]

def url(t): return '/%s/%s/' % (t['kit'], t['id'])
KIND = {'excel': ('Excel template', 'xlsx'), 'word': ('Word template', 'docx'), 'online': ('Online tool', '')}

# ------------------------------------------------------------------ page shell
NAV = ''.join("<a href='/%s/'>%s</a>" % (k['id'], esc(k['name'])) for k in TOOLKITS)
# footer: the toolkits in three short groups instead of one long column (owner: "footer should be organised")
FOOT_GROUPS = [('Accounts & business', ['accounting-gst', 'shop', 'sales', 'purchase', 'restaurant', 'rentals']),
               ('Work & services', ['construction', 'clinic', 'events', 'freelancers', 'projects', 'teachers']),
               ('Office & personal', ['office-hr', 'general-office', 'letters', 'contracts', 'home-money'])]
_left = [k['id'] for k in TOOLKITS if not any(k['id'] in g for _, g in FOOT_GROUPS)]
if _left: FOOT_GROUPS[-1][1].extend(_left)   # a new toolkit lands in the last group until it is placed
FOOT_KITS = ''.join("<nav class='ftr-nav' aria-label='%s'><b>%s</b>%s</nav>" % (esc(t), esc(t), ''.join("<a href='/%s/'>%s</a>" % (i, esc(KIT[i]['name'])) for i in ids if i in KIT))
                    for t, ids in FOOT_GROUPS)
CRIT = """html{background:#F6F8FC;color:#14213D;color-scheme:light}html[data-theme=dark]{background:#0B1220;color:#E7ECF7;color-scheme:dark}
body{margin:0;font-family:'Plus Jakarta Sans',system-ui,'Segoe UI',Roboto,Arial,sans-serif;overflow-x:hidden}
.tk-load{position:fixed;inset:0 0 auto 0;height:3px;z-index:99;background:linear-gradient(90deg,transparent,#2563EB,#60A5FA,transparent);background-size:50% 100%;background-repeat:no-repeat;animation:tkbar 1s linear infinite}
@keyframes tkbar{from{background-position:-50% 0}to{background-position:150% 0}}
.tk-ready .tk-load{display:none}
.tk-mount:empty,.tk-sk{min-height:320px;border-radius:16px;background:linear-gradient(100deg,rgba(127,140,170,.10) 30%,rgba(127,140,170,.22) 50%,rgba(127,140,170,.10) 70%);background-size:300% 100%;animation:tksh 1.2s ease-in-out infinite}
@keyframes tksh{from{background-position:100% 0}to{background-position:0 0}}
main{animation:tkin .35s ease-out both}@keyframes tkin{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
@media (prefers-reduced-motion:reduce){main,.tk-load,.tk-mount:empty{animation:none}}"""
THEME_JS = """(function(){var t;try{t=localStorage.getItem('tk-theme')}catch(e){}if(!t)t=matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light';document.documentElement.setAttribute('data-theme',t)})();"""

def page(path, title, desc, main, *, ld=None, ads=True, noindex=False, crumbs=None, og_type='website', body_cls='', tool=None, extra_head=''):
    canon = SITE + path
    graph = [{'@type': 'WebSite', '@id': SITE + '/#website', 'url': SITE + '/', 'name': NAME, 'inLanguage': 'en-IN',
              'publisher': {'@id': SITE + '/#org'},
              'potentialAction': {'@type': 'SearchAction', 'target': SITE + '/search/?q={search_term_string}', 'query-input': 'required name=search_term_string'}},
             {'@type': 'Organization', '@id': SITE + '/#org', 'name': NAME, 'url': SITE + '/', 'logo': SITE + '/brand/favicon-512.png',
              'parentOrganization': {'@type': 'Organization', 'name': 'GroupChatOn', 'url': 'https://www.groupchaton.com/'}}]
    if crumbs:
        graph.append({'@type': 'BreadcrumbList', 'itemListElement': [
            {'@type': 'ListItem', 'position': i + 1, 'name': n, 'item': SITE + u} for i, (n, u) in enumerate([('Home', '/')] + crumbs)]})
    graph += ld or []
    for g in graph:
        if g.get('@type') in ('WebPage', 'CollectionPage', 'AboutPage', 'ContactPage') and 'url' not in g: g.update(url=canon, name=title, description=desc)
    ldj = json.dumps({'@context': 'https://schema.org', '@graph': graph}, ensure_ascii=False).replace('</', '<\\/')
    bc = ''
    if crumbs:
        bc = "<nav class='crumbs' aria-label='Breadcrumb'><a href='/'>Home</a>" + ''.join(
            " <span aria-hidden='true'>›</span> " + ("<a href='%s'>%s</a>" % (u, esc(n)) if i < len(crumbs) - 1 else "<span aria-current='page'>%s</span>" % esc(n))
            for i, (n, u) in enumerate(crumbs)) + '</nav>'
    adsense = ("<script async src='https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=%s' crossorigin='anonymous'></script>" % AD_CLIENT) if ads else ''
    tool_js = "<script defer src='/assets/tools.js?v=%s'></script>" % ver('assets/tools.js') if tool else ''
    if tool and os.path.exists(os.path.join(HERE, 'assets', 'tools', tool + '.js')):   # a tool in its own file registers itself with TKTools
        tool_js += "<script defer src='/assets/tools/%s.js?v=%s'></script>" % (tool, ver('assets/tools/%s.js' % tool))
    return f"""<!doctype html>
<html lang='en-IN' data-theme='light'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1, viewport-fit=cover'>
<title>{esc(title)}</title>
<meta name='description' content='{esc(desc)}'>
<link rel='canonical' href='{canon}'>
{"<meta name='robots' content='noindex, follow'>" if noindex else "<meta name='robots' content='index, follow, max-image-preview:large'>"}
<meta property='og:type' content='{og_type}'><meta property='og:site_name' content='{NAME}'><meta property='og:locale' content='en_IN'>
<meta property='og:title' content='{esc(title)}'><meta property='og:description' content='{esc(desc)}'><meta property='og:url' content='{canon}'>
<meta property='og:image' content='{SITE}/brand/share.png'><meta property='og:image:width' content='1200'><meta property='og:image:height' content='630'>
<meta name='twitter:card' content='summary_large_image'><meta name='twitter:title' content='{esc(title)}'><meta name='twitter:description' content='{esc(desc)}'><meta name='twitter:image' content='{SITE}/brand/share.png'>
<meta name='theme-color' content='#2563EB'>
<link rel='icon' href='/brand/favicon.ico' sizes='any'><link rel='icon' type='image/png' sizes='32x32' href='/brand/favicon-32.png'>
<link rel='apple-touch-icon' href='/brand/apple-touch-icon.png'><link rel='manifest' href='/manifest.webmanifest'>
<style>{CRIT}</style>
<script>{THEME_JS}</script>
<link rel='preconnect' href='https://fonts.googleapis.com'><link rel='preconnect' href='https://fonts.gstatic.com' crossorigin>
<link rel='stylesheet' href='https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Sora:wght@600;700;800&display=swap'>
<link rel='stylesheet' href='/assets/toolkit.css?v={ver('assets/toolkit.css')}'>
<script type='application/ld+json'>{ldj}</script>
{extra_head}{adsense}
</head>
<body class='{body_cls}' data-ads='{1 if ads else 0}'>
<div class='tk-load' aria-hidden='true'></div>
<a class='skip' href='#main'>Skip to content</a>
<header class='hdr'>
  <div class='wrap hdr-in'>
    <a class='brand' href='/' aria-label='{NAME}: home'><img class='logo-l' src='/brand/logo-light.png' width='218' height='50' alt='GroupChatOn'><img class='logo-d' src='/brand/logo-dark.png' width='218' height='50' alt='GroupChatOn'><img class='logo-m' src='/brand/mark.svg' width='32' height='32' alt=''><span class='brand-t'>Toolkit</span></a>
    <nav class='nav' id='nav' aria-label='Toolkits'>{NAV}</nav>
    <div class='hdr-act'>
      <a class='ibtn' href='/search/' aria-label='Search' title='Search (/)'><svg viewBox='0 0 24 24' aria-hidden='true'><circle cx='11' cy='11' r='7'/><path d='m20 20-3.5-3.5'/></svg></a>
      <button class='ibtn' type='button' data-theme-toggle aria-label='Switch light or dark mode' title='Light / dark'><svg viewBox='0 0 24 24' aria-hidden='true'><path d='M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z'/></svg></button>
      <button class='ibtn' type='button' data-install hidden aria-label='Install the Toolkit app' title='Install app'><svg viewBox='0 0 24 24' aria-hidden='true'><path d='M12 3v12M7 10l5 5 5-5M5 21h14'/></svg></button>
      <button class='ibtn menu-btn' type='button' aria-controls='nav' aria-expanded='false' aria-label='Menu'><svg viewBox='0 0 24 24' aria-hidden='true'><path d='M4 7h16M4 12h16M4 17h16'/></svg></button>
    </div>
  </div>
</header>
<main id='main' class='wrap'>
{bc}
{main}
</main>
<footer class='ftr'>
  <div class='wrap ftr-in'>
    <div class='ftr-brand'><img src='/brand/mark.svg' width='40' height='40' alt=''><div><b>{NAME}</b><br><span>{TAG}. Free, private, works offline.</span></div></div>
    {FOOT_KITS}
    <nav class='ftr-nav' aria-label='Site'><b>Site</b><a href='/p/about.html'>About</a><a href='/p/contact.html'>Contact</a><a href='/my-details/'>My saved details</a><a href='/p/privacy-policy.html'>Privacy policy</a><a href='/p/disclaimer.html'>Disclaimer</a><a href='/p/terms-of-use.html'>Terms of use</a><a href='/search/'>Search</a></nav>
    <nav class='ftr-nav' aria-label='GroupChatOn'><b>GroupChatOn</b><a href='https://infinity.groupchaton.com/'>Infinity free tools ↗</a><a href='https://www.groupchaton.com/'>Novels ↗</a><a href='https://share.groupchaton.com/'>Share ↗</a><a href='https://help.groupchaton.com/'>Help ↗</a></nav>
  </div>
  <div class='wrap ftr-base'>
    <p>© {YEAR} GroupChatOn. All rights reserved. Microsoft, Excel and Word are trademarks of Microsoft Corporation; this site is not affiliated with Microsoft.</p>
    <p><span data-ns-credit>Designed &amp; Developed by: <a href='https://www.niveshsahayak.in/'>Nivesh Sahayak</a></span></p>
  </div>
</footer>
<script defer src='/assets/toolkit.js?v={ver('assets/toolkit.js')}'></script>
{tool_js}
<script async src='https://assets.niveshsahayak.in/assets/credit.js'></script>
</body>
</html>
"""

def ad(kind):   # filled by toolkit.js (never next to buttons, never on policy pages)
    return "<div class='ad-slot' data-ad='%s'></div>" % kind

def promo():
    return "<div class='promo-slot'></div>"

# ------------------------------------------------------------------ previews of the files
PREVIEW = {'GST-Invoice-Template.xlsx': ('Invoice', 'A1:J29'), 'GST-Monthly-Summary.xlsx': ('Summary', 'A3:F28'),
           'Client-Dues-Ageing.xlsx': ('Ageing', 'A3:G7'), 'Bank-Reconciliation-Statement.xlsx': ('Statement', 'A4:C16'),
           'Stock-Register.xlsx': ('Stock', 'A4:J9'), 'Daily-Cash-Book.xlsx': ('Entries', 'A3:F8'),
           'Student-Attendance-Register.xlsx': ('Register', 'A7:P13'), 'Marks-and-Report-Card.xlsx': ('Marks', 'A3:N11'),
           'Tuition-Fee-Register.xlsx': ('Fee Grid', 'A3:R6'), 'Salary-Sheet-and-Payslip.xlsx': ('Payslip', 'A1:E21'),
           'Employee-Leave-Tracker.xlsx': ('Balance', 'A3:I7'), 'Task-Tracker.xlsx': ('Tasks', 'A3:J9'), 'Expense-Claim-Form.xlsx': ('Claim', 'A1:F37')}
for _t in T: PREVIEW.update(_t.get('preview') or _t.get('previews') or {})   # new toolkits name their preview sheet in content/kits/<id>.py

def indian(n, dec):
    neg, n = n < 0, abs(n)
    s = ('%.' + str(dec) + 'f') % n
    ip, _, fp = s.partition('.')
    if len(ip) > 3:
        head, tail = ip[:-3], ip[-3:]
        head = ','.join([head[max(0, i - 2):i] for i in range(len(head), 0, -2)][::-1])
        ip = head + ',' + tail
    return ('-' if neg else '') + ip + ('.' + fp if fp else '')

def fmt_val(c):
    v, f = c.value, c.number_format or ''
    if v is None: return ''
    if isinstance(v, (datetime.datetime, datetime.date)):
        if 'mmmm' in f: return v.strftime('%B %Y')
        if 'mmm' in f: return v.strftime('%b-%y')
        if f.strip() == 'd': return str(v.day)
        return v.strftime('%d-%m-%Y')
    if isinstance(v, bool): return 'Yes' if v else 'No'
    if isinstance(v, (int, float)):
        if '%' in f: return ('%.1f%%' if '0.0' in f else '%.0f%%') % (v * 100)
        if '₹' in f: return '₹ ' + indian(v, 2 if '0.00' in f else 0)
        if '#,##0;;' in f: return indian(v, 0) if v else '–'
        return indian(v, 0) if float(v).is_integer() else ('%.2f' % v).rstrip('0').rstrip('.')
    return str(v)

def xlsx_preview(fn):
    sh, rng = PREVIEW[fn]
    ws = load_workbook(os.path.join(HERE, 'files', fn), data_only=True)[sh]
    merged = {}
    skip = set()
    for m in ws.merged_cells.ranges:
        merged[(m.min_row, m.min_col)] = (m.max_row - m.min_row + 1, m.max_col - m.min_col + 1)
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                if (r, c) != (m.min_row, m.min_col): skip.add((r, c))
    rows = []
    cells = ws[rng]
    last_col = cells[0][-1].column
    for row in cells:
        if all(c.value in (None, '') for c in row) and not any((c.row, c.column) in merged for c in row): continue
        tds = []
        for c in row:
            if (c.row, c.column) in skip: continue
            rs, cs = merged.get((c.row, c.column), (1, 1))
            cs = min(cs, last_col - c.column + 1)
            fill = (c.fill.fgColor.rgb or '') if c.fill and c.fill.fill_type == 'solid' else ''
            cls = {'FF2563EB': 'h', 'FFFFF7D6': 'i', 'FFF1F5F9': 'c', 'FFE8F0FE': 's'}.get(fill, '')
            if c.font and c.font.b: cls += ' b'
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool): cls += ' n'
            attrs = (" colspan='%d'" % cs if cs > 1 else '') + (" class='%s'" % cls.strip() if cls.strip() else '')
            tds.append('<td%s>%s</td>' % (attrs, esc(fmt_val(c))))
        rows.append('<tr>%s</tr>' % ''.join(tds))
    return "<div class='xl' role='region' aria-label='Preview of the %s sheet' tabindex='0'><div class='xl-tab'>%s</div><table>%s</table></div>" % (esc(sh), esc(sh), ''.join(rows))

def docx_preview(fn):
    d = Document(os.path.join(HERE, 'files', fn))
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    out, n = [], 0
    for el in d.element.body.iterchildren():
        if n > 34: break
        tag = el.tag.split('}')[1]
        if tag == 'p':
            p = Paragraph(el, d)
            t = p.text.strip()
            if not t: continue
            b = any(r.bold for r in p.runs)
            big = any(r.font.size and r.font.size.pt >= 13 for r in p.runs)
            al = ' style="text-align:center"' if p.alignment == 1 else (' style="text-align:right"' if p.alignment == 2 else '')
            li = p.style.name.startswith('List')
            out.append('<p class="%s"%s>%s%s</p>' % (('big ' if big else '') + ('b' if b else ''), al, '• ' if li else '', esc(t))); n += 1
        elif tag == 'tbl':
            tb = Table(el, d)
            trs = []
            for i, r in enumerate(tb.rows[:7]):
                seen, tds = set(), []
                for c in r.cells:
                    if id(c._tc) in seen: continue
                    seen.add(id(c._tc)); tds.append('<td>%s</td>' % esc(c.text.replace('\n', ' ').strip()))
                trs.append('<tr%s>%s</tr>' % (" class='h'" if i == 0 and tb.style and 'Grid' in tb.style.name and r.cells[0].text and r.cells[0].text in ('#', 'Component') else '', ''.join(tds)))
            out.append('<table>%s</table>' % ''.join(trs)); n += 3
    return "<div class='doc' role='region' aria-label='Preview of %s' tabindex='0'>%s<p class='more'>…</p></div>" % (esc(fn), ''.join(out))

# ------------------------------------------------------------------ building blocks
def card(t):
    k = KIT[t['kit']]
    lab = KIND[t['kind']][0]
    return ("<a class='card' href='%s' style='--cc:%s'><span class='card-ic' aria-hidden='true'>%s</span><span class='card-b'><b>%s</b>"
            "<small>%s</small><em class='kind kind-%s'>%s</em></span></a>") % (url(t), k['color'], t['icon'], esc(t['name']), esc(t['desc'][:120] + ('…' if len(t['desc']) > 120 else '')), t['kind'], lab)

def kit_card(k):
    items = [t for t in T if t['kit'] == k['id']]
    return ("<a class='kcard' href='/%s/' style='--cc:%s'><span class='kcard-ic' aria-hidden='true'>%s</span><b>%s</b><small>%s</small>"
            "<span class='kcard-n'>%d templates · %d online tool%s</span></a>") % (
        k['id'], k['color'], k['icon'], esc(k['name']), esc(k['who']), sum(1 for t in items if t['kind'] != 'online'),
        sum(1 for t in items if t['kind'] == 'online'), '' if sum(1 for t in items if t['kind'] == 'online') == 1 else 's')

def faq_html(faqs):
    return "<div class='faq'>" + ''.join("<details><summary>%s</summary><p>%s</p></details>" % (esc(q), esc(a)) for q, a in faqs) + '</div>'

def faq_ld(faqs):
    return {'@type': 'FAQPage', 'mainEntity': [{'@type': 'Question', 'name': q, 'acceptedAnswer': {'@type': 'Answer', 'text': a}} for q, a in faqs]}

def size_of(fn):
    b = os.path.getsize(os.path.join(HERE, 'files', fn))
    return '%d KB' % max(1, round(b / 1024))

# ------------------------------------------------------------------ pages
def item_page(t):
    k = KIT[t['kit']]
    kind, ext = KIND[t['kind']]
    same = [x for x in T if x['kit'] == t['kit'] and x['id'] != t['id']]
    other = [x for x in T if x['kit'] != t['kit'] and x['kind'] == t['kind']][:3]
    if t['kind'] == 'online':
        top = ("<div class='tk-tool' data-tool='%s'><div class='tk-mount'></div>"
               "<noscript><p class='note'>This tool needs JavaScript. Turn it on, or use the Excel template from this toolkit.</p></noscript></div>") % t['tool']
        files_note = ''
    else:
        btns = ''.join(("<a class='btn btn-dl' href='/files/%s' download data-dl='%s'><svg viewBox='0 0 24 24' aria-hidden='true'><path d='M12 3v12M7 10l5 5 5-5M5 21h14'/></svg>"
                        "<span>Download %s <small>.%s · %s</small></span></a>") % (fn, fn, esc(lab), fn.rsplit('.', 1)[1], size_of(fn)) for fn, lab in t['files'])
        prev = ''.join(xlsx_preview(fn) if fn.endswith('.xlsx') else docx_preview(fn) for fn, _ in t['files'])
        top = ("<div class='dl-box'><div class='dl-btns'>%s</div><p class='dl-note'>Free · no sign-up · opens in Microsoft %s, Google %s, LibreOffice and WPS Office. "
               "Files contain no macros.</p></div><h2>Preview</h2><p class='muted'>Real view of the file with its sample data (yellow cells are where you type).</p>%s") % (
            btns, 'Excel' if t['kind'] == 'excel' else 'Word', 'Sheets' if t['kind'] == 'excel' else 'Docs', prev)
    body = f"""<article class='item'>
<header class='item-hd'><span class='kind kind-{t['kind']}'>{kind}</span>
<h1>{esc(t['name'])}</h1>
<p class='lead'>{esc(t['lead'])}</p></header>
{top}
<h2>What's inside</h2>
<ul class='ticks'>{''.join('<li>%s</li>' % esc(s) for s in t['inside'])}</ul>
{ad('inArticle')}
<h2>How to use it</h2>
<ol class='steps'>{''.join('<li>%s</li>' % esc(s) for s in t['steps'])}</ol>
{promo()}
<h2>Tips</h2>
<ul class='tips'>{''.join('<li>%s</li>' % esc(s) for s in t['tips'])}</ul>
<h2>Questions and answers</h2>
{faq_html(t['faqs'])}
<section class='more'><h2>More in the {esc(k['name'])} toolkit</h2><div class='grid'>{''.join(card(x) for x in same)}</div>
{"<h2>Also useful</h2><div class='grid'>%s</div>" % ''.join(card(x) for x in other) if other else ''}
</section>
{ad('end')}
</article>"""
    ld = [{'@type': 'WebPage', 'url': SITE + url(t), 'name': t['title'], 'description': t['desc'], 'isPartOf': {'@id': SITE + '/#website'}, 'inLanguage': 'en-IN', 'dateModified': TODAY}, faq_ld(t['faqs'])]
    if t['kind'] == 'online':
        ld.append({'@type': 'WebApplication', 'name': t['name'], 'url': SITE + url(t), 'description': t['desc'], 'applicationCategory': 'BusinessApplication',
                   'operatingSystem': 'Any (web browser)', 'isAccessibleForFree': True, 'offers': {'@type': 'Offer', 'price': '0', 'priceCurrency': 'INR'},
                   'publisher': {'@id': SITE + '/#org'}})
    else:
        for fn, lab in t['files']:
            ld.append({'@type': 'DigitalDocument', 'name': lab, 'url': SITE + '/files/' + fn, 'description': t['desc'], 'isAccessibleForFree': True,
                       'encodingFormat': 'application/vnd.openxmlformats-officedocument.' + ('spreadsheetml.sheet' if fn.endswith('xlsx') else 'wordprocessingml.document'),
                       'inLanguage': 'en-IN', 'publisher': {'@id': SITE + '/#org'}, 'license': SITE + '/p/terms-of-use.html'})
    wr(url(t)[1:] + 'index.html', page(url(t), t['title'], t['desc'], body, ld=ld, crumbs=[(k['name'], '/%s/' % k['id']), (t['name'], url(t))],
                                         tool=t.get('tool') if t['kind'] == 'online' else None, body_cls='pg-item' + (' pg-tool' if t['kind'] == 'online' else '')))

def kit_page(k):
    items = [t for t in T if t['kit'] == k['id']]
    path = '/%s/' % k['id']
    blocks = ''.join("<section class='kit-item'><h3><a href='%s'>%s %s</a></h3><p>%s</p><ul>%s</ul><p><a class='go' href='%s'>%s →</a></p></section>" % (
        url(t), t['icon'], esc(t['name']), esc(t['lead']), ''.join('<li>%s</li>' % esc(s) for s in t['inside'][:3]), url(t),
        'Open the tool' if t['kind'] == 'online' else 'Preview and download') for t in items)
    body = f"""<header class='kit-hd' style='--cc:{k['color']}'><span class='kit-ic' aria-hidden='true'>{k['icon']}</span>
<h1>{esc(k['name'])} Toolkit</h1>
{''.join('<p class="lead">%s</p>' % esc(p) for p in k['intro'])}
<p class='who'><b>Made for:</b> {esc(k['who'])}</p></header>
<div class='grid'>{''.join(card(t) for t in items)}</div>
{ad('inArticle')}
<h2>What each one does</h2>
<div class='kit-items'>{blocks}</div>
{promo()}
<h2>Good to know</h2>
<ul class='tips'><li>Every template is free and contains no macros, so it opens without security warnings.</li>
<li>Yellow cells are for typing; grey cells calculate by themselves and are protected against accidental overwriting (Review › Unprotect Sheet, no password).</li>
<li>Online tools keep your details only in this browser on this device. Nothing is uploaded.</li>
<li>Tax rates and statutory limits mentioned in these files are a help, not advice: check the current rules before relying on them.</li></ul>
{ad('end')}"""
    ld = [{'@type': 'CollectionPage', 'isPartOf': {'@id': SITE + '/#website'}, 'inLanguage': 'en-IN'},
          {'@type': 'ItemList', 'itemListElement': [{'@type': 'ListItem', 'position': i + 1, 'url': SITE + url(t), 'name': t['name']} for i, t in enumerate(items)]}]
    wr(path[1:] + 'index.html', page(path, k['title'], k['desc'], body, ld=ld, crumbs=[(k['name'], path)], body_cls='pg-kit'))

HOME_FAQ = [('Are these templates really free?', 'Yes. Every Excel and Word template and every online tool on this site is free to use for your own work and your clients\' work, with no sign-up. The site is supported by ads.'),
            ('Do I need Microsoft Office?', 'No. The files are standard .xlsx and .docx files. They open in Microsoft Excel and Word, and also in Google Sheets and Docs, LibreOffice, WPS Office and Excel or Word on phones.'),
            ('Is my data safe?', 'Downloaded files stay on your computer. The online tools run inside your browser and save your details only on your own device; nothing you type is sent to us.'),
            ('Do the files contain macros?', 'No. The free templates use ordinary formulas only, so Excel does not show macro security warnings.'),
            ('Do the online tools work without internet?', 'Yes, after the first visit. Install Toolkit as an app (the download icon in the header, or your browser\'s Install option) and the online tools open even when you are offline.'),
            ('Can I ask for a template?', 'Yes. Use the Contact page and describe the register, form or report you need for your work. Popular requests are added in the next batch.')]

def home():
    online = [t for t in T if t['kind'] == 'online']
    popular = [T[i] for i in range(len(T)) if T[i]['id'] in ('gst-invoice-excel', 'salary-sheet-slip-excel', 'stock-register-excel', 'attendance-register-excel',
                                                              'bank-reconciliation-excel', 'task-tracker-excel')]
    nfiles = sum(len(t.get('files', [])) for t in T)
    body = f"""<section class='hero'>
<p class='eyebrow'>Free · private · works offline</p>
<h1>Office toolkits for the work you do every day</h1>
<p class='lead'>Ready Excel and Word templates and online tools for accountants, shopkeepers, teachers, HR and every office: GST invoices, stock registers, attendance, payslips, letters and more. Download, fill the yellow cells, done.</p>
<form class='hero-search' action='/search/' role='search'><label class='sr' for='q'>Search templates and tools</label>
<input id='q' name='q' type='search' placeholder='Search, e.g. GST invoice, salary slip, attendance' autocomplete='off'><button class='btn' type='submit'>Search</button></form>
<p class='hero-stats'><b>{len(TOOLKITS)}</b> toolkits · <b>{nfiles}</b> Excel and Word files · <b>{len(online)}</b> online tools</p>
</section>
<h2 class='sec-h'>Choose your toolkit</h2>
<div class='kgrid'>{''.join(kit_card(k) for k in TOOLKITS)}</div>
{promo()}
<h2 class='sec-h'>Online tools: fill, print, download</h2>
<p class='muted'>No download needed. Fill a form and get a print-ready document, a PDF, or a real Excel or Word file.</p>
<div class='grid'>{''.join(card(t) for t in online)}</div>
{ad('home')}
<h2 class='sec-h'>Most used templates</h2>
<div class='grid'>{''.join(card(t) for t in popular)}</div>
{promo()}
<section class='why'><h2 class='sec-h'>Why Toolkit</h2><div class='why-g'>
<div><b>Made for Indian work</b><p>Rupees in lakh and crore, amounts in words, CGST/SGST/IGST, PF and ESI, A4 printing.</p></div>
<div><b>Formulas you can trust</b><p>Every workbook is opened and recalculated in real Excel before it is published, with sample data you can check by hand.</p></div>
<div><b>Your data stays yours</b><p>No accounts, no uploads. Files live on your computer; online tools save only in your browser.</p></div>
<div><b>Simple by design</b><p>Yellow cells to type, grey cells that calculate, and a "How to use" sheet inside every file.</p></div></div></section>
{promo()}
<h2 class='sec-h'>Questions and answers</h2>
{faq_html(HOME_FAQ)}
{ad('end')}"""
    ld = [{'@type': 'WebPage', 'url': SITE + '/', 'name': NAME, 'description': '', 'isPartOf': {'@id': SITE + '/#website'}, 'inLanguage': 'en-IN'}, faq_ld(HOME_FAQ),
          {'@type': 'ItemList', 'itemListElement': [{'@type': 'ListItem', 'position': i + 1, 'url': SITE + '/%s/' % k['id'], 'name': k['name'] + ' Toolkit'} for i, k in enumerate(TOOLKITS)]}]
    title = 'Toolkit by GroupChatOn: Free Excel & Word Templates and Online Office Tools'
    desc = 'Free Excel and Word templates and online tools for accountants, shops, teachers, HR and offices: GST invoice, stock register, attendance, salary slip and more.'
    ld[0]['description'] = desc
    wr('index.html', page('/', title, desc, body, ld=ld, body_cls='pg-home'))

# ------------------------------------------------------------------ fixed pages (no ads, no sponsor cards)
def static_pages():
    upd = datetime.date.today().strftime('%d %B %Y')
    P = {}
    P['about'] = ('About Toolkit by GroupChatOn', 'Who makes Toolkit by GroupChatOn, why the templates are free, and how each Excel and Word file is checked before it is published.', 'AboutPage', f"""
<h1>About Toolkit</h1>
<p class='lead'>Toolkit by GroupChatOn is a free library of Office templates and online tools for people who do the same paperwork every day: accountants, shopkeepers, teachers, office managers and HR.</p>
<h2>What you will find here</h2>
<p>Ready Excel workbooks and Word documents grouped into toolkits by profession, and online tools that turn a short form into a print-ready document or a real Excel or Word file. Everything is written for Indian work practice: rupees in lakh and crore, GST, PF and ESI, A4 paper.</p>
<h2>How the files are made</h2>
<p>Each workbook is built with ordinary formulas (no macros), opened and recalculated in Microsoft Excel, and checked for errors before it goes live. Sample data in every file lets you check the arithmetic by hand. Every file has a "How to use" sheet or clear [bracketed] blanks.</p>
<h2>Part of GroupChatOn</h2>
<p>Toolkit is a sister site of <a href='https://infinity.groupchaton.com/'>Infinity</a> (609 free browser tools), the <a href='https://www.groupchaton.com/'>GroupChatOn novels</a> site and <a href='https://share.groupchaton.com/'>Share</a>. The sites are free to use and supported by advertising.</p>
<h2>Coming next</h2>
<p>More templates for every toolkit, an Excel add-in with a Toolkit ribbon tab, and a side panel inside Excel and Word. Ideas are welcome on the <a href='/p/contact.html'>contact page</a>.</p>""")
    P['contact'] = ('Contact Toolkit by GroupChatOn', 'Ask for a template, report a mistake in a file or send feedback to Toolkit by GroupChatOn.', 'ContactPage', f"""
<h1>Contact us</h1>
<p class='lead'>Found a mistake in a formula, need a template for your work, or have feedback? Messages for every GroupChatOn site go through one contact form on our main site.</p>
<p><a class='btn' href='https://www.groupchaton.com/p/contact-us.html?from=toolkit'>Open the contact form</a></p>
<h2>Help us answer faster</h2>
<ul><li><b>Start with "Toolkit"</b> so we know which site it is about.</li><li><b>The file or tool name</b>, for example "GST Invoice Register".</li><li><b>What went wrong</b>, or what the template you need should do.</li><li><b>Your email</b>, so we can reply.</li></ul>
<p>We usually reply within a few days.</p>""")
    P['privacy-policy'] = ('Privacy Policy | Toolkit by GroupChatOn', 'How Toolkit by GroupChatOn handles data: no accounts, online tools save only in your browser, files stay on your device; cookies and ads by Google.', 'WebPage', f"""
<h1>Privacy policy</h1>
<p class='muted'>Last updated: {upd}</p>
<p>This policy explains what information toolkit.groupchaton.com ("Toolkit", "we") collects and how it is used.</p>
<h2>What you type stays with you</h2>
<ul><li><b>Downloaded files</b> (Excel, Word) are stored and edited on your own device. We never see what you type into them.</li>
<li><b>Online tools</b> run entirely in your browser. Details you choose to keep (for example your business name) are saved in your browser's local storage on your device only. Clearing the site's data in your browser removes them.</li>
<li>We have no user accounts and no sign-up.</li></ul>
<h2>Information collected automatically</h2>
<p>Like most websites, our host (GitHub Pages) records standard technical information such as IP address, browser type and the pages requested, for security and reliability. We do not combine this with any other information.</p>
<h2>Cookies and advertising</h2>
<p>We show ads from Google AdSense. Google and its partners use cookies to show ads based on your visits to this and other websites. You can turn off personalised advertising at <a href='https://adssettings.google.com/' rel='noopener'>Google Ad Settings</a> or learn more at <a href='https://policies.google.com/technologies/ads' rel='noopener'>How Google uses cookies in advertising</a>. Visitors in some regions are asked for consent before personalised ads are shown.</p>
<p>Some cards are sponsored offers (labelled "Sponsored"). Clicking them takes you to the sponsor's own website, whose privacy policy then applies.</p>
<h2>Third-party services</h2>
<p>Pages load fonts from Google Fonts, the credit badge from niveshsahayak.in, and, when you download from an online tool, open-source libraries (ExcelJS, docx) from public code servers (cdnjs, jsDelivr). These services receive your IP address when they deliver files.</p>
<h2>Children</h2><p>The site is meant for working adults and is not directed at children under 13.</p>
<h2>Changes and contact</h2><p>If this policy changes, the date above changes. Questions: use the <a href='/p/contact.html'>contact page</a>.</p>""")
    P['disclaimer'] = ('Disclaimer | Toolkit by GroupChatOn', 'Toolkit templates and online tools are aids for everyday office work, not legal, tax or financial advice. Check rates and rules before relying on any figure.', 'WebPage', f"""
<h1>Disclaimer</h1>
<p class='muted'>Last updated: {upd}</p>
<p>The templates, online tools and articles on toolkit.groupchaton.com are provided free, for general information and to help with everyday office work.</p>
<h2>Not professional advice</h2>
<p>Nothing on this site is legal, tax, accounting, payroll or financial advice. GST rates, TDS, PF, ESI, professional tax and other rules change from time to time and depend on your situation. Before you rely on any figure, rate or format, check it against the current law and, where needed, with a qualified professional such as a chartered accountant.</p>
<h2>Your responsibility</h2>
<p>You are responsible for what you enter in a template or tool and for the documents you issue with it. Always review totals, tax amounts and wording before printing, signing or sending a document.</p>
<h2>No warranty</h2>
<p>We test every file carefully but cannot guarantee that it is free from errors or suitable for every purpose. The site and files are provided "as is". To the extent the law allows, we are not liable for any loss arising from their use.</p>
<h2>Trademarks</h2>
<p>Microsoft, Excel and Word are trademarks of Microsoft Corporation. Google Sheets and Google Docs are trademarks of Google LLC. Toolkit by GroupChatOn is independent and is not affiliated with, sponsored or endorsed by them.</p>
<h2>Advertising and sponsored links</h2>
<p>The site shows ads and some sponsored offers, which are labelled. We are not responsible for the content of advertisers' websites.</p>""")
    P['terms-of-use'] = ('Terms of Use | Toolkit by GroupChatOn', 'Terms for using Toolkit by GroupChatOn: free use of templates for your own and your clients\' work, no resale or re-publishing, and limits of liability.', 'WebPage', f"""
<h1>Terms of use</h1>
<p class='muted'>Last updated: {upd}</p>
<p>By using toolkit.groupchaton.com you agree to these terms.</p>
<h2>What you may do</h2>
<ul><li>Download, edit and use the templates for your own work and for work you do for your clients or employer, including in a business.</li>
<li>Print and share the documents you create with them (invoices, payslips, report cards and so on).</li>
<li>Use the online tools as often as you like.</li></ul>
<h2>What you may not do</h2>
<ul><li>Sell the templates, or re-publish them (changed or unchanged) on another website, app, marketplace or download collection.</li>
<li>Remove the credit from the template's About / How to use sheet or file properties when sharing the blank template itself.</li>
<li>Use the site in a way that harms it or other users, such as automated mass downloading.</li></ul>
<h2>Accuracy</h2><p>Please read the <a href='/p/disclaimer.html'>disclaimer</a>: the files are aids, not professional advice.</p>
<h2>Changes</h2><p>We may change the site, the files or these terms at any time. The date above shows the latest version.</p>
<h2>Law</h2><p>These terms are governed by the laws of India.</p>
<h2>Contact</h2><p>Questions: use the <a href='/p/contact.html'>contact page</a>.</p>""")
    for name, (title, desc, typ, body) in P.items():
        path = '/p/%s.html' % name
        crumbs = [(re.sub(r'\s*\|.*', '', title).replace(' Toolkit by GroupChatOn', '').replace('Contact Toolkit by GroupChatOn', 'Contact'), path)]
        wr(path[1:], page(path, title, desc, "<article class='prose'>%s</article>" % body, ld=[{'@type': typ, 'isPartOf': {'@id': SITE + '/#website'}}],
                          ads=False, crumbs=crumbs, body_cls='pg-static'))

def search_and_404():
    idx = [{'t': t['name'], 'd': t['desc'], 'u': url(t), 'k': KIT[t['kit']]['name'], 'i': t['icon'], 'kind': KIND[t['kind']][0],
            'w': ' '.join([t['title'], t['lead']] + t['inside'])} for t in T]
    idx += [{'t': k['name'] + ' Toolkit', 'd': k['desc'], 'u': '/%s/' % k['id'], 'k': 'Toolkit', 'i': k['icon'], 'kind': 'Toolkit', 'w': ' '.join(k['intro'])} for k in TOOLKITS]
    wr('data/index.json', json.dumps(idx, ensure_ascii=False, separators=(',', ':')))
    body = """<h1>Search templates and tools</h1>
<form class='hero-search' action='/search/' role='search' data-search><label class='sr' for='sq'>Search</label>
<input id='sq' name='q' type='search' placeholder='e.g. GST invoice, attendance, salary slip' autocomplete='off'><button class='btn' type='submit'>Search</button></form>
<p class='muted' data-search-count></p><div class='grid' data-search-results></div>"""
    wr('search/index.html', page('/search/', 'Search | Toolkit by GroupChatOn', 'Search all Toolkit by GroupChatOn Excel and Word templates and online office tools.', body, noindex=True, ads=False, body_cls='pg-search'))
    # /my-details/: the saved lists behind the suggestions in every online tool (owner: type details once, never again)
    body_md = """<h1>My saved details</h1>
<p class='lead'>Your own details, customers, suppliers and items are remembered on this device when you print or download from any Toolkit online tool. Next time they come up as suggestions while you type: pick one and its address, GSTIN, rate and other details fill in.</p>
<div data-masters><p class='muted'>Loading your saved details…</p></div>
<h2>Good to know</h2>
<ul>
<li><b>Private:</b> everything is kept in this browser only. Nothing is sent to us or to anyone else.</li>
<li><b>Nothing fills in by itself:</b> saved details only appear as suggestions. You choose when to use one.</li>
<li><b>Another phone or computer?</b> Tap <b>Download backup</b> here, then <b>Restore or import</b> on the other device.</li>
<li><b>Have a customer or item list in Excel?</b> Download the <b>Excel file</b>, paste your rows into the "Customers and suppliers" or "Items and services" sheet, save, then use <b>Restore or import</b>.</li>
<li><b>Clearing your browser data</b> also deletes these lists, so keep a backup.</li>
</ul>"""
    wr('my-details/index.html', page('/my-details/', 'My Saved Details | Toolkit by GroupChatOn', 'See, change, back up or move the business details, customers, suppliers and items that Toolkit online tools suggest while you type.',
                                     body_md, noindex=True, ads=False, crumbs=[('My saved details', '/my-details/')], body_cls='pg-md', tool='masters'))
    body404 = """<h1>Page not found</h1><p class='lead'>That address doesn't exist (any more). Try a search or pick a toolkit.</p>
<form class='hero-search' action='/search/' role='search'><label class='sr' for='nq'>Search</label><input id='nq' name='q' type='search' data-404-q placeholder='Search templates and tools'><button class='btn' type='submit'>Search</button></form>
<div class='kgrid'>%s</div>""" % ''.join(kit_card(k) for k in TOOLKITS)
    wr('404.html', page('/404.html', 'Page not found | Toolkit by GroupChatOn', 'The page you asked for was not found on Toolkit by GroupChatOn.', body404, noindex=True, ads=False, body_cls='pg-404'))

def site_files(paths):
    wr('robots.txt', 'User-agent: *\nAllow: /\nDisallow: /search/\n\nSitemap: %s/sitemap.xml\n' % SITE)
    wr('sitemap.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' +
       ''.join('  <url><loc>%s%s</loc><lastmod>%s</lastmod></url>\n' % (SITE, p, TODAY) for p in paths) + '</urlset>\n')
    wr('ads.txt', 'google.com, pub-5935207738848708, DIRECT, f08c47fec0942fa0\n')
    wr('CNAME', 'toolkit.groupchaton.com\n')
    wr('.nojekyll', '')
    wr('manifest.webmanifest', json.dumps({
        'id': '/', 'name': NAME, 'short_name': 'Toolkit', 'description': TAG + ': free templates and online tools for daily office work.',
        'start_url': '/?source=app', 'scope': '/', 'display': 'standalone', 'background_color': '#0B1220', 'theme_color': '#2563EB', 'lang': 'en-IN',
        'categories': ['business', 'productivity', 'education'],
        'icons': [{'src': '/brand/favicon-192.png', 'sizes': '192x192', 'type': 'image/png'}, {'src': '/brand/favicon-512.png', 'sizes': '512x512', 'type': 'image/png'},
                  {'src': '/brand/maskable-512.png', 'sizes': '512x512', 'type': 'image/png', 'purpose': 'maskable'}],
        'shortcuts': [{'name': t['name'], 'url': url(t)} for t in T if t['kind'] == 'online'][:4]}, ensure_ascii=False, indent=1))
    stamp = hashlib.md5(''.join(ver(p) for p in ('assets/toolkit.css', 'assets/toolkit.js', 'assets/tools.js')).encode() + TODAY.encode()).hexdigest()[:8]
    sw = open(os.path.join(HERE, 'assets', 'sw.js'), encoding='utf-8').read().replace('__VERSION__', stamp).replace('__PRECACHE__', json.dumps(
        ['/', '/offline.html', '/assets/toolkit.css?v=' + ver('assets/toolkit.css'), '/assets/toolkit.js?v=' + ver('assets/toolkit.js'),
         '/assets/tools.js?v=' + ver('assets/tools.js'), '/brand/logo-light.png', '/brand/logo-dark.png', '/brand/mark.svg', '/data/index.json'] +
        ['/assets/tools/%s.js?v=%s' % (t['tool'], ver('assets/tools/%s.js' % t['tool'])) for t in T if t['kind'] == 'online' and os.path.exists(os.path.join(HERE, 'assets', 'tools', t['tool'] + '.js'))] +
        [url(t) for t in T if t['kind'] == 'online']))
    wr('sw.js', sw)
    wr('offline.html', page('/offline.html', 'Offline | Toolkit by GroupChatOn', 'You are offline.', "<h1>You're offline</h1><p class='lead'>This page hasn't been saved on this device yet. The online tools you opened before work offline; downloaded files are already on your computer.</p><p><a class='btn' href='/'>Go to the home page</a></p>", noindex=True, ads=False))

def main():
    if os.path.isdir(OUT): shutil.rmtree(OUT)
    shutil.copytree(os.path.join(HERE, 'static'), OUT)
    shutil.copytree(os.path.join(HERE, 'files'), os.path.join(OUT, 'files'), ignore=shutil.ignore_patterns('~$*'))   # not Excel/Word lock files
    os.makedirs(os.path.join(OUT, 'assets'), exist_ok=True)
    for f in ('toolkit.css', 'toolkit.js', 'tools.js'): shutil.copy(os.path.join(HERE, 'assets', f), os.path.join(OUT, 'assets', f))
    if os.path.isdir(os.path.join(HERE, 'assets', 'tools')): shutil.copytree(os.path.join(HERE, 'assets', 'tools'), os.path.join(OUT, 'assets', 'tools'))
    home()
    for k in TOOLKITS: kit_page(k)
    for t in T: item_page(t)
    static_pages()
    search_and_404()
    paths = ['/'] + ['/%s/' % k['id'] for k in TOOLKITS] + [url(t) for t in T] + ['/p/%s.html' % n for n in ('about', 'contact', 'privacy-policy', 'disclaimer', 'terms-of-use')]
    site_files(paths)
    # every page: exactly one <h1>, unique title and description
    seen_t, seen_d = {}, {}
    for root, _, files in os.walk(OUT):
        for f in files:
            if not f.endswith('.html'): continue
            s = open(os.path.join(root, f), encoding='utf-8').read()
            n = len(re.findall(r'<h1[\s>]', s))
            assert n == 1, (root, f, n)
            ti = re.search(r'<title>(.*?)</title>', s).group(1); de = re.search(r"<meta name='description' content='(.*?)'>", s).group(1)
            assert ti not in seen_t, (ti, f); assert de not in seen_d, (de, f)
            seen_t[ti] = seen_d[de] = f
    print('built %d pages into %s' % (len(seen_t), OUT))

if __name__ == '__main__':
    main()
