"""File properties of every download (owner rule): Author / Last modified by / Company = Nivesh Sahayak, and the website
   https://www.niveshsahayak.in/ as the Hyperlink base (File › Info › Properties in Excel and Word) and in the comments.
   Excel writes the PC's Office user name into "last modified by" when check_office.ps1 saves a workbook, and python-docx
   leaves Company empty, so this runs at the end of check_office.ps1 (or by hand):
   python -X utf8 scripts/fix_meta.py [file names]   (no names = every .xlsx and .docx in files/)"""
import os, re, sys, zipfile, shutil, tempfile
DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'files')
NAME, WEB = b'Nivesh Sahayak', b'https://www.niveshsahayak.in/'

def tag(data, name, value):
    """set <name>value</name> in an Office properties part, adding it when missing"""
    new = b'<%s>%s</%s>' % (name, value, name)
    data, n = re.subn(rb'<%s>.*?</%s>|<%s/>' % (name, name, name), new, data, flags=re.S)
    return data if n else data.replace(b'</Properties>', new + b'</Properties>')

def fix(p):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(p)[1]).name
    with zipfile.ZipFile(p) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == 'docProps/core.xml':
                data = re.sub(rb'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', b'<cp:lastModifiedBy>' + NAME + b'</cp:lastModifiedBy>', data)
                data = re.sub(rb'<dc:creator>.*?</dc:creator>', b'<dc:creator>' + NAME + b'</dc:creator>', data)
            elif it.filename == 'docProps/app.xml':
                data = tag(data, b'Company', NAME)
                data = tag(data, b'HyperlinkBase', WEB)
                data = re.sub(rb'<Manager>.*?</Manager>|<Manager/>', b'', data, flags=re.S)
            zo.writestr(it, data)
    shutil.move(tmp, p)

for fn in sorted(sys.argv[1:] or os.listdir(DIR)):   # file names given = only those
    if fn.endswith(('.xlsx', '.docx')): fix(os.path.join(DIR, fn))
print('file properties set to Nivesh Sahayak + website')
