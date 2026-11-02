"""Excel writes the PC's Office user name into "last modified by" when check_office.ps1 saves a workbook.
   This puts the credit name back (called at the end of check_office.ps1)."""
import os, re, zipfile, shutil, tempfile
DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'files')
for fn in sorted(os.listdir(DIR)):
    if not fn.endswith('.xlsx'): continue
    p = os.path.join(DIR, fn)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    with zipfile.ZipFile(p) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == 'docProps/core.xml':
                data = re.sub(rb'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', b'<cp:lastModifiedBy>Nivesh Sahayak</cp:lastModifiedBy>', data)
            elif it.filename == 'docProps/app.xml':
                data = re.sub(rb'<Company>.*?</Company>', b'<Company>Nivesh Sahayak</Company>', data)
                data = re.sub(rb'<Manager>.*?</Manager>', b'', data)
            zo.writestr(it, data)
    shutil.move(tmp, p)
print('file properties set to Nivesh Sahayak')
