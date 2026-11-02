# Opens every file in files\ with real Excel / Word, recalculates, lists any error cells (#VALUE!, #REF!, #NAME?, ...),
# and saves the workbooks so the calculated values are stored in them. Run after make_files.py:
#   powershell -ExecutionPolicy Bypass -File scripts\check_office.ps1 [-Files a.xlsx,b.docx]   (-Files = only those files)
param([string[]]$Files)
$ErrorActionPreference = 'Stop'
$dir = Join-Path (Split-Path $PSScriptRoot -Parent) 'files'
$bad = 0
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  foreach ($f in Get-ChildItem $dir -Filter *.xlsx | Where-Object { -not $Files -or $Files -contains $_.Name }) {
    $wb = $xl.Workbooks.Open($f.FullName)
    $xl.CalculateFull()
    $errs = @()
    foreach ($ws in $wb.Worksheets) {
      $ur = $ws.UsedRange
      $vals = $ur.Value2
      if ($vals -is [array]) {
        $rows = $vals.GetLength(0); $cols = $vals.GetLength(1)
        for ($r = 1; $r -le $rows; $r++) { for ($c = 1; $c -le $cols; $c++) {
          $v = $vals[$r, $c]
          if ($v -is [int] -and $v -le -2146826200 -and $v -ge -2146826300) {
            $errs += ('{0}!{1}' -f $ws.Name, $ur.Cells($r, $c).Address($false, $false))
          } } }
      }
    }
    if ($errs.Count) { $bad += $errs.Count; Write-Output ("FAIL {0}: {1} error cells, e.g. {2}" -f $f.Name, $errs.Count, (($errs | Select-Object -First 8) -join ', ')) }
    else { Write-Output ("OK   {0}" -f $f.Name) }
    $wb.Save(); $wb.Close($false)
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
try {
  foreach ($f in Get-ChildItem $dir -Filter *.docx | Where-Object { -not $Files -or $Files -contains $_.Name }) {
    $d = $wd.Documents.Open($f.FullName, $false, $true)
    Write-Output ("OK   {0} ({1} pages)" -f $f.Name, $d.ComputeStatistics(2))
    $d.Close($false)
  }
} finally { $wd.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wd) }
python -X utf8 (Join-Path $PSScriptRoot 'fix_meta.py') @($Files | Where-Object { $_ -like '*.xlsx' })
if ($bad) { Write-Output "$bad error cells"; exit 1 } else { Write-Output 'all files open and calculate without errors' }
