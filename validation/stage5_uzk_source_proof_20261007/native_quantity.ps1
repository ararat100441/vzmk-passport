$ErrorActionPreference='Stop'
$appRoot=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$project=Join-Path (Split-Path $appRoot -Parent) 'Гараж казахстан'
$reports=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'project_uzk_documents.json') -Raw | ConvertFrom-Json
$excel=$null;$book=$null;$results=@()
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false
 foreach($passport in @('7 от 01.10.2026','8 от 02.10.2026')){
  $folder=Join-Path (Join-Path $project 'Паспорт') $passport
  $marks=@($reports | Where-Object { (Split-Path $_.path -Parent) -eq $folder } | ForEach-Object {$_.rows} | ForEach-Object {$_.mark})
  foreach($file in Get-ChildItem -LiteralPath $folder -File | Where-Object {$_.Name -like 'ВОМ*.xls*'}){
   $entry=[ordered]@{path=$file.FullName;before_sha=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant();macro_security=3;sheets=@()}
   try {
    $book=$excel.Workbooks.Open($file.FullName,0,$true)
    if(-not $book.ReadOnly){throw 'STOP workbook not read-only'}
    $entry.readonly=$book.ReadOnly
    foreach($sheet in $book.Worksheets){
     $used=$sheet.UsedRange;$raw=$used.Value2;$quantityRows=@();$visibleMass=0.0;$visibleQty=0.0
     for($i=1;$i -le $used.Rows.Count;$i++){
      $rowNumber=$used.Row+$i-1
      if($used.Columns.Count -lt 4){continue}
      $mark=$raw.GetValue($i,2)
      if($mark -in $marks){
       $quantityRows+=@{row=$rowNumber;mark=$mark;quantity=$raw.GetValue($i,4);hidden=$sheet.Rows.Item($rowNumber).Hidden;cells=@($raw.GetValue($i,1),$mark,$raw.GetValue($i,3),$raw.GetValue($i,4))}
      }
      if($i -ge 4 -and $mark -is [string] -and $mark -match '\d' -and -not $sheet.Rows.Item($rowNumber).Hidden){
       if($raw.GetValue($i,4) -is [double]){$visibleQty+=$raw.GetValue($i,4)}
       if($used.Columns.Count -ge 9 -and $raw.GetValue($i,9) -is [double]){$visibleMass+=$raw.GetValue($i,9)}
      }
     }
     $entry.sheets+=@{name=$sheet.Name;used=$used.Address($false,$false);filter_mode=$sheet.FilterMode;matches=$quantityRows;visible_mass=$visibleMass;visible_quantity=$visibleQty}
    }
    $entry.status='PASS_READ_ONLY_NATIVE_OPEN'
   } finally {if($null -ne $book){$book.Close($false);$book=$null}}
   $entry.after_sha=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
   $entry.unchanged=$entry.before_sha -eq $entry.after_sha
   if(-not $entry.unchanged){throw 'STOP VOM changed'}
   $results+=[pscustomobject]$entry
  }
 }
} finally {if($null -ne $book){$book.Close($false)};if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
$results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'historical_vom_quantities.json') -Encoding utf8
$results | ConvertTo-Json -Depth 10
