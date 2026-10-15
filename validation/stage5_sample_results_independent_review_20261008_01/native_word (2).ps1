param([Parameter(Mandatory=$true)][string]$InputPath)
$ErrorActionPreference='Stop'
$outDir=Join-Path $PSScriptRoot 'native_word'
[void][IO.Directory]::CreateDirectory($outDir)
$env:TEMP=Join-Path $outDir 'tmp'
$env:TMP=$env:TEMP
[void][IO.Directory]::CreateDirectory($env:TEMP)
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class IndependentWordPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$existing=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word=$null; $doc=$null; $owned=$false
$result=[ordered]@{path=$InputPath;before_sha=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLower();status='STARTED';open_method='Word object model Documents.Open(path,False,True,False): ReadOnly=True; AddToRecentFiles=False; OpenAndRepair=False (API default)';close_without_save=$false;existing_word_pids=$existing;owned_instance=$false;pdf_export=$false;ui_input=$false}
try {
    if($result.before_sha -ne '8d3a2769baf6d381a0675804ef7f0806896f1b611e1498ea58677d3fdac1fd1b'){throw 'STOP: exact TEST DOCX SHA mismatch'}
    $word=New-Object -ComObject Word.Application
    $new=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object {$existing -notcontains $_.Id} | Select-Object -ExpandProperty Id)
    if($new.Count -ne 1 -or $word.Documents.Count -ne 0){throw 'Instance not exclusively owned; user Word left untouched'}
    $owned=$true; $result.owned_instance=$true; $result.word_pid=$new[0]
    $word.Visible=$false; $word.DisplayAlerts=-1; $word.AutomationSecurity=3
    $result.version=$word.Version; $result.build=$word.Build
    $doc=$word.Documents.Open($InputPath,$false,$true,$false)
    [uint32]$openedPid=0
    [void][IndependentWordPid]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd,[ref]$openedPid)
    if($openedPid -ne $result.word_pid){throw 'Document outside owned process'}
    if(-not $doc.ReadOnly){throw 'Document not opened ReadOnly'}
    $doc.Repaginate(); $doc.ActiveWindow.View.Type=3
    $result.readonly=$doc.ReadOnly; $result.pages=$doc.ComputeStatistics(2)
    $result.sections=$doc.Sections.Count; $result.tables=$doc.Tables.Count
    $result.first_table_rows=$doc.Tables.Item(1).Rows.Count; $result.hyperlinks=$doc.Hyperlinks.Count
    $result.orientation=$doc.PageSetup.Orientation; $result.page_width=$doc.PageSetup.PageWidth; $result.page_height=$doc.PageSetup.PageHeight
    $result.text=$doc.Content.Text; $result.page_texts=@(); $result.images=@()
    $width=2480; $height=[int][Math]::Round($width*$result.page_height/$result.page_width)
    for($pi=1;$pi -le $result.pages;$pi++){
        $start=$doc.GoTo(1,1,$pi).Start
        if($pi -lt $result.pages){$end=$doc.GoTo(1,1,$pi+1).Start}else{$end=$doc.Content.End}
        $result.page_texts+=@($doc.Range($start,$end).Text)
        [byte[]]$pageBytes=$doc.ActiveWindow.Panes.Item(1).Pages.Item($pi).EnhMetaFileBits
        $memory=[IO.MemoryStream]::new($pageBytes,$false)
        $meta=[Drawing.Imaging.Metafile]::new($memory)
        $bitmap=[Drawing.Bitmap]::new($width,$height); $graphics=[Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([Drawing.Color]::White); $graphics.DrawImage($meta,0,0,$width,$height)
        $imagePath=Join-Path $outDir ('page-'+$pi+'.png')
        $bitmap.Save($imagePath,[Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose(); $bitmap.Dispose(); $meta.Dispose(); $memory.Dispose()
        $result.images+=$imagePath
    }
    $result.status='PASS_NORMAL_READONLY_OPEN'
}catch{$result.status='FAIL';$result.error=$_.Exception.Message}
finally {
    if($null -ne $doc){$doc.Close(0);$result.close_without_save=$true}
    if($owned -and $null -ne $word){$word.Quit(0)}
    if($null -ne $word){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    $result.after_sha=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLower()
    $result.unchanged=$result.before_sha -eq $result.after_sha
    $result.existing_word_processes_preserved=@($existing | Where-Object {Get-Process -Id $_ -ErrorAction SilentlyContinue}).Count -eq $existing.Count
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outDir 'native_word.json') -Encoding UTF8
    [pscustomobject]$result | Select-Object status,error,pages,sections,tables,first_table_rows,hyperlinks,readonly,unchanged,owned_instance,existing_word_processes_preserved | ConvertTo-Json
}
if($result.status -ne 'PASS_NORMAL_READONLY_OPEN' -or -not $result.unchanged -or -not $result.close_without_save){exit 1}
