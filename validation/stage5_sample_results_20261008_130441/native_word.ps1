param([Parameter(Mandatory=$true)][string]$InputPath)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class SampleWordPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$prefix=Join-Path $PSScriptRoot 'native_word'
$existingPids=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word=$null; $doc=$null; $owned=$false
$result=[ordered]@{path=$InputPath;before_sha=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLower();status='STARTED';open_method='Documents.Open(path,False,True,False); OpenAndRepair=False (API default); ReadOnly=True';close_without_save=$true;pdf_export=$false;existing_word_pids=$existingPids;owned_instance=$false}
try {
    $word=New-Object -ComObject Word.Application
    $newPids=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object {$existingPids -notcontains $_.Id} | Select-Object -ExpandProperty Id)
    if($newPids.Count -ne 1 -or $word.Documents.Count -ne 0){throw 'Word instance is not exclusively owned; user instance left untouched'}
    [uint32]$taskWordPid=$newPids[0]
    $owned=$true; $result.owned_instance=$true; $result.word_pid=$taskWordPid
    $result.version=$word.Version; $result.build=$word.Build
    $word.Visible=$false; $word.DisplayAlerts=-1; $word.AutomationSecurity=3; $word.Options.UpdateLinksAtOpen=$false
    $doc=$word.Documents.Open($InputPath,$false,$true,$false)
    [uint32]$openedWordPid=0
    [void][SampleWordPid]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd,[ref]$openedWordPid)
    if($openedWordPid -ne $taskWordPid){throw 'Opened document is outside the owned Word process'}
    if(-not $doc.ReadOnly){throw 'Document did not open ReadOnly'}
    $doc.Repaginate(); $doc.ActiveWindow.View.Type=3
    $result.readonly=$doc.ReadOnly
    $result.pages=$doc.ComputeStatistics(2); $result.sections=$doc.Sections.Count
    $result.tables=$doc.Tables.Count; $result.first_table_rows=$doc.Tables.Item(1).Rows.Count
    $result.hyperlinks=$doc.Hyperlinks.Count; $result.text=$doc.Content.Text
    $result.orientation=$doc.PageSetup.Orientation; $result.page_width=$doc.PageSetup.PageWidth; $result.page_height=$doc.PageSetup.PageHeight
    $result.page_texts=@()
    $images=@(); $width=2480; $height=[int][Math]::Round($width*$doc.PageSetup.PageHeight/$doc.PageSetup.PageWidth)
    for($pi=1;$pi -le $result.pages;$pi++){
        $start=$doc.GoTo(1,1,$pi).Start
        if($pi -lt $result.pages){$end=$doc.GoTo(1,1,$pi+1).Start}else{$end=$doc.Content.End}
        $result.page_texts+=@($doc.Range($start,$end).Text)
        [byte[]]$pageBytes=$doc.ActiveWindow.Panes.Item(1).Pages.Item($pi).EnhMetaFileBits
        $memory=[IO.MemoryStream]::new($pageBytes,$false); $meta=[Drawing.Imaging.Metafile]::new($memory)
        $bitmap=[Drawing.Bitmap]::new($width,$height); $graphics=[Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([Drawing.Color]::White); $graphics.DrawImage($meta,0,0,$width,$height)
        $imagePath=Join-Path $PSScriptRoot ('page-'+$pi+'.png')
        $bitmap.Save($imagePath,[Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose(); $bitmap.Dispose(); $meta.Dispose(); $memory.Dispose(); $images+=$imagePath
    }
    $result.images=$images; $result.status='PASS_NORMAL_READONLY_OPEN'
} catch {$result.status='FAIL'; $result.error=$_.Exception.Message}
finally {
    if($null -ne $doc){$doc.Close(0)}
    if($owned -and $null -ne $word){$word.Quit(0)}
    if($null -ne $word){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    $result.after_sha=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLower()
    $result.unchanged=$result.before_sha -eq $result.after_sha
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath ($prefix+'.json') -Encoding utf8
    [pscustomobject]$result | Select-Object status,error,pages,sections,tables,first_table_rows,hyperlinks,unchanged,owned_instance | ConvertTo-Json
}
if($result.status -ne 'PASS_NORMAL_READONLY_OPEN' -or -not $result.unchanged){exit 1}
