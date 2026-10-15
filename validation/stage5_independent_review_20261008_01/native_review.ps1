param([Parameter(Mandatory=$true)][string]$InputPath,[Parameter(Mandatory=$true)][string]$Prefix)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
$word=$null;$doc=$null;$owned=$false;$priorLinks=$null; $prePids=@(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id); Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ReviewWordPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p); }'
$result=[ordered]@{path=$InputPath;before_sha=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLower();status='STARTED';open_method='Documents.Open; normal; ReadOnly; no OpenAndRepair';macro_security=3;close_without_save=$true;pdf_export=$false}
try {
    $word=New-Object -ComObject Word.Application; $newPids=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object {$_.Id -notin $prePids} | ForEach-Object Id); if($newPids.Count -ne 1){throw "Could not prove ownership of new Word process; STOP"}; $owned=$true; $result.word_pid=$newPids[0]; $result.preexisting_word_pids=$prePids; if($word.Documents.Count -ne 0){throw "Unexpected documents in new Word process"}; $priorLinks=$word.Options.UpdateLinksAtOpen
    $word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3;$word.Options.UpdateLinksAtOpen=$false
    $doc=$word.Documents.Open($InputPath,$false,$true,$false)
    if(-not $doc.ReadOnly){throw 'Not ReadOnly'}
    $doc.Repaginate();$doc.ActiveWindow.View.Type=3
    $result.readonly=$doc.ReadOnly
    $result.pages=$doc.ComputeStatistics(2)
    $result.sections=$doc.Sections.Count
    $result.tables=$doc.Tables.Count
    $result.first_table_rows=$doc.Tables.Item(1).Rows.Count
    $result.hyperlinks=$doc.Hyperlinks.Count
    $result.text=$doc.Content.Text
    $result.orientation=$doc.PageSetup.Orientation
    $result.page_width=$doc.PageSetup.PageWidth
    $result.page_height=$doc.PageSetup.PageHeight
    $images=@()
    $width=1754;$height=[int][Math]::Round($width*$doc.PageSetup.PageHeight/$doc.PageSetup.PageWidth)
    for($pi=1;$pi -le $result.pages;$pi++){
        [byte[]]$pageBytes=$doc.ActiveWindow.Panes.Item(1).Pages.Item($pi).EnhMetaFileBits
        $memory=[IO.MemoryStream]::new($pageBytes,$false)
        $meta=[Drawing.Imaging.Metafile]::new($memory)
        $bitmap=[Drawing.Bitmap]::new($width,$height)
        $graphics=[Drawing.Graphics]::FromImage($bitmap);$graphics.Clear([Drawing.Color]::White)
        $graphics.DrawImage($meta,0,0,$width,$height)
        $imagePath=$Prefix+'_'+$pi+'.png'
        $bitmap.Save($imagePath,[Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose();$bitmap.Dispose();$meta.Dispose();$memory.Dispose()
        $images+=$imagePath
    }
    $result.images=$images
    $result.status='PASS_NORMAL_READONLY_OPEN'
} catch {
    $result.status='FAIL';$result.error=$_.Exception.Message
} finally {
    if($null -ne $doc){$doc.Close(0)}
    if($null -ne $word -and $owned){if($null -ne $priorLinks){$word.Options.UpdateLinksAtOpen=$priorLinks};$word.Quit(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    $result.after_sha=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLower()
    $result.unchanged=$result.before_sha -eq $result.after_sha
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath ($Prefix+'.json') -Encoding utf8
    [pscustomobject]$result | Select-Object status,error,pages,sections,tables,first_table_rows,hyperlinks,unchanged | ConvertTo-Json
}


