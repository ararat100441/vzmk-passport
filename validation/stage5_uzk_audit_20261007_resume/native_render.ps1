$ErrorActionPreference='Stop'
$word=$null;$doc=$null;$results=@()
Add-Type -AssemblyName System.Drawing
$inputs=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'word_inputs.json') -Raw | ConvertFrom-Json
$labels=@('garage_first','garage_second','tempo_first')
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false;$word.DisplayAlerts=0;$word.AutomationSecurity=3;$word.Options.UpdateLinksAtOpen=$false
    for($di=0;$di -lt $inputs.Count;$di++){
        $inputItem=$inputs[$di]
        $before=(Get-FileHash -LiteralPath $inputItem.path -Algorithm SHA256).Hash.ToLowerInvariant()
        $doc=$word.Documents.Open($inputItem.path,$false,$true,$false)
        if(-not $doc.ReadOnly){throw 'STOP: not read-only'}
        $doc.Repaginate();$doc.ActiveWindow.View.Type=3
        $pageCount=$doc.ComputeStatistics(2)
        $ratio=$doc.PageSetup.PageWidth/$doc.PageSetup.PageHeight
        $width=1754;$height=[int][Math]::Round($width/$ratio)
        $images=@()
        for($pi=1;$pi -le $pageCount;$pi++){
            [byte[]]$pageBytes=$doc.ActiveWindow.Panes.Item(1).Pages.Item($pi).EnhMetaFileBits
            $memory=[IO.MemoryStream]::new($pageBytes,$false)
            $meta=[Drawing.Imaging.Metafile]::new($memory)
            $bitmap=[Drawing.Bitmap]::new($width,$height)
            $graphics=[Drawing.Graphics]::FromImage($bitmap);$graphics.Clear([Drawing.Color]::White)
            $graphics.DrawImage($meta,0,0,$width,$height)
            $name=$labels[$di]+'_'+$pi+'.png'
            $bitmap.Save((Join-Path $PSScriptRoot $name),[Drawing.Imaging.ImageFormat]::Png)
            $graphics.Dispose();$bitmap.Dispose();$meta.Dispose();$memory.Dispose()
            $images+=$name
        }
        $doc.Close(0);$doc=$null
        $after=(Get-FileHash -LiteralPath $inputItem.path -Algorithm SHA256).Hash.ToLowerInvariant()
        $results+=@{path=$inputItem.path;pages=$pageCount;images=$images;width=$width;height=$height;before=$before;after=$after;unchanged=($before -eq $after)}
    }
}finally{
    if($null -ne $doc){$doc.Close(0)}
    if($null -ne $word){$word.Quit(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
    $results | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'native_render.json') -Encoding utf8
    $results | ConvertTo-Json -Depth 6
}
