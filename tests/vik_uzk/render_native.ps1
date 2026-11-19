param([string]$Directory,[double]$PageWidth,[double]$PageHeight)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
$width=2480
$height=[int][Math]::Round($width*$PageHeight/$PageWidth)
foreach($page in Get-ChildItem -LiteralPath $Directory -Filter 'page-*.emf'){
    $outputImage=Join-Path $Directory ($page.BaseName+'.png')
    if(Test-Path -LiteralPath $outputImage){throw 'PNG exists; no overwrite'}
    $memory=[IO.MemoryStream]::new([IO.File]::ReadAllBytes($page.FullName),$false)
    $meta=[Drawing.Imaging.Metafile]::new($memory)
    $bitmap=[Drawing.Bitmap]::new($width,$height)
    $graphics=[Drawing.Graphics]::FromImage($bitmap)
    try{
        $graphics.Clear([Drawing.Color]::White)
        $graphics.DrawImage($meta,0,0,$width,$height)
        $bitmap.Save($outputImage,[Drawing.Imaging.ImageFormat]::Png)
    }finally{
        $graphics.Dispose();$bitmap.Dispose();$meta.Dispose();$memory.Dispose()
    }
}
