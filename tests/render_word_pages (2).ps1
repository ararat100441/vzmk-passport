param([Parameter(Mandatory=$true)][string]$Directory)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Get-ChildItem -LiteralPath $Directory -Filter '*.emf' -File | ForEach-Object {
    $meta = [System.Drawing.Imaging.Metafile]::FromFile($_.FullName)
    $bitmap = New-Object System.Drawing.Bitmap(1240,1754)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.DrawImage($meta, (New-Object System.Drawing.Rectangle(0,0,1240,1754)))
        $bitmap.Save([System.IO.Path]::ChangeExtension($_.FullName,'.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
        $meta.Dispose()
    }
}
