$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$reviewRoot = $PSScriptRoot
foreach ($reviewLabel in @('new','first','manual')) {
    $nativeFolder = Join-Path $reviewRoot ('native_' + $reviewLabel)
    foreach ($emfFile in Get-ChildItem -LiteralPath $nativeFolder -Filter '*.emf' -File) {
        $memory = New-Object System.IO.MemoryStream(,[IO.File]::ReadAllBytes($emfFile.FullName))
        $metafile = [Drawing.Imaging.Metafile]::new($memory)
        if ($emfFile.BaseName.StartsWith('page-')) {
            $renderWidth = 2480
            $renderHeight = [int][math]::Round($renderWidth * 841.9 / 595.3)
        } else {
            $renderWidth = 2200
            $renderHeight = [int][math]::Round($renderWidth * $metafile.Height / $metafile.Width)
        }
        $bitmap = [Drawing.Bitmap]::new($renderWidth, $renderHeight)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([Drawing.Color]::White)
            $graphics.DrawImage($metafile, 0, 0, $renderWidth, $renderHeight)
            $bitmap.Save((Join-Path $nativeFolder ($emfFile.BaseName + '.png')), [Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $graphics.Dispose(); $bitmap.Dispose(); $metafile.Dispose(); $memory.Dispose()
        }
        Write-Output ($reviewLabel + '/' + $emfFile.BaseName + '.png')
    }
}
