param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Destination
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$sourceImage = [System.Drawing.Image]::FromFile((Resolve-Path -LiteralPath $Source).Path)
$frames = @()
try {
    foreach ($size in @(16,20,24,32,48,64,128,256)) {
        $bitmap = New-Object System.Drawing.Bitmap($size,$size,[System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $buffer = New-Object System.IO.MemoryStream
        try {
            $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.DrawImage($sourceImage,0,0,$size,$size)
            $bitmap.Save($buffer,[System.Drawing.Imaging.ImageFormat]::Png)
            $frames += ,$buffer.ToArray()
        } finally { $graphics.Dispose(); $bitmap.Dispose(); $buffer.Dispose() }
    }
} finally { $sourceImage.Dispose() }
$output = [System.IO.File]::Open($Destination,[System.IO.FileMode]::CreateNew)
$writer = New-Object System.IO.BinaryWriter($output)
try {
    $writer.Write([uint16]0); $writer.Write([uint16]1); $writer.Write([uint16]$frames.Count)
    $offset = 6 + 16 * $frames.Count
    $sizes = @(16,20,24,32,48,64,128,256)
    for ($i=0; $i -lt $frames.Count; $i++) {
        $encodedSize = $sizes[$i] % 256
        $writer.Write([byte]$encodedSize); $writer.Write([byte]$encodedSize)
        $writer.Write([byte]0); $writer.Write([byte]0)
        $writer.Write([uint16]1); $writer.Write([uint16]32)
        $writer.Write([uint32]$frames[$i].Length); $writer.Write([uint32]$offset)
        $offset += $frames[$i].Length
    }
    foreach ($frame in $frames) { $writer.Write([byte[]]$frame) }
} finally { $writer.Dispose(); $output.Dispose() }
Get-Item -LiteralPath $Destination | Select-Object FullName,Length
