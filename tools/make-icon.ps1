# Genera l'icona originale dell'app (src/PowerPresentations.App/icon.ico): schermo di presentazione con triangolo "play" su fondo arancio.
Add-Type -AssemblyName System.Drawing
$out = Join-Path $PSScriptRoot '..\src\PowerPresentations.App\icon.ico'
$s = 256
$bmp = New-Object System.Drawing.Bitmap $s, $s
$g = [System.Drawing.Graphics]::FromImage($bmp); $g.SmoothingMode = 'AntiAlias'
$g.Clear([System.Drawing.Color]::Transparent)
$path = New-Object System.Drawing.Drawing2D.GraphicsPath
$r = 48; $path.AddArc(8, 8, $r, $r, 180, 90); $path.AddArc($s - 8 - $r, 8, $r, $r, 270, 90); $path.AddArc($s - 8 - $r, $s - 8 - $r, $r, $r, 0, 90); $path.AddArc(8, $s - 8 - $r, $r, $r, 90, 90); $path.CloseFigure()
$g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 122, 69))), $path)
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 14, 14, 14)), 12
$g.DrawRectangle($pen, 52, 70, 152, 108)
$g.DrawLine($pen, 128, 178, 128, 204); $g.DrawLine($pen, 98, 206, 158, 206)
$tri = [System.Drawing.Point[]]@([System.Drawing.Point]::new(112, 98), [System.Drawing.Point]::new(112, 150), [System.Drawing.Point]::new(156, 124))
$g.FillPolygon((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 14, 14, 14))), $tri)
$g.Dispose()
$ms = New-Object System.IO.MemoryStream; $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $png = $ms.ToArray()
$fs = [System.IO.File]::Create($out); $w = New-Object System.IO.BinaryWriter $fs
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]1)
$w.Write([byte]0); $w.Write([byte]0); $w.Write([byte]0); $w.Write([byte]0); $w.Write([uint16]1); $w.Write([uint16]32); $w.Write([uint32]$png.Length); $w.Write([uint32]22)
$w.Write($png); $w.Close(); $fs.Close()
"icon: $out"
