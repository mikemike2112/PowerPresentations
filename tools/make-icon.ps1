# Genera src/PowerPresentations.App/icon.ico (16,24,32,48,64,128,256) da assets/icon-source.png (icona fornita dall'autore del progetto).
# Gli angoli esterni diventano trasparenti (flood fill dei pixel quasi bianchi/trasparenti collegati ai bordi). Stampa il colore di sfondo campionato.
Add-Type -AssemblyName System.Drawing
$repo = Split-Path $PSScriptRoot -Parent
$src = [System.Drawing.Bitmap]::FromFile((Join-Path $repo 'assets\icon-source.png'))
$n = 256
$base = New-Object System.Drawing.Bitmap $n, $n, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($base); $g.InterpolationMode = 'HighQualityBicubic'; $g.SmoothingMode = 'AntiAlias'; $g.PixelOffsetMode = 'HighQuality'
$g.Clear([System.Drawing.Color]::Transparent); $g.DrawImage($src, 0, 0, $n, $n); $g.Dispose()
# trasparenza degli angoli: pixel chiari collegati ai bordi
$vis = New-Object 'bool[,]' $n, $n; $q = New-Object System.Collections.Generic.Queue[object]
function IsBg($c) { $c.A -lt 40 -or ($c.R -gt 235 -and $c.G -gt 235 -and $c.B -gt 235) }
foreach ($i in 0..($n - 1)) { foreach ($pt in @(@($i, 0), @($i, ($n - 1)), @(0, $i), @(($n - 1), $i))) { $q.Enqueue($pt) } }
while ($q.Count -gt 0) {
    $p = $q.Dequeue(); $x = $p[0]; $y = $p[1]
    if ($x -lt 0 -or $y -lt 0 -or $x -ge $n -or $y -ge $n -or $vis[$x, $y]) { continue }
    if (-not (IsBg $base.GetPixel($x, $y))) { continue }
    $vis[$x, $y] = $true; $base.SetPixel($x, $y, [System.Drawing.Color]::Transparent)
    $q.Enqueue(@(($x + 1), $y)); $q.Enqueue(@(($x - 1), $y)); $q.Enqueue(@($x, ($y + 1))); $q.Enqueue(@($x, ($y - 1)))
}
$c = $base.GetPixel(30, 128); "accento campionato: #{0:X2}{1:X2}{2:X2}" -f $c.R, $c.G, $c.B
$sizes = 16, 24, 32, 48, 64, 128, 256; $pngs = @()
foreach ($s in $sizes) {
    $b = New-Object System.Drawing.Bitmap $s, $s, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $gg = [System.Drawing.Graphics]::FromImage($b); $gg.InterpolationMode = 'HighQualityBicubic'; $gg.SmoothingMode = 'AntiAlias'; $gg.PixelOffsetMode = 'HighQuality'; $gg.Clear([System.Drawing.Color]::Transparent)
    $gg.DrawImage($base, 0, 0, $s, $s); $gg.Dispose()
    $ms = New-Object System.IO.MemoryStream; $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $pngs += , $ms.ToArray(); $b.Dispose()
}
$out = Join-Path $repo 'src\PowerPresentations.App\icon.ico'
$fs = [System.IO.File]::Create($out); $w = New-Object System.IO.BinaryWriter $fs
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$sizes.Count)
$off = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) { $s = $sizes[$i]; $d = if ($s -ge 256) { 0 } else { $s }; $w.Write([byte]$d); $w.Write([byte]$d); $w.Write([byte]0); $w.Write([byte]0); $w.Write([uint16]1); $w.Write([uint16]32); $w.Write([uint32]$pngs[$i].Length); $w.Write([uint32]$off); $off += $pngs[$i].Length }
foreach ($p in $pngs) { $w.Write($p) }
$w.Close(); $fs.Close()
# anteprima per controllo visivo (tutte le dimensioni su fondo a scacchi)
$sheet = New-Object System.Drawing.Bitmap 640, 300; $sg = [System.Drawing.Graphics]::FromImage($sheet); $sg.Clear([System.Drawing.Color]::FromArgb(60, 60, 60))
$x = 10; foreach ($i in 0..($sizes.Count - 1)) { $im = [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream (, $pngs[$i]))); $sg.DrawImage($im, $x, 10, $sizes[$i], $sizes[$i]); $x += $sizes[$i] + 10 }
$sheet.Save((Join-Path $env:TEMP 'pp-icon-sheet.png')); "ico: $out"
