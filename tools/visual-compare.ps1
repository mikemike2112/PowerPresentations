# Confronto visivo: PDF da PowerPoint (COM) e PDF dal clone, rasterizzati con Windows.Data.Pdf (tools/PdfRaster),
# affiancati slide per slide in <OutDir>\<nome>.png con la differenza media di pixel. Poi si guardano le immagini.
# Uso: visual-compare.ps1 -OutDir dir [-Names a,b,c]
param([string]$OutDir, [string[]]$Names = @(), [int]$Width = 960)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$corpus = Join-Path $repo 'tests\corpus\office-made'
$app = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
$ras = Join-Path $repo 'tools\PdfRaster\bin\Debug\net10.0-windows10.0.19041.0\PdfRaster.exe'
$OutDir = [IO.Path]::GetFullPath($OutDir)
$names = @($Names | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
if ($names.Count -eq 0) { $names = Get-ChildItem $corpus -Filter *.pptx | ForEach-Object { $_.BaseName } }
$office = Join-Path $OutDir 'office-pdf'; $clone = Join-Path $OutDir 'clone-pdf'
New-Item -ItemType Directory -Force $OutDir, $office, $clone | Out-Null
$list = Join-Path $OutDir 'list.txt'
$names | ForEach-Object { Join-Path $corpus ($_ + '.pptx') } | Set-Content $list -Encoding UTF8
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'office-inspect.ps1') -ListFile $list -OutJson (Join-Path $OutDir 'office.json') -PdfDir $office -TimeoutSec 40
foreach ($n in $names) {
    $p = Start-Process $app -ArgumentList '--export-pdf', (Join-Path $corpus ($n + '.pptx')), (Join-Path $clone ($n + '.pdf')) -PassThru
    if (-not $p.WaitForExit(60000)) { $p.Kill(); Write-Output "TIMEOUT clone $n"; continue }
    foreach ($side in @(@($office, 'o'), @($clone, 'c'))) {
        $pdf = Join-Path $side[0] ($n + '.pdf'); if (-not (Test-Path $pdf)) { Write-Output "manca $pdf"; continue }
        $d = Join-Path $OutDir ($side[1] + '\' + $n)
        if (Test-Path $d) { Remove-Item $d -Recurse -Force }
        & $ras raster $pdf $d $Width | Out-Null
    }
    $r = & $ras compare (Join-Path $OutDir "o\$n") (Join-Path $OutDir "c\$n") (Join-Path $OutDir ($n + '.png')) 560
    Write-Output ("{0}: {1}" -f $n, $r)
}
