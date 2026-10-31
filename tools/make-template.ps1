# Genera il modello incorporato "Nuova presentazione" (src/PowerPresentations.Core/Template.pptx) col vero PowerPoint.
. (Join-Path $PSScriptRoot 'Common.ps1')
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$out = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\src\PowerPresentations.Core\Template.pptx'))
$ctx = Start-Ppt -TimeoutSec 60
try {
    $app = $ctx.App; $p = $app.Presentations.Add(0)
    $fs = $p.SlideMaster.Theme.ThemeFontScheme; $fs.MajorFont.Item(1).Name = 'Calibri Light'; $fs.MinorFont.Item(1).Name = 'Calibri'
    $null = $p.Slides.AddSlide(1, $p.SlideMaster.CustomLayouts.Item(1))
    if (Test-Path $out) { Remove-Item $out }
    $p.SaveAs($out); $p.Close()
} finally { Stop-Ppt $ctx }
$zip = [System.IO.Compression.ZipFile]::Open($out, 'Update')
try {
    foreach ($name in 'docProps/core.xml', 'docProps/app.xml', 'ppt/theme/theme1.xml') {
        $e = $zip.GetEntry($name); if (-not $e) { continue }
        $sr = New-Object System.IO.StreamReader($e.Open()); $x = $sr.ReadToEnd(); $sr.Close()
        $x = [regex]::Replace($x, '<dc:creator>.*?</dc:creator>', '<dc:creator>PowerPresentations</dc:creator>')
        $x = [regex]::Replace($x, '<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy>PowerPresentations</cp:lastModifiedBy>')
        $x = $x -replace '<a:theme ([^>]*)name="[^"]*"', '<a:theme $1name="Tema PowerPresentations"'
        $e.Delete(); $ne = $zip.CreateEntry($name)
        $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
    }
} finally { $zip.Dispose() }
"template: $out"
