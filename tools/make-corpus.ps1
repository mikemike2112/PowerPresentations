# Genera il corpus tests/corpus/office-made con il vero PowerPoint (COM). Rigenerabile.
# Usa solo file creati qui; Author/LastAuthor = "Corpus" (post-elaborazione dello zip).
param([string]$OutDir = (Join-Path $PSScriptRoot '..\tests\corpus\office-made'), [string]$Only = '')
. (Join-Path $PSScriptRoot 'Common.ps1')
Add-Type -AssemblyName System.Drawing, System.IO.Compression, System.IO.Compression.FileSystem
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force $OutDir | Out-Null
$tmp = Join-Path $env:TEMP 'pp-corpus'; New-Item -ItemType Directory -Force $tmp | Out-Null

# immagine di prova generata qui (nessun dato personale)
$png = Join-Path $tmp 'sample.png'
$bmp = New-Object System.Drawing.Bitmap 400, 300
$g = [System.Drawing.Graphics]::FromImage($bmp)
$br = New-Object System.Drawing.Drawing2D.LinearGradientBrush ([System.Drawing.Point]::new(0, 0)), ([System.Drawing.Point]::new(400, 300)), ([System.Drawing.Color]::FromArgb(255, 80, 140, 255)), ([System.Drawing.Color]::FromArgb(255, 255, 122, 69))
$g.FillRectangle($br, 0, 0, 400, 300)
$g.DrawString('Corpus', (New-Object System.Drawing.Font('Arial', 40)), [System.Drawing.Brushes]::White, 90, 120)
$g.Dispose(); $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function RGB($r, $g, $b) { return $r + 256 * $g + 65536 * $b }
function Lay($pres, $i) { return $pres.SlideMaster.CustomLayouts.Item($i) }
function NewSlide($pres, $layout) { return $pres.Slides.AddSlide($pres.Slides.Count + 1, (Lay $pres $layout)) }
function Title($s, $t) { $s.Shapes.Title.TextFrame.TextRange.Text = $t }
function SetAuthor($pres) {
    $bt = [System.__ComObject]; $f = [Reflection.BindingFlags]
    foreach ($n in 'Author', 'Last Author') {
        try { $it = $bt.InvokeMember('Item', $f::GetProperty, $null, $pres.BuiltInDocumentProperties, @($n)); $bt.InvokeMember('Value', $f::SetProperty, $null, $it, @('Corpus')) | Out-Null } catch {}
    }
}
function Anonymize($path) {
    $zip = [System.IO.Compression.ZipFile]::Open($path, 'Update')
    try {
        foreach ($name in 'docProps/core.xml', 'docProps/app.xml') {
            $e = $zip.GetEntry($name); if (-not $e) { continue }
            $sr = New-Object System.IO.StreamReader($e.Open()); $x = $sr.ReadToEnd(); $sr.Close()
            $x = [regex]::Replace($x, '<dc:creator>.*?</dc:creator>', '<dc:creator>Corpus</dc:creator>')
            $x = [regex]::Replace($x, '<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy>Corpus</cp:lastModifiedBy>')
            $x = [regex]::Replace($x, '<Company>.*?</Company>', '<Company></Company>')
            $e.Delete(); $ne = $zip.CreateEntry($name)
            $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
        }
    } finally { $zip.Dispose() }
}

$defs = [ordered]@{}
$defs['title-slide'] = { param($p, $a)
    $s = NewSlide $p 1; Title $s 'Titolo della presentazione'; $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Sottotitolo di prova' }
$defs['layouts-all'] = { param($p, $a)
    for ($i = 1; $i -le 11; $i++) { $s = NewSlide $p $i; try { Title $s ("Layout " + $i) } catch {} } }
$defs['master-custom'] = { param($p, $a)
    $m = $p.SlideMaster; $m.Background.Fill.ForeColor.RGB = (RGB 250 240 230)
    $m.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = (RGB 160 30 30); $m.Shapes.Title.TextFrame.TextRange.Font.Size = 40
    $r = $m.Shapes.AddShape(1, 0, 500, 960, 40); $r.Fill.ForeColor.RGB = (RGB 160 30 30); $r.Line.Visible = 0
    $s = NewSlide $p 2; Title $s 'Master personalizzato'; $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Primo punto`rSecondo punto" }
$defs['theme-custom-colors'] = { param($p, $a)
    $cs = $p.SlideMaster.Theme.ThemeColorScheme
    $cs.Colors(5).RGB = (RGB 200 30 30); $cs.Colors(6).RGB = (RGB 30 160 60); $cs.Colors(7).RGB = (RGB 30 60 200)
    $s = NewSlide $p 6; Title $s 'Colori del tema'
    for ($i = 0; $i -lt 3; $i++) { $sh = $s.Shapes.AddShape(1, 60 + $i * 200, 200, 160, 120); $sh.Fill.ForeColor.ObjectThemeColor = 5 + $i } }
$defs['theme-custom-fonts'] = { param($p, $a)
    $fs = $p.SlideMaster.Theme.ThemeFontScheme
    $fs.MajorFont.Item(1).Name = 'Georgia'; $fs.MinorFont.Item(1).Name = 'Verdana'
    $s = NewSlide $p 2; Title $s 'Font del tema'; $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Testo in font minore del tema' }
$defs['textbox-basic'] = { param($p, $a)
    $s = NewSlide $p 7; $t = $s.Shapes.AddTextbox(1, 100, 100, 400, 60); $t.TextFrame.TextRange.Text = 'Casella di testo semplice'
    $t2 = $s.Shapes.AddTextbox(1, 100, 250, 300, 100); $t2.TextFrame.WordWrap = -1; $t2.TextFrame.TextRange.Text = 'Un testo molto più lungo che deve andare a capo automaticamente dentro la casella stretta.' }
$defs['text-formatting'] = { param($p, $a)
    $s = NewSlide $p 7; $t = $s.Shapes.AddTextbox(1, 60, 60, 800, 300); $tr = $t.TextFrame.TextRange
    $tr.Text = "Grassetto`rCorsivo`rSottolineato`rGrande rosso`rPiccolo`rVerdana blu`rApice e pedice"
    $tr.Paragraphs(1).Font.Bold = -1; $tr.Paragraphs(2).Font.Italic = -1; $tr.Paragraphs(3).Font.Underline = -1
    $tr.Paragraphs(4).Font.Size = 36; $tr.Paragraphs(4).Font.Color.RGB = (RGB 220 0 0); $tr.Paragraphs(5).Font.Size = 10
    $tr.Paragraphs(6).Font.Name = 'Verdana'; $tr.Paragraphs(6).Font.Color.RGB = (RGB 0 0 200); $tr.Paragraphs(7).Characters(1, 5).Font.Superscript = -1 }
$defs['bullets-numbered'] = { param($p, $a)
    $s = NewSlide $p 2; Title $s 'Elenchi'; $tr = $s.Shapes.Placeholders(2).TextFrame.TextRange
    $tr.Text = "Uno`rDue`rSotto-punto`rTre"; $tr.Paragraphs(3).IndentLevel = 2
    $t = $s.Shapes.AddTextbox(1, 500, 380, 400, 120); $t.TextFrame.TextRange.Text = "Primo`rSecondo`rTerzo"
    $t.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1; $t.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 2 }
$defs['paragraph-align-spacing'] = { param($p, $a)
    $s = NewSlide $p 7; $t = $s.Shapes.AddTextbox(1, 60, 60, 800, 360); $tr = $t.TextFrame.TextRange
    $tr.Text = "Sinistra`rCentro`rDestra`rInterlinea 2,0 prima riga`rseconda riga dello stesso paragrafo molto lungo da andare a capo davvero più volte"
    $tr.Paragraphs(2).ParagraphFormat.Alignment = 2; $tr.Paragraphs(3).ParagraphFormat.Alignment = 3
    $tr.Paragraphs(4).ParagraphFormat.SpaceWithin = 2; $tr.Paragraphs(1).ParagraphFormat.SpaceAfter = 12 }
$defs['shapes-basic'] = { param($p, $a)
    $s = NewSlide $p 7; $types = 1, 5, 9, 7, 8, 4, 2, 3, 10, 6, 11, 33, 34, 35, 36, 92, 52, 51, 21, 22
    $i = 0; foreach ($ty in $types) { $x = 30 + ($i % 5) * 180; $y = 30 + [math]::Floor($i / 5) * 125; $sh = $s.Shapes.AddShape($ty, $x, $y, 150, 100); $sh.TextFrame.TextRange.Text = "$ty"; $i++ } }
$defs['shape-fills-lines'] = { param($p, $a)
    $s = NewSlide $p 7
    $a1 = $s.Shapes.AddShape(1, 50, 50, 250, 150); $a1.Fill.TwoColorGradient(1, 1); $a1.Fill.ForeColor.RGB = (RGB 255 200 0); $a1.Fill.BackColor.RGB = (RGB 200 0 100)
    $a2 = $s.Shapes.AddShape(5, 350, 50, 250, 150); $a2.Fill.ForeColor.RGB = (RGB 0 150 200); $a2.Fill.Transparency = 0.5; $a2.Line.Weight = 6; $a2.Line.ForeColor.RGB = (RGB 0 0 100)
    $a3 = $s.Shapes.AddShape(9, 650, 50, 250, 150); $a3.Fill.Visible = 0; $a3.Line.DashStyle = 4; $a3.Line.Weight = 3
    $a4 = $s.Shapes.AddShape(1, 50, 300, 250, 150); $a4.Fill.Patterned(10); $a4.Line.Visible = 0 }
$defs['shape-rotation-flip'] = { param($p, $a)
    $s = NewSlide $p 7; $a1 = $s.Shapes.AddShape(33, 100, 150, 250, 100); $a1.Rotation = 30; $a1.TextFrame.TextRange.Text = 'Ruotata'
    $a2 = $s.Shapes.AddShape(33, 500, 150, 250, 100); $a2.Flip(0); $a2.TextFrame.TextRange.Text = 'Specchiata'
    $a3 = $s.Shapes.AddShape(7, 300, 330, 200, 150); $a3.Flip(1) }
$defs['connectors-lines'] = { param($p, $a)
    $s = NewSlide $p 7; $l = $s.Shapes.AddLine(100, 100, 400, 100); $l.Line.Weight = 3; $l.Line.EndArrowheadStyle = 2
    $b1 = $s.Shapes.AddShape(1, 100, 200, 120, 60); $b2 = $s.Shapes.AddShape(1, 400, 350, 120, 60)
    $c = $s.Shapes.AddConnector(2, 0, 0, 10, 10); $c.ConnectorFormat.BeginConnect($b1, 4); $c.ConnectorFormat.EndConnect($b2, 1); $c.Line.EndArrowheadStyle = 2 }
$defs['picture-basic'] = { param($p, $a)
    $s = NewSlide $p 6; Title $s 'Immagine'; $null = $s.Shapes.AddPicture($png, 0, -1, 200, 150, 400, 300) }
$defs['picture-crop-rotate'] = { param($p, $a)
    $s = NewSlide $p 7; $pic = $s.Shapes.AddPicture($png, 0, -1, 100, 100, 400, 300); $pic.PictureFormat.CropLeft = 100; $pic.PictureFormat.CropTop = 50
    $p2 = $s.Shapes.AddPicture($png, 0, -1, 550, 100, 200, 150); $p2.Rotation = 15; $p2.Line.Weight = 4 }
$defs['table-basic'] = { param($p, $a)
    $s = NewSlide $p 6; Title $s 'Tabella'; $t = $s.Shapes.AddTable(4, 3, 80, 150, 800, 240)
    for ($r = 1; $r -le 4; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text = "R${r}C$c" } } }
$defs['table-merged-cells'] = { param($p, $a)
    $s = NewSlide $p 7; $t = $s.Shapes.AddTable(4, 4, 80, 100, 800, 300)
    for ($r = 1; $r -le 4; $r++) { for ($c = 1; $c -le 4; $c++) { $t.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text = "$r,$c" } }
    $t.Table.Cell(1, 1).Merge($t.Table.Cell(1, 3)); $t.Table.Cell(2, 4).Merge($t.Table.Cell(4, 4))
    $t.Table.Cell(3, 2).Shape.Fill.ForeColor.RGB = (RGB 255 230 100) }
$defs['table-style-nobands'] = { param($p, $a)
    $s = NewSlide $p 7; $t = $s.Shapes.AddTable(3, 3, 80, 100, 600, 200); $t.Table.FirstRow = 0; $t.Table.HorizBanding = 0
    $t.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'Senza stile marcato'; $t.Table.Cell(2, 2).Shape.TextFrame.TextRange.Font.Bold = -1 }
$defs['notes-speaker'] = { param($p, $a)
    $s = NewSlide $p 2; Title $s 'Con note'; $s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Prima nota del relatore.`rSeconda riga delle note."
    $s2 = NewSlide $p 6; Title $s2 'Seconda slide'; $s2.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Note della seconda slide' }
$defs['transitions'] = { param($p, $a)
    $codes = 3849, 1281, 3855, 2819, 1793, 2305
    foreach ($c in $codes) { $s = NewSlide $p 6; Title $s "Transizione $c"; try { $s.SlideShowTransition.EntryEffect = $c } catch {} } }
$defs['animations'] = { param($p, $a)
    $s = NewSlide $p 7; $sh = $s.Shapes.AddShape(1, 100, 100, 200, 100); $null = $s.TimeLine.MainSequence.AddEffect($sh, 10, 0, 1)
    $sh2 = $s.Shapes.AddShape(9, 400, 100, 200, 100); $null = $s.TimeLine.MainSequence.AddEffect($sh2, 2, 0, 3) }
$defs['hidden-slide'] = { param($p, $a)
    $s = NewSlide $p 6; Title $s 'Visibile'; $h = NewSlide $p 6; Title $h 'Nascosta'; $h.SlideShowTransition.Hidden = -1; $s3 = NewSlide $p 6; Title $s3 'Visibile 2' }
$defs['sections'] = { param($p, $a)
    1..4 | ForEach-Object { $s = NewSlide $p 6; Title $s "Slide $_" }
    $null = $p.SectionProperties.AddBeforeSlide(1, 'Introduzione'); $null = $p.SectionProperties.AddBeforeSlide(3, 'Sviluppo') }
$defs['slide-background'] = { param($p, $a)
    $s = NewSlide $p 6; Title $s 'Sfondo colore'; $s.FollowMasterBackground = 0; $s.Background.Fill.ForeColor.RGB = (RGB 30 40 80)
    $s.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = (RGB 255 255 255)
    $s2 = NewSlide $p 6; Title $s2 'Sfondo sfumato'; $s2.FollowMasterBackground = 0; $s2.Background.Fill.TwoColorGradient(1, 1); $s2.Background.Fill.ForeColor.RGB = (RGB 255 220 200); $s2.Background.Fill.BackColor.RGB = (RGB 200 220 255) }
$defs['autofit-shrink'] = { param($p, $a)
    $s = NewSlide $p 2; Title $s 'Testo che si rimpicciolisce'
    $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = (1..14 | ForEach-Object { "Punto numero $_ con del testo abbastanza lungo da occupare spazio" }) -join "`r"
    $s.Shapes.Placeholders(2).TextFrame2.AutoSize = 2 }
$defs['group-shapes'] = { param($p, $a)
    $s = NewSlide $p 7; $a1 = $s.Shapes.AddShape(1, 100, 100, 150, 100); $a1.Name = 'g1'; $a2 = $s.Shapes.AddShape(9, 200, 150, 150, 100); $a2.Name = 'g2'
    $a2.Fill.ForeColor.RGB = (RGB 0 160 80); $g = $s.Shapes.Range([object[]]@('g1', 'g2')).Group(); $g.Left = 400; $g.Top = 200 }
$defs['hyperlinks'] = { param($p, $a)
    $s = NewSlide $p 7; $t = $s.Shapes.AddTextbox(1, 100, 100, 600, 60); $t.TextFrame.TextRange.Text = 'Collegamento di prova'
    $t.TextFrame.TextRange.ActionSettings(1).Hyperlink.Address = 'https://example.org/' }
$defs['comments'] = { param($p, $a)
    $s = NewSlide $p 6; Title $s 'Con commento'; try { $null = $s.Comments.Add2(100, 100, 'Corpus', 'C', 'Commento di prova', 'AD', 'Corpus') } catch { $null = $s.Comments.Add(100, 100, 'Corpus', 'C', 'Commento di prova') } }
$defs['footer-slidenumber'] = { param($p, $a)
    $s = NewSlide $p 6; Title $s 'Piè di pagina'; $s.HeadersFooters.SlideNumber.Visible = -1; $s.HeadersFooters.Footer.Visible = -1; $s.HeadersFooters.Footer.Text = 'Piè di pagina di prova' }
$defs['custom-layout'] = { param($p, $a)
    $cl = $p.SlideMaster.CustomLayouts.Add($p.SlideMaster.CustomLayouts.Count + 1); $cl.Name = 'Layout personalizzato'
    $r = $cl.Shapes.AddShape(9, 800, 400, 100, 100); $r.Fill.ForeColor.RGB = (RGB 255 122 69)
    $s = $p.Slides.AddSlide(1, $cl); try { $s.Shapes.Title.TextFrame.TextRange.Text = 'Layout personalizzato' } catch {} }
$defs['size-4x3'] = { param($p, $a)
    $p.PageSetup.SlideWidth = 720; $p.PageSetup.SlideHeight = 540; $s = NewSlide $p 2; Title $s 'Formato 4:3'; $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Contenuto' }
$defs['shadow-effects'] = { param($p, $a)
    $s = NewSlide $p 7; $a1 = $s.Shapes.AddShape(5, 100, 100, 300, 150); $a1.Shadow.Visible = -1; $a1.Shadow.OffsetX = 6; $a1.Shadow.OffsetY = 6
    $a2 = $s.Shapes.AddShape(1, 500, 100, 300, 150); $a2.ThreeD.BevelTopType = 3; $a2.Glow.Radius = 8 }
# smartart: AddSmartArt non disponibile via COM su questa installazione (SmartArtLayouts vuoto) -> escluso dal corpus, vedi COMPAT.md
$defs['wordart'] = { param($p, $a)
    $s = NewSlide $p 7; $null = $s.Shapes.AddTextEffect(0, 'WordArt', 'Arial', 60, 0, 0, 100, 100) }
$defs['mixed-realistic-1'] = { param($p, $a)
    $s = NewSlide $p 1; Title $s 'Piano trimestrale'; $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Reparto vendite'
    $s2 = NewSlide $p 2; Title $s2 'Obiettivi'; $s2.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Crescita del 10%`rNuovi mercati`rRidurre i costi"
    $s2.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Spiegare i tre obiettivi'
    $s3 = NewSlide $p 6; Title $s3 'Risultati'; $t = $s3.Shapes.AddTable(3, 3, 80, 150, 500, 150)
    for ($r = 1; $r -le 3; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text = "$r-$c" } }
    $null = $s3.Shapes.AddPicture($png, 0, -1, 620, 150, 280, 210); $s3.SlideShowTransition.EntryEffect = 3849 }
$defs['mixed-realistic-2'] = { param($p, $a)
    $s = NewSlide $p 3; Title $s 'Sezione'; $s.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Dettagli'
    $s2 = NewSlide $p 4; Title $s2 'Due contenuti'; $s2.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Sinistra A`rSinistra B"; $s2.Shapes.Placeholders(3).TextFrame.TextRange.Text = "Destra A`rDestra B"
    $s3 = NewSlide $p 5; Title $s3 'Confronto'
    $s4 = NewSlide $p 7; $sh = $s4.Shapes.AddShape(5, 100, 100, 300, 120); $sh.TextFrame.TextRange.Text = 'Forma con testo'; $sh.Fill.ForeColor.RGB = (RGB 255 122 69)
    $s4.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Note finali' }

$ctx = Start-Ppt -TimeoutSec 90
$app = $ctx.App
$log = @()
try {
    foreach ($k in $defs.Keys) {
        if ($Only -and $k -ne $Only) { continue }
        Ping-Ppt $ctx
        $out = Join-Path $OutDir ($k + '.pptx')
        $pres = $null
        try {
            $pres = $app.Presentations.Add(0)
            SetAuthor $pres
            & $defs[$k] $pres $app
            if (Test-Path $out) { Remove-Item $out }
            $pres.SaveAs($out)
            $pres.Close(); $pres = $null
            Anonymize $out
            $log += "OK   $k"
        } catch {
            $log += "FAIL $k : $($_.Exception.Message)"
            try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
        }
    }
} finally { Stop-Ppt $ctx }
$log
