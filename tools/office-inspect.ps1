# Apre uno o piu' .pptx nel VERO PowerPoint (COM, WithWindow=msoFalse, sola lettura) e scrive un JSON con
# il contenuto letto da Office: n. slide, forme, testo, tabelle, note, transizioni. Rileva apertura fallita,
# dialoghi modali (watchdog per PID) e ripristino. Opzionale: PDF e PNG per slide (SaveCopyAs / Slide.Export).
# Uso: office-inspect.ps1 (-Files a.pptx,b.pptx | -ListFile elenco.txt) -OutJson out.json [-PdfDir dir] [-PngDir dir] [-TimeoutSec 25]
param([string[]]$Files, [string]$ListFile = '', [string]$OutJson, [string]$PdfDir = '', [string]$PngDir = '', [int]$TimeoutSec = 25)
. (Join-Path $PSScriptRoot 'Common.ps1')
if ($ListFile) { $Files = @(Get-Content $ListFile -Encoding UTF8 | Where-Object { $_ }) } else { $Files = @($Files | ForEach-Object { $_ -split "," }) }
$results = @()
$ctx = $null

function Read-Pres($pres, $path) {
    $o = [ordered]@{ slideCount = $pres.Slides.Count; width = $pres.PageSetup.SlideWidth; height = $pres.PageSetup.SlideHeight; slides = @(); sections = $pres.SectionProperties.Count; customShows = $pres.SlideShowSettings.NamedSlideShows.Count }
    foreach ($s in $pres.Slides) {
        $sl = [ordered]@{ index = $s.SlideIndex; layout = $s.CustomLayout.Name; hidden = [int]$s.SlideShowTransition.Hidden; transition = [int]$s.SlideShowTransition.EntryEffect; notes = ''; shapes = @(); comments = $s.Comments.Count; advOn = [int]$s.SlideShowTransition.AdvanceOnTime; advTime = [double]$s.SlideShowTransition.AdvanceTime }
        $sl.commentList = @(); try { foreach ($c in $s.Comments) { $rc = 0; try { $rc = $c.Replies.Count } catch {}; $sl.commentList += [ordered]@{ author = [string]$c.Author; text = [string]$c.Text; replies = $rc } } } catch {}
        try { foreach ($ns in $s.NotesPage.Shapes) { if ($ns.Type -eq 14 -and $ns.PlaceholderFormat.Type -eq 2) { $sl.notes = $ns.TextFrame.TextRange.Text } } } catch {}
        foreach ($sh in $s.Shapes) {
            $e = [ordered]@{ name = $sh.Name; type = [int]$sh.Type; text = ''; table = $null; l = [math]::Round($sh.Left, 1); t = [math]::Round($sh.Top, 1); w = [math]::Round($sh.Width, 1); h = [math]::Round($sh.Height, 1) }
            try { if ($sh.HasTextFrame -ne 0 -and $sh.TextFrame.HasText -ne 0) { $e.text = $sh.TextFrame.TextRange.Text } } catch {}
            try { if ($sh.HasTextFrame -ne 0 -and $sh.TextFrame.HasText -ne 0) { $e.bullet = [int]$sh.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Bullet.Type } } catch {}
            try { if ($sh.HasTextFrame -ne 0 -and $sh.TextFrame.HasText -ne 0) { $e.link = [string]$sh.TextFrame.TextRange.Runs(1).ActionSettings(1).Hyperlink.Address } } catch {}
            try { if ($sh.HasTable -ne 0) { $tb = $sh.Table; $cells = @(); for ($r = 1; $r -le $tb.Rows.Count; $r++) { for ($c = 1; $c -le $tb.Columns.Count; $c++) { $cells += $tb.Cell($r, $c).Shape.TextFrame.TextRange.Text } }; $e.table = [ordered]@{ rows = $tb.Rows.Count; cols = $tb.Columns.Count; cells = $cells } } } catch {}
            try { if ($sh.Type -eq 6) { $e.group = $sh.GroupItems.Count } } catch {}
            $sl.shapes += $e
        }
        $o.slides += $sl
    }
    return $o
}

try {
    $ctx = Start-Ppt -TimeoutSec $TimeoutSec
    foreach ($f in $Files) {
        $full = [System.IO.Path]::GetFullPath($f)
        $r = [ordered]@{ path = $full; ok = $false; error = ''; killed = $false }
        Ping-Ppt $ctx
        $pres = $null
        try {
            $pres = $ctx.App.Presentations.Open($full, -1, 0, 0)   # ReadOnly, Untitled=false, WithWindow=msoFalse
            Ping-Ppt $ctx
            $r.ok = $true
            $r.content = Read-Pres $pres $full
            # rilevatore di ripristino: PowerPoint non espone un flag; un file riparato apre ma segnala ProtectedView/riparazione tramite finestre
            $r.windows = $ctx.App.Windows.Count
            $r.protectedView = $ctx.App.ProtectedViewWindows.Count
            if ($PdfDir) { New-Item -ItemType Directory -Force $PdfDir | Out-Null; $pdf = Join-Path $PdfDir ([IO.Path]::GetFileNameWithoutExtension($full) + '.pdf'); if (Test-Path $pdf) { Remove-Item $pdf }; $pres.SaveCopyAs($pdf, 32); $r.pdf = $pdf }
            if ($PngDir) { $d = Join-Path $PngDir ([IO.Path]::GetFileNameWithoutExtension($full)); New-Item -ItemType Directory -Force $d | Out-Null; foreach ($s in $pres.Slides) { $s.Export((Join-Path $d ('slide{0:00}.png' -f $s.SlideIndex)), 'PNG', 1280, [int](1280 * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)) } }
        } catch {
            $r.error = $_.Exception.Message
        }
        try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch {}
        if ($ctx.Owned -and -not (Test-PptAlive $ctx)) {
            # watchdog ha ucciso l'istanza: dialogo modale o blocco
            $r.killed = $true; $r.ok = $false; if (-not $r.error) { $r.error = 'istanza bloccata (dialogo modale?) uccisa dal watchdog' }
            try { Remove-Item ($ctx.Heartbeat + '.killed') -ErrorAction SilentlyContinue } catch {}
            Stop-Ppt $ctx
            $ctx = Start-Ppt -TimeoutSec $TimeoutSec
        }
        $results += [pscustomobject]$r
    }
} finally {
    if ($ctx) { Stop-Ppt $ctx }
    $json = ConvertTo-Json -InputObject @($results) -Depth 12
    [System.IO.File]::WriteAllText($OutJson, $json, (New-Object System.Text.UTF8Encoding($false)))
}
