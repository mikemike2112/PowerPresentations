# Prova end-to-end della UI reale con UI Automation + input reale: apre un file del corpus (copia), alterna la scheda Transizioni,
# applica una transizione, duplica una slide, riordina con drag & drop reale delle miniature, salva con Ctrl+S,
# avvia la presentazione a schermo intero (F5), cattura SOLO le finestre dell'app e verifica il file salvato.
# Uso: ui-smoke.ps1 [-Shots dir]
param([string]$Shots = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms, System.Drawing, System.IO.Compression.FileSystem
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class W { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y); [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, int f, int e); [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint f); [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
 [DllImport("user32.dll")] public static extern void mouse_event(int f, int dx, int dy, int d, int e); [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
 [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; } }
'@
[void][W]::SetProcessDPIAware()
$repo = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
$work = Join-Path $env:TEMP 'pp-ui-smoke'; New-Item -ItemType Directory -Force $work | Out-Null
$file = Join-Path $work 'smoke.pptx'; Copy-Item (Join-Path $repo 'tests\corpus\office-made\mixed-realistic-1.pptx') $file -Force
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
function Find($root, $name) { $c = New-Object System.Windows.Automation.PropertyCondition ($AE::NameProperty), $name; $root.FindFirst($TS::Descendants, $c) }
function Click($e) { $r = $e.Current.BoundingRectangle; [void][W]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds 250; [W]::mouse_event(2, 0, 0, 0, 0); [W]::mouse_event(4, 0, 0, 0, 0); Start-Sleep -Milliseconds 700 }
function Press($root, $name) {
    $e = Find $root $name; if (-not $e) { throw "controllo non trovato: $name" }
    $o = $null
    if ($e.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$o)) { $o.Invoke() }
    elseif ($e.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$o)) { $o.Select() }
    elseif ($e.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$o)) { for ($t = 0; $t -lt 4 -and $o.Current.ToggleState -ne 'On'; $t++) { Front $h; Start-Sleep -Milliseconds 300; Click $e } }   # clic reale, ripetuto finche' lo stato e' On (condizione, non attesa fissa)
    else { throw "nessun pattern per $name" }
    Start-Sleep -Milliseconds 600
}
function Front($h) { if ([W]::GetForegroundWindow() -eq $h) { return }; for ($i = 0; $i -lt 6; $i++) { [void][W]::SetWindowPos($h, [IntPtr]::new(-1), 0, 0, 0, 0, 0x43); [W]::keybd_event(0x12, 0, 0, 0); [W]::keybd_event(0x12, 0, 2, 0); [void][W]::SetForegroundWindow($h); Start-Sleep -Milliseconds 300; if ([W]::GetForegroundWindow() -eq $h) { return } } }   # topmost + foreground (trucco ALT contro il blocco del focus); i clic reali non finiscono su altre finestre
function Shot($h, $path) {
    $r = New-Object W+RECT; [void][W]::GetWindowRect($h, [ref]$r)
    $bmp = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T); $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][W]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose(); $bmp.Save($path); $bmp.Dispose()   # PrintWindow: cattura la SOLA finestra anche se coperta
}
$p = Start-Process $exe -ArgumentList $file -PassThru
$result = [ordered]@{}
try {
    $deadline = (Get-Date).AddSeconds(20); while (-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) { Start-Sleep -Milliseconds 300; $p.Refresh(); if ((Get-Date) -gt $deadline) { throw 'finestra non apparsa' } }
    Start-Sleep -Seconds 2
    $h = $p.MainWindowHandle; Front $h
    $root = $AE::FromHandle($h)
    Press $root 'Transizioni'; Press $root 'Dissolvenza'
    Press $root 'Home'; Press $root 'Duplica'
    # drag & drop reale: prima miniatura -> terza
    $items = $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem)))
    $result.thumbnails = $items.Count
    for ($try = 0; $try -lt 4; $try++) {
        $items = $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem)))
        if ($items[1].Current.Name -like '*Obiettivi*') { break }     # condizione, non attesa fissa: riordinato
        $a = $items[0].Current.BoundingRectangle; $b = $items[2].Current.BoundingRectangle
        Front $h
        [void][W]::SetCursorPos([int]($a.X + $a.Width / 2), [int]($a.Y + $a.Height / 2)); Start-Sleep -Milliseconds 300
        [W]::mouse_event(2, 0, 0, 0, 0); Start-Sleep -Milliseconds 300
        for ($i = 1; $i -le 30; $i++) { [void][W]::SetCursorPos([int]($a.X + $a.Width / 2), [int]($a.Y + $a.Height / 2 + ($b.Y - $a.Y) * $i / 30)); Start-Sleep -Milliseconds 50 }
        [void][W]::SetCursorPos([int]($b.X + $b.Width / 2), [int]($b.Y + $b.Height / 2)); Start-Sleep -Milliseconds 400
        [W]::mouse_event(4, 0, 0, 0, 0); Start-Sleep -Seconds 1
    }
    $result.dragTries = $try; $result.names = (($root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem))) | % { $_.Current.Name }) -join "/")
    if ($Shots) { New-Item -ItemType Directory -Force $Shots | Out-Null }  # la finestra principale non si cattura dallo schermo (ombre/bordi includono altre finestre): vedi screenshot.ps1
    Front $h; Click (Find $root 'Salva'); Start-Sleep -Seconds 2   # clic reale sul pulsante Salva della barra del titolo
    # presentazione a schermo intero
    $show = $null
    for ($t = 0; $t -lt 4 -and -not $show; $t++) {
        Front $h; Click (Find $root 'Avvia presentazione'); Start-Sleep -Seconds 2
        $wins = $AE::RootElement.FindAll($TS::Children, (New-Object System.Windows.Automation.PropertyCondition ($AE::ProcessIdProperty), $p.Id))
        foreach ($w in $wins) { if ($w.Current.Name -eq 'Presentazione') { $show = $w } }
    }
    $result.showWindow = [bool]$show
    if ($show) {
        $br = $show.Current.BoundingRectangle; $result.showSize = "$([int]$br.Width)x$([int]$br.Height)"
        if ($Shots) { Shot ([IntPtr]$show.Current.NativeWindowHandle) (Join-Path $Shots 'app-schermo-intero-reale.png') }
        Front ([IntPtr]$show.Current.NativeWindowHandle); [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}'); Start-Sleep -Seconds 2
        if ($Shots) { Shot ([IntPtr]$show.Current.NativeWindowHandle) (Join-Path $Shots 'app-schermo-intero-reale-2.png') }
        Front ([IntPtr]$show.Current.NativeWindowHandle); [System.Windows.Forms.SendKeys]::SendWait('{ESC}'); Start-Sleep -Seconds 1; Front $h
    }
    $o = $null; if ($root.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$o)) { $o.Close() }
    if (-not $p.WaitForExit(10000)) { $result.closed = $false } else { $result.closed = $true }
} finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
# verifica del file salvato
$z = [System.IO.Compression.ZipFile]::OpenRead($file)
try {
    $rd = { param($n) $s = New-Object System.IO.StreamReader($z.GetEntry($n).Open()); $t = $s.ReadToEnd(); $s.Close(); $t }
    $pres = & $rd 'ppt/presentation.xml'; $rels = & $rd 'ppt/_rels/presentation.xml.rels'
    $order = [regex]::Matches($pres, '<p:sldId [^>]*r:id="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
    $titles = @(); $trans = @()
    foreach ($rid in $order) {
        $t = [regex]::Match($rels, "Id=`"$rid`"[^>]*Target=`"([^`"]+)`"|Target=`"([^`"]+)`"[^>]*Id=`"$rid`"")
        $target = if ($t.Groups[1].Value) { $t.Groups[1].Value } else { $t.Groups[2].Value }
        $x = & $rd ('ppt/' + $target.TrimStart('/').Replace('ppt/', ''))
        $titles += [regex]::Match($x, '<a:t>([^<]*)</a:t>').Groups[1].Value; $trans += [bool]($x -match '<p:fade')
    }
    $result.savedSlideCount = $order.Count; $result.savedTitles = ($titles -join ' | '); $result.savedFade = ($trans -join ',')
} finally { $z.Dispose() }
$result | ConvertTo-Json -Compress
