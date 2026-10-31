# Prova end-to-end della UI reale con UI Automation + input reale: apre un file del corpus (copia), alterna la scheda Transizioni,
# applica una transizione, duplica una slide, riordina con drag & drop reale delle miniature, salva con Ctrl+S,
# avvia la presentazione a schermo intero (F5), cattura SOLO le finestre dell'app e verifica il file salvato.
# Uso: ui-smoke.ps1 [-Shots dir]
param([string]$Shots = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms, System.Drawing, System.IO.Compression.FileSystem
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class W { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
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
function Press($root, $name) {
    $e = Find $root $name; if (-not $e) { throw "controllo non trovato: $name" }
    $o = $null
    if ($e.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$o)) { $o.Invoke() }
    elseif ($e.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$o)) { $o.Select() }
    elseif ($e.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$o)) { $r = $e.Current.BoundingRectangle; [void][W]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds 200; [W]::mouse_event(2, 0, 0, 0, 0); [W]::mouse_event(4, 0, 0, 0, 0) }  # clic reale: Toggle() di UIA non genera Click
    else { throw "nessun pattern per $name" }
    Start-Sleep -Milliseconds 600
}
function Shot($h, $path) {
    $r = New-Object W+RECT; [void][W]::GetWindowRect($h, [ref]$r)
    $bmp = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T)
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($r.L, $r.T, 0, 0, $bmp.Size); $g.Dispose(); $bmp.Save($path); $bmp.Dispose()
}
$p = Start-Process $exe -ArgumentList $file -PassThru
$result = [ordered]@{}
try {
    $deadline = (Get-Date).AddSeconds(20); while (-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) { Start-Sleep -Milliseconds 300; $p.Refresh(); if ((Get-Date) -gt $deadline) { throw 'finestra non apparsa' } }
    Start-Sleep -Seconds 2
    $h = $p.MainWindowHandle; [void][W]::SetForegroundWindow($h)
    $root = $AE::FromHandle($h)
    Press $root 'Transizioni'; Press $root 'Dissolvenza'
    Press $root 'Home'; Press $root 'Duplica'
    # drag & drop reale: prima miniatura -> terza
    $items = $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem)))
    $result.thumbnails = $items.Count
    $a = $items[0].Current.BoundingRectangle; $b = $items[2].Current.BoundingRectangle
    [void][W]::SetForegroundWindow($h); Start-Sleep -Milliseconds 300
    [W]::SetCursorPos([int]($a.X + $a.Width / 2), [int]($a.Y + $a.Height / 2)); Start-Sleep -Milliseconds 200
    [W]::mouse_event(2, 0, 0, 0, 0); Start-Sleep -Milliseconds 200
    for ($i = 1; $i -le 20; $i++) { [W]::SetCursorPos([int]($a.X + $a.Width / 2), [int]($a.Y + $a.Height / 2 + ($b.Y - $a.Y) * $i / 20)); Start-Sleep -Milliseconds 30 }
    Start-Sleep -Milliseconds 300; [W]::mouse_event(4, 0, 0, 0, 0); Start-Sleep -Seconds 1
    if ($Shots) { New-Item -ItemType Directory -Force $Shots | Out-Null }  # la finestra principale non si cattura dallo schermo (ombre/bordi includono altre finestre): vedi screenshot.ps1
    [void][W]::SetForegroundWindow($h); [System.Windows.Forms.SendKeys]::SendWait('^s'); Start-Sleep -Seconds 2
    # presentazione a schermo intero
    [System.Windows.Forms.SendKeys]::SendWait('{F5}'); Start-Sleep -Seconds 2
    $wins = $AE::RootElement.FindAll($TS::Children, (New-Object System.Windows.Automation.PropertyCondition ($AE::ProcessIdProperty), $p.Id))
    $show = $null; foreach ($w in $wins) { if ($w.Current.Name -eq 'Presentazione') { $show = $w } }
    $result.showWindow = [bool]$show
    if ($show) {
        $br = $show.Current.BoundingRectangle; $result.showSize = "$([int]$br.Width)x$([int]$br.Height)"
        if ($Shots) { Shot ([IntPtr]$show.Current.NativeWindowHandle) (Join-Path $Shots 'app-schermo-intero-reale.png') }
        [System.Windows.Forms.SendKeys]::SendWait('{RIGHT}'); Start-Sleep -Seconds 2
        if ($Shots) { Shot ([IntPtr]$show.Current.NativeWindowHandle) (Join-Path $Shots 'app-schermo-intero-reale-2.png') }
        [System.Windows.Forms.SendKeys]::SendWait('{ESC}'); Start-Sleep -Seconds 1
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
