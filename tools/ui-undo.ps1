# (variante: annulla/ripristina e scheda File) Prova end-to-end della UI reale con UI Automation + input reale: apre un file del corpus (copia), alterna la scheda Transizioni,
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
$res = [ordered]@{}
try {
    $deadline = (Get-Date).AddSeconds(20); while (-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) { Start-Sleep -Milliseconds 300; $p.Refresh(); if ((Get-Date) -gt $deadline) { throw 'finestra non apparsa' } }
    Start-Sleep -Seconds 2; $h = $p.MainWindowHandle; Front $h; $root = $AE::FromHandle($h)
    function Items { $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem))).Count }
    function Until($want, $name) { for ($t = 0; $t -lt 5 -and (Items) -ne $want; $t++) { Front $h; Click (Find $root $name); Start-Sleep -Milliseconds 500 } ; Items }
    $res.start = Items
    Press $root 'Duplica'; $res.afterDuplicate = Items
    $res.afterUndo = Until 3 'Annulla'
    $res.afterRedo = Until 4 'Ripristina azione'
    # scheda File a tutta pagina
    for ($t = 0; $t -lt 4 -and -not (Find $root 'Esporta come PDF'); $t++) { Front $h; Click (Find $root 'File') }; $res.backstage = [bool](Find $root 'Esporta come PDF')
    Click (Find $root 'Indietro'); $res.backstageClosed = -not ((Find $root 'Esporta come PDF') -and (Find $root 'Esporta come PDF').Current.IsOffscreen -eq $false)
    # Ctrl+Z da tastiera con il focus sulle miniature
    Front $h; Click (Find $root 'Annulla'); $res.afterUndoBtn2 = Items; Click (Find $root 'Ripristina azione'); $items = $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem))); Click $items[0]; $f = $AE::FocusedElement; $res.focus = "$($f.Current.ControlType.ProgrammaticName) $($f.Current.Name)"; [System.Windows.Forms.SendKeys]::SendWait('^z'); Start-Sleep -Milliseconds 800; $res.afterCtrlZ = Items
    for ($t = 0; $t -lt 4 -and (Items) -ne 4; $t++) { Front $h; [System.Windows.Forms.SendKeys]::SendWait('^y'); Start-Sleep -Milliseconds 800 }; $res.afterCtrlY = Items; $res.title = $root.Current.Name
    Front $h; Click (Find $root 'Chiudi')   # clic reale sul pulsante Chiudi della barra del titolo
    # documento modificato: appare la richiesta di salvataggio; si risponde No
    Start-Sleep -Seconds 1
    $d = $AE::RootElement.FindAll($TS::Children, (New-Object System.Windows.Automation.PropertyCondition ($AE::ProcessIdProperty), $p.Id))
    $no = $null
    for ($t = 0; $t -lt 10 -and -not $no; $t++) {
        if ($t -gt 0 -and -not (Find $root 'Non salvare')) { $res.closeRetry = $t; Front $h; Click (Find $root 'Chiudi') }; $no = Find $root 'Non salvare'; if (-not $no) { foreach ($w in $AE::RootElement.FindAll($TS::Children, [System.Windows.Automation.Condition]::TrueCondition)) { if ($w.Current.ProcessId -eq $p.Id) { $no = Find $w 'Non salvare'; if ($no) { break } } } }
        Start-Sleep -Milliseconds 400
    }
    $res.dialogShown = [bool]$no; $res.top = (($AE::RootElement.FindAll($TS::Children, [System.Windows.Automation.Condition]::TrueCondition) | ? { $_.Current.ProcessId -eq $p.Id } | % { $_.Current.Name + '/' + $_.Current.ClassName }) -join ';'); $res.rootChildren = (($root.FindAll($TS::Children, [System.Windows.Automation.Condition]::TrueCondition) | % { $_.Current.ControlType.ProgrammaticName + ':' + $_.Current.Name }) -join ';')
    for ($t = 0; $no -and $t -lt 4 -and -not $p.HasExited; $t++) { Click $no; [void]$p.WaitForExit(2000) }   # clic reale su Non salvare (nuova richiesta di salvataggio)
    $res.closed = $p.WaitForExit(10000)
} finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
$res | ConvertTo-Json -Compress
