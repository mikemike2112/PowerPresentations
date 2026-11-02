# (riquadro Commenti) Prova end-to-end con input reale: apre il riquadro dal comando Commento, modifica ed elimina commenti con clic e tastiera reali, Ctrl+Z, cambio slide; cattura SOLO la finestra dell'app.
# Prova end-to-end dei comandi del ribbon con input reale (tastiera) e UI Automation: Ctrl+A/C/V/Z, Trova e sostituisci,
# Seleziona, Elenco puntato, Collegamento, Griglia/Righello, Applica a tutte. Verifica poi i file salvati dal test C#.
# Uso: ui-commands.ps1 -Find testo -Repl nuovo
param([string]$Src = '', [string]$Work = '', [string]$Shots = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class W { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y); [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, int f, int e); [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint f);
 [DllImport("user32.dll")] public static extern void mouse_event(int f, int dx, int dy, int d, int e); [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r); [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
 [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; } }
'@
[void][W]::SetProcessDPIAware()
$repo = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
if (-not $Work) { $Work = Join-Path $env:TEMP 'pp-ui-commands' }
New-Item -ItemType Directory -Force $Work | Out-Null
$file = Join-Path $Work 'cmt.pptx'; Copy-Item $Src $file -Force
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
function Find($root, $name) { $c = New-Object System.Windows.Automation.PropertyCondition ($AE::NameProperty), $name; $root.FindFirst($TS::Descendants, $c) }
function Front($h) { if ([W]::GetForegroundWindow() -eq $h) { return }; for ($i = 0; $i -lt 6; $i++) { [void][W]::SetWindowPos($h, [IntPtr]::new(-1), 0, 0, 0, 0, 0x43); [W]::keybd_event(0x12, 0, 0, 0); [W]::keybd_event(0x12, 0, 2, 0); [void][W]::SetForegroundWindow($h); Start-Sleep -Milliseconds 300; if ([W]::GetForegroundWindow() -eq $h) { return } } }
function Click($e) { $r = $e.Current.BoundingRectangle; [void][W]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds 250; [W]::mouse_event(2, 0, 0, 0, 0); [W]::mouse_event(4, 0, 0, 0, 0); Start-Sleep -Milliseconds 600 }
function Press($root, $name) {
    $e = $null; for ($t = 0; $t -lt 10 -and -not $e; $t++) { $e = Find $root $name; if (-not $e) { Start-Sleep -Milliseconds 300 } }
    if (-not $e) { throw "controllo non trovato: $name" }
    $o = $null
    if ($e.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$o)) { $o.Invoke() }
    elseif ($e.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$o)) { $o.Select() }
    elseif ($e.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$o)) { Front $script:h; Click $e }
    else { throw "nessun pattern per $name" }
    Start-Sleep -Milliseconds 700
}
function Status { $r = $script:root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::Text))); foreach ($x in $r) { if ($x.Current.Name -like 'Diapositiva * di *' -or $x.Current.Name -like 'Forma *' -or $x.Current.Name -like '*copiat*' -or $x.Current.Name -like '*sostituzion*') { return $x.Current.Name } }; return '' }
function InDialog($name, $timeout = 8) {
    $end = (Get-Date).AddSeconds($timeout)
    while ((Get-Date) -lt $end) {
        foreach ($w in $AE::RootElement.FindAll($TS::Children, [System.Windows.Automation.Condition]::TrueCondition)) { if ($w.Current.ProcessId -eq $script:p.Id) { $e = Find $w $name; if ($e) { return @($w, $e) } } }
        Start-Sleep -Milliseconds 300
    }
    throw "finestra/controllo non trovato: $name"
}
function FindEdit($root, $name) { $c = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition ($AE::NameProperty), $name), (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::Edit))); $root.FindFirst($TS::Descendants, $c) }
function SetText($e, $v) { $o = $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern); $o.SetValue($v) }
function Keys($k) { Front $script:h; [System.Windows.Forms.SendKeys]::SendWait($k); Start-Sleep -Milliseconds 900 }
function Save($name) { Keys '^s'; Start-Sleep -Milliseconds 500; Copy-Item $file (Join-Path $Work $name) -Force }

function Shot($h, $path) {
    $r = New-Object W+RECT; [void][W]::GetWindowRect($h, [ref]$r)
    $bmp = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T); $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][W]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose(); $bmp.Save($path); $bmp.Dispose()   # PrintWindow: la SOLA finestra, mai lo schermo intero
}
function Btns { @($script:root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::Button))) | ? { $_.Current.Name -like 'Elimina commento *' }).Count }
function PaneTexts { (@($script:root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::Text))) | % { $_.Current.Name } | ? { $_ }) -join ' | ') }   # tutti i testi della finestra: il riquadro e' l'unico a contenere i commenti
function ClickBtn($name) { $e = $null; for ($t = 0; $t -lt 10 -and -not $e; $t++) { $e = Find $script:root $name; if (-not $e) { Start-Sleep -Milliseconds 300 } }; if (-not $e) { throw "controllo non trovato: $name" }; Front $script:h; Click $e }
function ClickInDialog($name) { $d = InDialog $name; Click $d[1] }
function Thumbs { @($script:root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem)))) }
$p = Start-Process $exe -ArgumentList "`"$file`"" -PassThru
$res = [ordered]@{}
try {
    $deadline = (Get-Date).AddSeconds(25); while (-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) { Start-Sleep -Milliseconds 300; $p.Refresh(); if ((Get-Date) -gt $deadline) { throw 'finestra non apparsa' } }
    Start-Sleep -Seconds 2; $h = $p.MainWindowHandle; Front $h; $root = $AE::FromHandle($h)
    $res.paneBefore = [bool](Find $root 'Elenco commenti' | ? { $_.Current.IsOffscreen -eq $false })
    Press $root 'Inserisci'; Press $root 'Commento'
    for ($t = 0; $t -lt 10 -and (Btns) -eq 0; $t++) { Start-Sleep -Milliseconds 300 }
    $res.listedButtons = Btns; $res.listedTexts = PaneTexts
    if ($Shots) { Shot $h (Join-Path $Shots 'commenti-riquadro.png') }
    # modifica del 3o elemento (Secondo)
    ClickBtn 'Modifica commento 3'; $d = InDialog 'Salva'
    Start-Sleep -Milliseconds 500; [System.Windows.Forms.SendKeys]::SendWait('Secondo modificato'); Start-Sleep -Milliseconds 400
    Click $d[1]; Start-Sleep -Milliseconds 600
    $res.afterEditTexts = PaneTexts; Save 'edited.pptx'
    # elimina il 1o (ha una risposta): conferma
    ClickBtn 'Elimina commento 1'; $c = InDialog 'Elimina'; $res.confirmShown = $true
    Click $c[1]; Start-Sleep -Milliseconds 600
    $res.afterDeleteButtons = Btns; $res.afterDeleteTexts = PaneTexts; Save 'deleted.pptx'
    if ($Shots) { Shot $h (Join-Path $Shots 'commenti-dopo-elimina.png') }
    # un solo Ctrl+Z riporta commento e risposta
    Keys '^z'; $res.afterUndoButtons = Btns; $res.afterUndoTexts = PaneTexts; Save 'undone.pptx'
    # cambio slide: il riquadro si aggiorna
    $th = Thumbs; Front $h; Click $th[1]; $res.otherSlideButtons = Btns; $res.otherSlideTexts = PaneTexts
    if ($Shots) { Shot $h (Join-Path $Shots 'commenti-altra-slide.png') }
    $th = Thumbs; Front $h; Click $th[0]; $res.backButtons = Btns
} catch { $res.error = $_.Exception.Message + ' @' + $_.InvocationInfo.ScriptLineNumber } finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
$res | ConvertTo-Json -Compress
