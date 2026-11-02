# Prova end-to-end dei comandi del ribbon con input reale (tastiera) e UI Automation: Ctrl+A/C/V/Z, Trova e sostituisci,
# Seleziona, Elenco puntato, Collegamento, Griglia/Righello, Applica a tutte. Verifica poi i file salvati dal test C#.
# Uso: ui-commands.ps1 -Find testo -Repl nuovo
param([string]$Find = 'a', [string]$Repl = 'b', [string]$Work = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class W { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y); [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, int f, int e); [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint f);
 [DllImport("user32.dll")] public static extern void mouse_event(int f, int dx, int dy, int d, int e); }
'@
[void][W]::SetProcessDPIAware()
$repo = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
if (-not $Work) { $Work = Join-Path $env:TEMP 'pp-ui-commands' }
New-Item -ItemType Directory -Force $Work | Out-Null
$file = Join-Path $Work 'cmd.pptx'; Copy-Item (Join-Path $repo 'tests\corpus\office-made\mixed-realistic-1.pptx') $file -Force
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

$p = Start-Process $exe -ArgumentList "`"$file`"" -PassThru
$res = [ordered]@{}
try {
    $deadline = (Get-Date).AddSeconds(25); while (-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) { Start-Sleep -Milliseconds 300; $p.Refresh(); if ((Get-Date) -gt $deadline) { throw 'finestra non apparsa' } }
    Start-Sleep -Seconds 2; $h = $p.MainWindowHandle; Front $h; $root = $AE::FromHandle($h)
    # area della slide: sopra le note
    $notes = Find $root 'Note del relatore'; $nr = $notes.Current.BoundingRectangle
    Front $h; [void][W]::SetCursorPos([int]($nr.X + 40), [int]($nr.Y - 40)); Start-Sleep -Milliseconds 200; [W]::mouse_event(2, 0, 0, 0, 0); [W]::mouse_event(4, 0, 0, 0, 0); Start-Sleep -Milliseconds 500
    # Ctrl+A / Ctrl+C / Ctrl+V / Ctrl+Z
    Keys '^a'; $res.afterSelectAll = Status
    Keys '^c'; $res.afterCopy = Status
    Keys '^v'; Save 'pasted.pptx'; $res.afterPaste = Status
    Keys '^z'; Save 'undone.pptx'
    Keys '^a'; Keys '^x'; Save 'cut.pptx'; Keys '^v'; Save 'cutpasted.pptx'
    Keys '^z'; Keys '^z'
    # Trova e sostituisci
    Press $root 'Home'; Press $root 'Trova e sostituisci'
    $dlg = InDialog 'Sostituisci tutto'; $w = $dlg[0]
    SetText (FindEdit $w 'Trova') $Find; SetText (FindEdit $w 'Sostituisci con') $Repl
    $b = $dlg[1].GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern); $b.Invoke(); Start-Sleep -Milliseconds 900
    $res.afterReplace = Status; Save 'replaced.pptx'
    # Seleziona -> prima forma, elenco puntato, collegamento
    Press $root 'Seleziona'
    $tAll = $null; for ($t = 0; $t -lt 20 -and -not $tAll; $t++) { $tAll = Find $root 'Seleziona tutto'; if (-not $tAll) { Start-Sleep -Milliseconds 300 } }   # la finestra della scelta e' figlia della principale in UIA
    if (-not $tAll) { throw 'elenco Seleziona non apparso' }
    $items = $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::ListItem)), (New-Object System.Windows.Automation.PropertyCondition ($AE::IsOffscreenProperty), $false))))
    $items = @($items | ? { $_.Current.Name -notlike 'Diapositiva *' })
    $res.pickItems = $items.Count
    $items[1].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep -Milliseconds 400
    [W]::keybd_event(0x0D, 0, 0, 0); [W]::keybd_event(0x0D, 0, 2, 0); Start-Sleep -Milliseconds 900; $res.afterPick = Status
    Press $root 'Elenco puntato'; $res.afterBullet = Status
    Press $root 'Inserisci'; Press $root 'Collegamento'
    $dlg = InDialog 'OK'; SetText (FindEdit $dlg[0] 'Collegamento') 'https://example.org/ui'
    $dlg[1].GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Start-Sleep -Milliseconds 900
    Save 'shape.pptx'
    # Griglia e Righello (scheda Visualizza)
    Press $root 'Visualizza'
    foreach ($n in 'Griglia', 'Righello') {
        $c = $null; for ($t = 0; $t -lt 10 -and -not $c; $t++) { $c = Find $root $n; if (-not $c) { Start-Sleep -Milliseconds 300 } }
        $res["$n-enabled"] = $c.Current.IsEnabled
        $tg = $c.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); $tg.Toggle(); Start-Sleep -Milliseconds 500
        $res["$n-on"] = ($tg.Current.ToggleState -eq 'On')
    }
    # Applica a tutte (scheda Transizioni)
    Press $root 'Transizioni'; Press $root 'Dissolvenza'; Press $root 'Applica a tutte'; Save 'trans.pptx'
    # inerti rimasti
    Press $root 'Inserisci'
    $res.iconsEnabled = (Find $root 'Icone').Current.IsEnabled
    $res.videoEnabled = (Find $root 'Video').Current.IsEnabled
} catch { $res.error = $_.Exception.Message + ' @' + $_.InvocationInfo.ScriptLineNumber } finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
$res | ConvertTo-Json -Compress
