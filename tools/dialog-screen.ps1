# Cattura dallo SCHERMO VERO (CopyFromScreen, non resa interna) la scheda della finestra di salvataggio, a scala Windows reale.
# Uso: dialog-screen.ps1 -Out file.png [-Exe percorso.exe]   (l'app viene avviata e chiusa per PID; la finestra e' topmost)
param([string]$Out, [string]$Exe = '', [string]$File = '', [switch]$Light)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class DS { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern uint GetDpiForSystem(); [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int a, out RECT r, int s); [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r); [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; } }
'@
[void][DS]::SetProcessDPIAware()
$repo = Split-Path $PSScriptRoot -Parent
if (-not $Exe) { $Exe = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe' }
if (-not $File) { $File = Join-Path $repo 'tests\corpus\office-made\title-slide.pptx' }
$env:PP_SETTINGS = Join-Path $env:TEMP 'pp-dialog-screen.json'; Remove-Item $env:PP_SETTINGS -ErrorAction SilentlyContinue; if ($Light) { Set-Content $env:PP_SETTINGS '{"Dark":false}' }
$env:PP_DEFAULT_FOLDER = 'C:\Documenti'
$p = Start-Process $Exe -ArgumentList $File, '--screenshot', (Join-Path $env:TEMP 'pp-unused.png'), 'dialog-live' -PassThru
try {
    $AE = [System.Windows.Automation.AutomationElement]; $dlg = $null
    for ($i = 0; $i -lt 40 -and -not $dlg; $i++) {
        Start-Sleep -Milliseconds 300
        foreach ($w in $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) { if ($w.Current.ProcessId -eq $p.Id -and $w.Current.ControlType -eq [System.Windows.Automation.ControlType]::Window -and $w.Current.Name -like 'Vuoi salvare*') { $dlg = $w } }
    }
    if (-not $dlg) { throw 'finestra di salvataggio non trovata' }
    Start-Sleep -Milliseconds 800
    $r = $dlg.Current.BoundingRectangle; $dpi = [DS]::GetDpiForSystem()
    $fr = New-Object DS+RECT; [void][DS]::DwmGetWindowAttribute([IntPtr]$dlg.Current.NativeWindowHandle, 9, [ref]$fr, 16)   # DWMWA_EXTENDED_FRAME_BOUNDS: riquadro visibile esatto, senza ombra
    $x = $fr.L; $y = $fr.T; $w = $fr.R - $fr.L; $h = $fr.B - $fr.T
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($x, $y, 0, 0, $bmp.Size); $g.Dispose(); $bmp.Save($Out); $bmp.Dispose()
    "dpi=$dpi scala=$([math]::Round($dpi / 96 * 100))% rect=$($r.Width)x$($r.Height) out=$Out"
} finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
