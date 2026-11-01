# Prova con UI reale della scheda Powers: clic REALE sull'interruttore "Modalità scura", verifica del file delle impostazioni,
# riavvio e controllo del tema all'avvio (colore dello sfondo in uno screenshot dell'app), file impostazioni corrotto senza crash.
param([string]$Shots = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class WP { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h); [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y); [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, int f, int e); [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint f); [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f); [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r); [DllImport("user32.dll")] public static extern void mouse_event(int f, int dx, int dy, int d, int e); [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; } }
'@
[void][WP]::SetProcessDPIAware()
$repo = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
$file = Join-Path $repo 'tests\corpus\office-made\mixed-realistic-1.pptx'
$set = Join-Path $env:TEMP 'pp-powers-settings.json'; Remove-Item $set -ErrorAction SilentlyContinue
$env:PP_SETTINGS = $set
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
function Find($root, $name) { $root.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::NameProperty), $name)) }
function Click($e) { $r = $e.Current.BoundingRectangle; [void][WP]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds 250; [WP]::mouse_event(2, 0, 0, 0, 0); [WP]::mouse_event(4, 0, 0, 0, 0); Start-Sleep -Milliseconds 700 }
function Front($h) { for ($i = 0; $i -lt 6; $i++) { [void][WP]::SetWindowPos($h, [IntPtr]::new(-1), 0, 0, 0, 0, 0x43); [WP]::keybd_event(0x12, 0, 0, 0); [WP]::keybd_event(0x12, 0, 2, 0); [void][WP]::SetForegroundWindow($h); Start-Sleep -Milliseconds 300; if ([WP]::GetForegroundWindow() -eq $h) { return } } }   # topmost + foreground (trucco ALT contro il blocco del focus); i clic reali non finiscono su altre finestre
function Shot($h, $path) {
    $r = New-Object WP+RECT; [void][WP]::GetWindowRect($h, [ref]$r)
    $bmp = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T); $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][WP]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose(); $bmp.Save($path); $bmp.Dispose()
}
function Start-App($more) { $p = Start-Process $exe -ArgumentList (@($file) + $more) -PassThru; $d = (Get-Date).AddSeconds(20); while ((-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) -and (Get-Date) -lt $d) { Start-Sleep -Milliseconds 300; $p.Refresh() }; Start-Sleep -Seconds 2; $p }
function Lum($png) { $b = [System.Drawing.Bitmap]::FromFile($png); $c = $b.GetPixel(5, [int]($b.Height / 2)); $b.Dispose(); [int](($c.R + $c.G + $c.B) / 3) }
$res = [ordered]@{}
$p = Start-App @()
try {
    $h = $p.MainWindowHandle; Front $h; $root = $AE::FromHandle($h)
    $res.defaultDarkInFile = (-not (Test-Path $set))
    $sw = $null
    for ($i = 0; $i -lt 6 -and -not $sw; $i++) {
        Front $h; Click (Find $root 'Powers'); Start-Sleep -Milliseconds 800
        $c1 = [System.Windows.Automation.PropertyCondition]::new($AE::NameProperty, 'Modalità scura'); $c2 = [System.Windows.Automation.PropertyCondition]::new($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
        $sw = $root.FindFirst($TS::Descendants, [System.Windows.Automation.AndCondition]::new($c1, $c2))
        if ($sw -and $sw.Current.IsOffscreen) { $sw = $null }
    }
    if ($Shots) { New-Item -ItemType Directory -Force $Shots | Out-Null; Shot $h (Join-Path $Shots 'powers-scuro.png') }
    Front $h; Click $sw                      # clic reale sull'interruttore
    Start-Sleep -Milliseconds 800
    if ($Shots) { Shot $h (Join-Path $Shots 'powers-chiaro.png') }
    $res.settingsAfterToggle = (Get-Content $set -Raw).Trim()
} finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
# riavvio: deve partire chiaro
$png = Join-Path $env:TEMP 'pp-powers-restart.png'
$q = Start-Process $exe -ArgumentList $file, '--screenshot', $png, 'main' -PassThru; [void]$q.WaitForExit(40000); if (-not $q.HasExited) { Stop-Process -Id $q.Id -Force }
$res.restartLuminance = Lum $png
# file corrotto: nessun crash, parte scuro
Set-Content $set '{ questo non è json <<<' -Encoding UTF8
$png2 = Join-Path $env:TEMP 'pp-powers-corrupt.png'
$q = Start-Process $exe -ArgumentList $file, '--screenshot', $png2, 'main' -PassThru; [void]$q.WaitForExit(40000); $res.corruptExit = if ($q.HasExited) { $q.ExitCode } else { -1 }; if (-not $q.HasExited) { Stop-Process -Id $q.Id -Force }
$res.corruptLuminance = Lum $png2
$res | ConvertTo-Json -Compress
