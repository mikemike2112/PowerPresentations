# Ridimensiona la finestra con UI Automation (TransformPattern/WindowPattern: NESSUN input reale) e cattura la SOLA finestra con PrintWindow.
# Verifica anche il pulsante ingrandisci/ripristina (nome accessibile) a ogni cambio di stato. Uso: ui-resize.ps1 -Shots dir
param([string]$Shots)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class RS { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f); [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
 [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; } }
'@
[void][RS]::SetProcessDPIAware()
$repo = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $repo 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
$file = Join-Path $repo 'tests\corpus\office-made\mixed-realistic-1.pptx'
$env:PP_SETTINGS = Join-Path $env:TEMP 'pp-resize.json'; Remove-Item $env:PP_SETTINGS -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $Shots | Out-Null
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
function Shot($h, $path) { $r = New-Object RS+RECT; [void][RS]::GetWindowRect($h, [ref]$r); $b = New-Object System.Drawing.Bitmap ($r.R - $r.L), ($r.B - $r.T); $g = [System.Drawing.Graphics]::FromImage($b); $hdc = $g.GetHdc(); [void][RS]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose(); $b.Save($path); $b.Dispose() }
function BtnNames($root) { $c = [System.Windows.Automation.PropertyCondition]::new($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button); (($root.FindAll($TS::Descendants, $c) | % { $_.Current.Name }) -join '|') }
$p = Start-Process $exe -ArgumentList $file -PassThru; $res = [ordered]@{}
try {
    $d = (Get-Date).AddSeconds(20); while ((-not $p.MainWindowHandle -or $p.MainWindowHandle -eq 0) -and (Get-Date) -lt $d) { Start-Sleep -Milliseconds 300; $p.Refresh() }
    Start-Sleep -Seconds 2; $h = $p.MainWindowHandle; $root = $AE::FromHandle($h)
    $win = $root.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern); $tr = $root.GetCurrentPattern([System.Windows.Automation.TransformPattern]::Pattern)
    function MaxName { $c = [System.Windows.Automation.PropertyCondition]::new($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button); ($root.FindAll($TS::Descendants, $c) | ? { $_.Current.Name -in 'Ingrandisci', 'Ripristina giù' } | % { $_.Current.Name }) -join '' }
    $res.maxNameNormal = MaxName
    $win.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Maximized); Start-Sleep -Milliseconds 900
    $res.maxNameMaximized = MaxName; Shot $h (Join-Path $Shots 'resize-massimizzata.png')
    $win.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal); Start-Sleep -Milliseconds 900
    $res.maxNameBack = MaxName
    foreach ($w in 1100, 800, 640) {
        $tr.Resize($w * 1.25, 760); Start-Sleep -Milliseconds 1000      # larghezze in DIP a scala 125%
        $r = $root.Current.BoundingRectangle; $res["w$w"] = "$([int]$r.Width)x$([int]$r.Height)"
        Shot $h (Join-Path $Shots "resize-$w.png")
        if ($w -eq 800) {
            $c = [System.Windows.Automation.PropertyCondition]::new($AE::NameProperty, 'Altri comandi'); $more = $root.FindFirst($TS::Descendants, $c)
            $res.overflowButton = [bool]$more -and -not $more.Current.IsOffscreen
            if ($more) {
                $ip = $null; if ($more.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$ip)) { $ip.Invoke(); Start-Sleep -Milliseconds 700 }
                $names = @(); foreach ($e in $AE::RootElement.FindAll($TS::Descendants, [System.Windows.Automation.PropertyCondition]::new($AE::ProcessIdProperty, $p.Id))) { if ($e.Current.Name -in 'Trova e sostituisci', 'Seleziona') { $names += $e.Current.Name } }
                $res.overflowItems = ($names | Sort-Object -Unique) -join ','
            }
        }
    }
    $res.minWidthReached = $res.w640
} finally { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
$res | ConvertTo-Json -Compress
