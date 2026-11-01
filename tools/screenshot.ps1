# Avvia l'app su un file e salva uno screenshot della SOLA finestra (RenderTargetBitmap, nessun dato dello schermo).
# modi: main | edit | light | present (finestra presentazione + vista relatore) | slideN
param([string]$File, [string]$Out, [string]$Mode = '')
$exe = Join-Path (Split-Path $PSScriptRoot -Parent) 'src\PowerPresentations.App\bin\Debug\net10.0-windows\PowerPresentations.App.exe'
$a = @($File, '--screenshot', $Out); if ($Mode) { $a += $Mode }
$env:PP_SETTINGS = Join-Path $env:TEMP ('pp-shot-' + $Mode + '.json'); Remove-Item $env:PP_SETTINGS -ErrorAction SilentlyContinue   # impostazioni isolate: lo screenshot parte dal tema predefinito (scuro)
if ($Mode -like 'dialog*') { $env:PP_DEFAULT_FOLDER = 'C:\Documenti' }
$p = Start-Process $exe -ArgumentList $a -PassThru
if (-not $p.WaitForExit(60000)) { Stop-Process -Id $p.Id -Force; Write-Output "TIMEOUT $Mode" } else { Write-Output "$Mode exit $($p.ExitCode)" }
