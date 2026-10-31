# Verifica che il watchdog uccida SOLO il PID di PowerPoint avviato qui quando il battito si ferma (simula un dialogo modale).
. (Join-Path $PSScriptRoot 'Common.ps1')
$ctx = Start-Ppt -TimeoutSec 3
if (-not $ctx.Owned) { Write-Output 'SKIP: un PowerPoint dell''utente era gia'' aperto, test non eseguito'; Stop-Ppt $ctx; exit 2 }
$pid0 = $ctx.Pid
Start-Sleep -Seconds 8   # nessun Ping: il watchdog deve intervenire
$alive = [bool](Get-Process -Id $pid0 -ErrorAction SilentlyContinue)
$killedMark = Test-Path ($ctx.Heartbeat + '.killed')
Stop-Ppt $ctx
Remove-Item ($ctx.Heartbeat + '.killed') -ErrorAction SilentlyContinue
if (-not $alive -and $killedMark) { Write-Output 'WATCHDOG_OK' ; exit 0 } else { Write-Output "WATCHDOG_FAIL alive=$alive mark=$killedMark"; exit 1 }
