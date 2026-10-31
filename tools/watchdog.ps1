# Uccide il PID di PowerPoint (solo quello indicato) se il battito non si aggiorna entro TimeoutSec
# oppure se il processo proprietario e' terminato. Serve contro i dialoghi modali (es. ripristino file).
param([int]$PptPid, [int]$OwnerPid, [string]$Heartbeat, [int]$TimeoutSec = 60)
while ($true) {
    Start-Sleep -Seconds 1
    if (-not (Get-Process -Id $PptPid -ErrorAction SilentlyContinue)) { exit 0 }
    $ownerGone = -not (Get-Process -Id $OwnerPid -ErrorAction SilentlyContinue)
    $stale = $false
    if (Test-Path $Heartbeat) { $stale = ((Get-Date) - (Get-Item $Heartbeat).LastWriteTime).TotalSeconds -gt $TimeoutSec } else { exit 0 }
    if ($ownerGone -or $stale) {
        Stop-Process -Id $PptPid -Force -ErrorAction SilentlyContinue
        if ($stale) { Set-Content ($Heartbeat + '.killed') 'killed' }
        exit 0
    }
}
