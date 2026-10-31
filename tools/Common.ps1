# Funzioni comuni: avvio di PowerPoint con PID isolato e watchdog. Solo POWERPNT.EXE avviato da questi script.
$ErrorActionPreference = 'Stop'

function Start-Ppt {
    param([int]$TimeoutSec = 60)
    $before = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $app = New-Object -ComObject PowerPoint.Application
    Start-Sleep -Milliseconds 500
    $after = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $new = @($after | Where-Object { $before -notcontains $_ })
    $ctx = [pscustomobject]@{ App = $app; Pid = $(if ($new.Count -ge 1) { $new[0] } else { 0 }); Owned = ($new.Count -ge 1); Heartbeat = $null; Watchdog = $null }
    if ($ctx.Owned) {
        $ctx.Heartbeat = Join-Path $env:TEMP ("ppt-hb-" + $ctx.Pid + ".txt")
        Set-Content $ctx.Heartbeat (Get-Date).Ticks
        $wd = Join-Path $PSScriptRoot 'watchdog.ps1'
        $ctx.Watchdog = Start-Process powershell -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile', '-File', $wd, '-PptPid', $ctx.Pid, '-OwnerPid', $PID, '-Heartbeat', $ctx.Heartbeat, '-TimeoutSec', $TimeoutSec)
    }
    return $ctx
}

function Ping-Ppt($ctx) { if ($ctx.Heartbeat) { Set-Content $ctx.Heartbeat (Get-Date).Ticks } }

function Test-PptAlive($ctx) { return $ctx.Owned -and [bool](Get-Process -Id $ctx.Pid -ErrorAction SilentlyContinue) }

function Stop-Ppt($ctx) {
    # chiude SOLO l'istanza avviata da noi; se era di proprieta' dell'utente non la tocca
    if ($ctx.Owned) {
        try { if (Test-PptAlive $ctx) { $ctx.App.DisplayAlerts = 1; $ctx.App.Quit() } } catch {}
    }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.App) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($ctx.Owned) {
        $deadline = (Get-Date).AddSeconds(8)
        while ((Get-Process -Id $ctx.Pid -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
        if (Get-Process -Id $ctx.Pid -ErrorAction SilentlyContinue) { Stop-Process -Id $ctx.Pid -Force }
        if ($ctx.Watchdog -and -not $ctx.Watchdog.HasExited) { Stop-Process -Id $ctx.Watchdog.Id -Force -ErrorAction SilentlyContinue }
        if ($ctx.Heartbeat) { Remove-Item $ctx.Heartbeat -ErrorAction SilentlyContinue }
    }
}
