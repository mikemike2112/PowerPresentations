# Inventario del ribbon di PowerPoint con UI Automation: apre PowerPoint con finestra VISIBILE su una presentazione vuota
# creata qui, visita ogni scheda e registra gruppi e comandi (nome, tipo di controllo). Nessun documento dell'utente.
# Uso: ribbon-inventory.ps1 -OutJson ribbon.json
param([string]$OutJson = (Join-Path $PSScriptRoot 'ribbon-inventory.json'))
. (Join-Path $PSScriptRoot 'Common.ps1')
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
$ctx = Start-Ppt -TimeoutSec 180
$out = [ordered]@{}
try {
    $app = $ctx.App
    $pres = $app.Presentations.Add(-1)          # con finestra
    $null = $pres.Slides.AddSlide(1, $pres.SlideMaster.CustomLayouts.Item(2))
    $app.Visible = -1; Start-Sleep -Seconds 3; Ping-Ppt $ctx
    $root = $AE::FromHandle([IntPtr]$app.HWND)
    $tabs = $root.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::TabItem)))
    $names = @(); foreach ($t in $tabs) { $names += $t.Current.Name }
    $types = @([System.Windows.Automation.ControlType]::Button, [System.Windows.Automation.ControlType]::SplitButton, [System.Windows.Automation.ControlType]::CheckBox, [System.Windows.Automation.ControlType]::ComboBox, [System.Windows.Automation.ControlType]::MenuItem, [System.Windows.Automation.ControlType]::RadioButton)
    foreach ($tabName in $names) {
        Ping-Ppt $ctx
        $tab = $root.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.AndCondition (New-Object System.Windows.Automation.PropertyCondition ($AE::NameProperty), $tabName), (New-Object System.Windows.Automation.PropertyCondition ($AE::ControlTypeProperty), ([System.Windows.Automation.ControlType]::TabItem))))
        $o = $null
        if ($tab -and $tab.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$o)) { try { $o.Select() } catch {} }
        Start-Sleep -Milliseconds 1200
        $groups = [ordered]@{}
        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $all = $root.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($e in $all) {
            try {
                $ct = $e.Current.ControlType
                if ($types -notcontains $ct) { continue }
                $r = $e.Current.BoundingRectangle
                if ($e.Current.IsOffscreen -or $r.Width -le 0 -or $r.Y -lt 0) { continue }
                # solo la fascia del ribbon: sotto le schede e sopra l'area della diapositiva
                $tr = $tab.Current.BoundingRectangle
                if ($r.Y -lt $tr.Y + $tr.Height -or $r.Y -gt $tr.Y + 260) { continue }
                $grp = ''; $p = $walker.GetParent($e)
                while ($p -and $p.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) { if ($p.Current.ControlType -eq [System.Windows.Automation.ControlType]::Group -or $p.Current.ControlType -eq [System.Windows.Automation.ControlType]::Pane) { if ($p.Current.Name) { $grp = $p.Current.Name; break } }; $p = $walker.GetParent($p) }
                $n = $e.Current.Name; if (-not $n) { continue }
                if (-not $groups.Contains($grp)) { $groups[$grp] = @() }
                $groups[$grp] += ("{0} [{1}]" -f $n, $ct.ProgrammaticName.Replace('ControlType.', ''))
            } catch {}
        }
        $out[$tabName] = $groups
    }
    $pres.Saved = -1; $pres.Close()
} finally { Stop-Ppt $ctx }
$json = ConvertTo-Json -InputObject $out -Depth 6
[System.IO.File]::WriteAllText($OutJson, $json, (New-Object System.Text.UTF8Encoding($false)))
"schede: " + ($out.Keys -join ', ')
