<#
  uia.ps1 - read / click the Edge page through the Windows UI Automation tree.
  ASCII-only source: pass Chinese labels through command-line arguments.

  dump      : list element names of the web document
  text      : print only Text / ListItem / Hyperlink names (compact page reading)
  click     : invoke / select / toggle the element whose Name equals -Name
  clicknear : same, but only elements whose ancestor Name contains -Contains
  count     : element count

  -Anchor picks the document that contains that element name (e.g. the exam page),
  so a window with several tabs does not resolve to the wrong document.
#>
param(
    [string]$TitleMatch = '',
    [ValidateSet('dump', 'click', 'clicknear', 'count', 'text')]
    [string]$Action = 'dump',
    [ValidateSet('doc', 'win')]
    [string]$Scope = 'doc',
    [string]$Anchor = '',
    [string]$Name = '',
    [string]$Contains = '',
    [int]$Index = 0,
    [int]$Max = 900,
    [int]$Delay = 250
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

if (-not ('UiaMouse' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class UiaMouse {
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(90);
        mouse_event(0x0002, 0, 0, 0, IntPtr.Zero);
        System.Threading.Thread.Sleep(50);
        mouse_event(0x0004, 0, 0, 0, IntPtr.Zero);
    }
}
"@
}

$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]

function Get-AppWindow {
    $root = $AE::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)
    $wins = $root.FindAll($TS::Children, $cond)
    if ($TitleMatch) {
        foreach ($w in $wins) { if ($w.Current.Name -like ('*' + $TitleMatch + '*')) { return $w } }
    }
    foreach ($w in $wins) { if ($w.Current.ClassName -eq 'Chrome_WidgetWin_1') { return $w } }
    return $null
}

function Get-Target($win) {
    if ($Scope -eq 'win') { return $win }
    $docCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Document)
    $docs = $win.FindAll($TS::Descendants, $docCond)
    if ($docs.Count -eq 0) { return $win }
    if ($Anchor) {
        $anchorCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Anchor)
        foreach ($d in $docs) {
            $hit = $d.FindFirst($TS::Descendants, $anchorCond)
            if ($hit) { return $d }
        }
    }
    return $docs[0]
}

function Invoke-Element($el) {
    $p = $null
    if ($el.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$p)) { $p.Invoke(); return 'invoke' }
    $p = $null
    if ($el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$p)) { $p.Select(); return 'select' }
    $p = $null
    if ($el.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$p)) { $p.Toggle(); return 'toggle' }
    $r = $el.Current.BoundingRectangle
    if ($r.Width -le 0 -or $r.Height -le 0) { return 'noclick' }
    $cx = [int]($r.X + $r.Width / 2)
    $cy = [int]($r.Y + $r.Height / 2)
    [UiaMouse]::Click($cx, $cy)
    return ('mouse-at-' + $cx + ',' + $cy)
}

$win = Get-AppWindow
if (-not $win) { Write-Output 'ERR no-window'; exit 2 }
Start-Sleep -Milliseconds $Delay

$target = Get-Target $win
$all = $target.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

switch ($Action) {

    'count' { Write-Output ('ELEMENTS ' + $all.Count); break }

    'text' {
        $i = 0
        foreach ($e in $all) {
            $ct = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
            if ($ct -ne 'Text' -and $ct -ne 'ListItem' -and $ct -ne 'Hyperlink') { continue }
            $n = $e.Current.Name
            if (-not $n) { continue }
            $i++
            if ($i -gt $Max) { break }
            Write-Output (($n -replace '\s+', ' ').Trim())
        }
        break
    }

    'dump' {
        $i = 0
        foreach ($e in $all) {
            $i++
            if ($i -gt $Max) { break }
            $n = $e.Current.Name
            if ($n -and $n.Trim().Length -gt 0) {
                $ct = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
                $clean = ($n -replace '\s+', ' ').Trim()
                Write-Output ($ct + '|' + $clean)
            }
        }
        Write-Output ('TOTAL ' + $all.Count)
        break
    }

    'click' {
        $hits = @()
        foreach ($e in $all) { if ($e.Current.Name -eq $Name) { $hits += $e } }
        if ($hits.Count -le $Index) { Write-Output ('ERR no-match hits=' + $hits.Count); exit 3 }
        Write-Output ((Invoke-Element $hits[$Index]) + ' hits=' + $hits.Count + ' idx=' + $Index)
        break
    }

    'clicknear' {
        $hits = @()
        foreach ($e in $all) {
            if ($e.Current.Name -ne $Name) { continue }
            $cur = $e
            for ($k = 0; $k -lt 6; $k++) {
                $parent = $null
                try { $parent = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($cur) } catch { $parent = $null }
                if (-not $parent) { break }
                if ($parent.Current.Name -and $parent.Current.Name -like ('*' + $Contains + '*')) { $hits += $e; break }
                $cur = $parent
            }
        }
        if ($hits.Count -le $Index) { Write-Output ('ERR no-near hits=' + $hits.Count); exit 3 }
        $si = $null
        try { if ($hits[$Index].TryGetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern, [ref]$si)) { $si.ScrollIntoView(); Start-Sleep -Milliseconds 350 } } catch { }
        Write-Output ((Invoke-Element $hits[$Index]) + ' near hits=' + $hits.Count)
        break
    }
}
