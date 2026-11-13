<#
  cu.ps1 - dependency-free Windows "computer use" helper.
  Reads the screen you already have open; never starts a browser.

  Actions:
    shot      capture full virtual screen (or one window with -Match) to PNG
    windows   list visible top-level windows (handle + title)
    rect      print foreground / matched window rect
    activate  bring a window to front by title match
    click / rclick / dblclick / move   pointer actions at -X -Y
    type      paste -Text into the focused control
    key       send -Keys (SendKeys syntax, e.g. "{ENTER}" "^t" "{PGDN}")
    scroll    wheel by -Amount (positive = up, negative = down)

  Examples:
    powershell -NoProfile -ExecutionPolicy Bypass -File cu.ps1 -Action shot
    powershell -NoProfile -ExecutionPolicy Bypass -File cu.ps1 -Action shot -Match "Edge"
    powershell -NoProfile -ExecutionPolicy Bypass -File cu.ps1 -Action windows
    powershell -NoProfile -ExecutionPolicy Bypass -File cu.ps1 -Action click -X 900 -Y 430
    powershell -NoProfile -ExecutionPolicy Bypass -File cu.ps1 -Action scroll -Amount -3
    powershell -NoProfile -ExecutionPolicy Bypass -File cu.ps1 -Action key -Keys "{PGDN}"
#>
[CmdletBinding()]
param(
    [ValidateSet('shot', 'windows', 'rect', 'activate', 'click', 'rclick', 'dblclick', 'move', 'type', 'key', 'scroll', 'wclick')]
    [string]$Action = 'shot',
    [string]$Out = '',
    [int]$X = 0,
    [int]$Y = 0,
    [string]$Text = '',
    [string]$Keys = '',
    [int]$Amount = 3,
    [string]$Match = '',
    [int]$WaitMs = 250
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ('Cu32' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Cu32 {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, IntPtr extra);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
"@
}

try { [void][Cu32]::SetProcessDPIAware() } catch { }

function Get-TopWindows {
    $script:cuWins = New-Object System.Collections.ArrayList
    $cb = [Cu32+EnumProc] {
        param([IntPtr]$h, [IntPtr]$l)
        if ([Cu32]::IsWindowVisible($h)) {
            $sb = New-Object System.Text.StringBuilder 512
            [void][Cu32]::GetWindowTextW($h, $sb, 512)
            $t = $sb.ToString()
            if ($t -and $t.Trim().Length -gt 0) {
                [void]$script:cuWins.Add([pscustomobject]@{ hwnd = $h.ToInt64(); title = $t })
            }
        }
        return $true
    }
    [void][Cu32]::EnumWindows($cb, [IntPtr]::Zero)
    return $script:cuWins
}

function Find-Window([string]$needle) {
    $all = Get-TopWindows
    $hit = @($all | Where-Object { $_.title -like ('*' + $needle + '*') })
    if ($hit.Count -eq 0) { return $null }
    return $hit[0]
}

function Invoke-Wheel([int]$notches) {
    if ($notches -eq 0) { return }
    $delta = [int]([math]::Sign($notches) * 120)
    $steps = [math]::Abs($notches)
    $u = [BitConverter]::ToUInt32([BitConverter]::GetBytes($delta), 0)
    for ($i = 0; $i -lt $steps; $i++) {
        [void][Cu32]::mouse_event(0x0800, 0, 0, $u, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 40
    }
}

switch ($Action) {

    'windows' {
        Get-TopWindows | Sort-Object title | Format-Table -AutoSize hwnd, title | Out-String -Width 400
        break
    }

    'rect' {
        if ($Match) {
            $w = Find-Window $Match
            if (-not $w) { Write-Output ('ERR no-window ' + $Match); exit 2 }
            $h = [IntPtr]$w.hwnd
        }
        else {
            $h = [Cu32]::GetForegroundWindow()
        }
        $r = New-Object Cu32+RECT
        [void][Cu32]::GetWindowRect($h, [ref]$r)
        Write-Output ("RECT " + $r.Left + " " + $r.Top + " " + ($r.Right - $r.Left) + " " + ($r.Bottom - $r.Top))
        break
    }

    'activate' {
        if (-not $Match) { Write-Output 'ERR match-required'; exit 2 }
        $w = Find-Window $Match
        if (-not $w) { Write-Output ('ERR no-window ' + $Match); exit 2 }
        $h = [IntPtr]$w.hwnd
        [void][Cu32]::ShowWindow($h, 9)
        Start-Sleep -Milliseconds 120
        try {
            $sh = New-Object -ComObject WScript.Shell
            [void]$sh.AppActivate($w.title)
        }
        catch { }
        [void][Cu32]::SetForegroundWindow($h)
        Start-Sleep -Milliseconds $WaitMs
        Write-Output ('ACTIVATED ' + $w.title)
        break
    }

    'shot' {
        if ($Out) {
            $path = $Out
        }
        else {
            $path = Join-Path $PSScriptRoot ('shots\shot-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '.png')
        }
        $dir = Split-Path -Parent $path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
        }

        if ($Match) {
            $w = Find-Window $Match
            if (-not $w) { Write-Output ('ERR no-window ' + $Match); exit 2 }
            $r = New-Object Cu32+RECT
            [void][Cu32]::GetWindowRect([IntPtr]$w.hwnd, [ref]$r)
            $left = $r.Left; $top = $r.Top
            $width = $r.Right - $r.Left; $height = $r.Bottom - $r.Top
            if ($width -le 0 -or $height -le 0) { Write-Output 'ERR empty-window'; exit 2 }
        }
        else {
            $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
            $left = $vs.Left; $top = $vs.Top; $width = $vs.Width; $height = $vs.Height
        }

        $bmp = New-Object System.Drawing.Bitmap($width, $height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($left, $top, 0, 0, (New-Object System.Drawing.Size($width, $height)))
        $g.Dispose()
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        Write-Output ('SHOT ' + (Resolve-Path -LiteralPath $path).Path + ' ' + $width + 'x' + $height)
        break
    }

    'wclick' {
        if (-not $Match) { Write-Output 'ERR match-required'; exit 2 }
        $w = Find-Window $Match
        if (-not $w) { Write-Output ('ERR no-window ' + $Match); exit 2 }
        $r = New-Object Cu32+RECT
        [void][Cu32]::GetWindowRect([IntPtr]$w.hwnd, [ref]$r)
        $sx = $r.Left + $X
        $sy = $r.Top + $Y
        [void][Cu32]::SetCursorPos($sx, $sy)
        Start-Sleep -Milliseconds 90
        [void][Cu32]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 40
        [void][Cu32]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
        Write-Output ('WCLICK ' + $sx + ' ' + $sy)
        break
    }

    'move' {
        [void][Cu32]::SetCursorPos($X, $Y)
        Write-Output ('MOVED ' + $X + ' ' + $Y)
        break
    }

    { $_ -in 'click', 'rclick', 'dblclick' } {
        [void][Cu32]::SetCursorPos($X, $Y)
        Start-Sleep -Milliseconds 80
        if ($Action -eq 'click') {
            [void][Cu32]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 30
            [void][Cu32]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
        }
        elseif ($Action -eq 'rclick') {
            [void][Cu32]::mouse_event(0x0008, 0, 0, 0, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 30
            [void][Cu32]::mouse_event(0x0010, 0, 0, 0, [IntPtr]::Zero)
        }
        else {
            for ($i = 0; $i -lt 2; $i++) {
                [void][Cu32]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
                [void][Cu32]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
                Start-Sleep -Milliseconds 60
            }
        }
        Write-Output ('CLICK ' + $Action + ' ' + $X + ' ' + $Y)
        break
    }

    'type' {
        if ($Text.Length -eq 0) { Write-Output 'ERR text-required'; exit 2 }
        Set-Clipboard -Value $Text
        Start-Sleep -Milliseconds 60
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds $WaitMs
        Write-Output ('TYPED ' + $Text.Length + ' chars')
        break
    }

    'key' {
        if ($Keys.Length -eq 0) { Write-Output 'ERR keys-required'; exit 2 }
        [System.Windows.Forms.SendKeys]::SendWait($Keys)
        Start-Sleep -Milliseconds $WaitMs
        Write-Output ('KEY ' + $Keys)
        break
    }

    'scroll' {
        Invoke-Wheel $Amount
        Write-Output ('SCROLL ' + $Amount)
        break
    }
}
