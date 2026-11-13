#requires -Version 5.1
<#
  Start-ExamHelper.ps1  v0.2.0 - 智慧树/在线学堂单元测试助手（公开版）

  本版针对四个问题做了改造：
  1. 界面不再置顶、默认停靠屏幕右下角；点“自动作答”后窗口收进托盘，浏览器全程不被遮挡。
  2. 作答改为“读当前选中状态 → 只点差异项”，重复运行不会把已选项点掉，也不会重复点同一题。
  3. 图片选项题可选“跳过”或“停留等我做”，程序不替用户猜图。
  4. 填完与提交前都会核对答题卡（逐题回读选项状态），列出未答题号，可一键跳到第一道未答。

  用法： powershell -NoProfile -ExecutionPolicy Bypass -File .\Start-ExamHelper.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$ErrorActionPreference = 'Stop'

$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]

if (-not ('ExamWin' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ExamWin {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(90);
        mouse_event(0x0002, 0, 0, 0, IntPtr.Zero);
        System.Threading.Thread.Sleep(50);
        mouse_event(0x0004, 0, 0, 0, IntPtr.Zero);
    }
    public static void Focus(IntPtr h) {
        ShowWindow(h, 9);
        SetForegroundWindow(h);
    }
}
"@
}

try { [void][ExamWin]::SetProcessDPIAware() } catch { }

# ---------------------------------------------------------------- 通用动作

function Nap([int]$min, [int]$max) {
    $ms = $min
    if ($max -gt $min) { $ms = Get-Random -Minimum $min -Maximum $max }
    Start-Sleep -Milliseconds $ms
    [System.Windows.Forms.Application]::DoEvents()
    return $ms
}

function Get-BrowserWindow([string]$match) {
    $root = $AE::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)
    $wins = $root.FindAll($TS::Children, $cond)
    if ($match) {
        foreach ($w in $wins) {
            if ($w.Current.Name -like ('*' + $match + '*')) { return $w }
        }
    }
    foreach ($w in $wins) {
        if ($w.Current.ClassName -eq 'Chrome_WidgetWin_1') { return $w }
    }
    return $null
}

function Focus-Browser([string]$match) {
    $win = Get-BrowserWindow $match
    if (-not $win) { return $false }
    [ExamWin]::Focus([IntPtr]$win.Current.NativeWindowHandle)
    return $true
}

function Get-ExamDoc([string]$match) {
    $win = Get-BrowserWindow $match
    if (-not $win) { return $null }
    $docCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Document)
    $docs = $win.FindAll($TS::Descendants, $docCond)
    if ($docs.Count -eq 0) { return $win }
    $anchorCond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, '下一题')
    foreach ($d in $docs) {
        if ($d.FindFirst($TS::Descendants, $anchorCond)) { return $d }
    }
    return $docs[0]
}

function Get-Items($doc) {
    $out = @()
    if (-not $doc) { return $out }
    $all = $doc.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($e in $all) {
        $ctName = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
        if ($ctName -ne 'Text' -and $ctName -ne 'Button' -and $ctName -ne 'ListItem') { continue }
        $n = $e.Current.Name
        if (-not $n) { continue }
        $out += [pscustomobject]@{ T = $ctName; N = ($n -replace '\s+', ' ').Trim(); E = $e }
    }
    return $out
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
    [ExamWin]::Click([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
    return 'mouse'
}

function Click-ByName($doc, [string]$name) {
    foreach ($it in (Get-Items $doc)) {
        if ($it.T -eq 'Button' -and $it.N -eq $name) { return (Invoke-Element $it.E) }
    }
    return 'no-button'
}

function Jump-To($doc, [int]$num) {
    foreach ($it in (Get-Items $doc)) {
        if ($it.T -eq 'ListItem' -and $it.N -eq [string]$num) { return (Invoke-Element $it.E) }
    }
    return 'no-cell'
}

function Get-TotalQuestions($doc) {
    $max = 0
    foreach ($it in (Get-Items $doc)) {
        if ($it.T -eq 'ListItem' -and $it.N -match '^\d+$') {
            $v = [int]$it.N
            if ($v -gt $max) { $max = $v }
        }
    }
    return $max
}

function Get-Completion($doc) {
    foreach ($it in (Get-Items $doc)) {
        if ($it.T -eq 'Text' -and $it.N -match '^完成率\s*(\d+)%') { return [int]$matches[1] }
    }
    return -1
}

function Read-Question($doc) {
    $items = Get-Items $doc
    $start = -1
    for ($i = 0; $i -lt $items.Count; $i++) {
        if ($items[$i].T -eq 'Text' -and $items[$i].N -match '^\d+\.?$') { $start = $i; break }
    }
    if ($start -lt 0) { return $null }

    $num = [int](($items[$start].N -replace '\.$', '').Trim())
    $type = ''
    $stem = ''
    $cur = ''
    $opts = [ordered]@{}
    for ($j = $start + 1; $j -lt $items.Count; $j++) {
        $n = $items[$j].N
        if ($items[$j].T -eq 'Button') { break }
        if ($n.Length -gt 2 -and $n[0] -eq [char]0x3010 -and $n[$n.Length - 1] -eq [char]0x3011) {
            $type = $n.Substring(1, $n.Length - 2)
            continue
        }
        if ($n -match '^\(\d+[^)]*\)$') { continue }
        if ($n -match '^[A-D]\.$') { $cur = $n.Substring(0, 1); $opts[$cur] = ''; continue }
        if ($cur) {
            if ($opts[$cur] -eq '') { $opts[$cur] = $n }
            continue
        }
        if (-not $stem) { $stem = $n } else { $stem = $stem + ' ' + $n }
    }
    $hasText = $false
    foreach ($k in $opts.Keys) { if ($opts[$k]) { $hasText = $true } }
    return [pscustomobject]@{ Num = $num; Type = $type; Stem = $stem; Options = $opts; HasText = $hasText }
}

# ---------------------------------------------------------------- 屏幕采样

function Get-WindowShot([string]$match) {
    $win = Get-BrowserWindow $match
    if (-not $win) { return $null }
    $r = New-Object ExamWin+RECT
    [void][ExamWin]::GetWindowRect([IntPtr]$win.Current.NativeWindowHandle, [ref]$r)
    $w = $r.Right - $r.Left
    $h = $r.Bottom - $r.Top
    if ($w -le 0 -or $h -le 0) { return $null }
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, (New-Object System.Drawing.Size($w, $h)))
    $g.Dispose()
    return [pscustomobject]@{ Bitmap = $bmp; Left = $r.Left; Top = $r.Top }
}

function Test-BlueAt($bmp, [int]$x, [int]$y) {
    if ($x -lt 0 -or $y -lt 0 -or $x -ge $bmp.Width -or $y -ge $bmp.Height) { return $false }
    $p = $bmp.GetPixel($x, $y)
    return (($p.B - $p.R) -gt 35 -and $p.B -gt 110)
}

# 选项左侧小圆点 / 小方框是否处于选中态
function Get-SelectedLetters($doc, [string]$match) {
    $shot = Get-WindowShot $match
    if (-not $shot) { return @() }
    $sel = @()
    try {
        foreach ($it in (Get-Items $doc)) {
            if ($it.T -ne 'Text') { continue }
            if ($it.N -notmatch '^[A-D]\.$') { continue }
            $letter = $it.N.Substring(0, 1)
            $r = $it.E.Current.BoundingRectangle
            if ($r.Width -le 0 -or $r.Height -le 0) { continue }
            $rx = [int]($r.Left - $shot.Left)
            $ry = [int]($r.Top - $shot.Top + $r.Height / 2)
            $hit = 0
            for ($x = $rx - 24; $x -le $rx - 4; $x += 2) {
                for ($y = $ry - 6; $y -le $ry + 6; $y += 3) {
                    if (Test-BlueAt $shot.Bitmap $x $y) { $hit++ }
                }
            }
            if ($hit -ge 3) { $sel += $letter }
        }
    }
    finally { $shot.Bitmap.Dispose() }
    return $sel
}

# ---------------------------------------------------------------- 模型调用

function Ask-Model([string]$baseUrl, [string]$model, [string]$key, $q) {
    $pairs = @()
    foreach ($k in $q.Options.Keys) { $pairs += ($k + '. ' + $q.Options[$k]) }
    $sys = '你是考试答题助手。只输出选项字母，不要解释。单选题与判断题输出 1 个字母；多选题输出多个大写字母连写，例如 ABCD。'
    $user = '题型：' + $q.Type + "`n题干：" + $q.Stem + "`n选项：`n" + ($pairs -join "`n")
    $body = @{
        model       = $model
        temperature = 0
        messages    = @(
            @{ role = 'system'; content = $sys },
            @{ role = 'user'; content = $user }
        )
    } | ConvertTo-Json -Depth 6

    $uri = $baseUrl.TrimEnd('/') + '/chat/completions'
    $headers = @{ Authorization = ('Bearer ' + $key); 'Content-Type' = 'application/json' }
    $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 90
    $text = [string]$resp.choices[0].message.content
    $letters = ''
    foreach ($m in [regex]::Matches($text.ToUpper(), '[A-D]')) {
        if ($letters -notlike ('*' + $m.Value + '*')) { $letters += $m.Value }
    }
    return [pscustomobject]@{ Letters = $letters; Raw = $text }
}

# ---------------------------------------------------------------- 作答：只点差异项

function Apply-Answer($doc, [string]$want, [string]$match) {
    $sel = @(Get-SelectedLetters $doc $match)
    $wantSet = @()
    foreach ($c in $want.ToCharArray()) { $wantSet += [string]$c }

    $toUnclick = @()
    foreach ($s in $sel) { if ($wantSet -notcontains $s) { $toUnclick += $s } }
    $toClick = @()
    foreach ($w in $wantSet) { if ($sel -notcontains $w) { $toClick += $w } }

    foreach ($c in $toUnclick) {
        foreach ($it in (Get-Items $doc)) {
            if ($it.T -eq 'Text' -and $it.N -eq ($c + '.')) { [void](Invoke-Element $it.E) }
        }
        Nap 240 480 | Out-Null
    }
    foreach ($c in $toClick) {
        foreach ($it in (Get-Items $doc)) {
            if ($it.T -eq 'Text' -and $it.N -eq ($c + '.')) { [void](Invoke-Element $it.E) }
        }
        Nap 300 620 | Out-Null
    }
    Nap 250 450 | Out-Null

    $after = @(Get-SelectedLetters $doc $match)
    $ok = $true
    foreach ($w in $wantSet) { if ($after -notcontains $w) { $ok = $false } }
    foreach ($a in $after) { if ($wantSet -notcontains $a) { $ok = $false } }

    return [pscustomobject]@{
        Target    = $want
        Before    = ($sel -join '')
        After     = ($after -join '')
        Unclicked = ($toUnclick -join '')
        Clicked   = ($toClick -join '')
        Ok        = $ok
    }
}

# ---------------------------------------------------------------- 答题卡核对

function Scan-Missing([string]$match, [switch]$Quiet) {
    $doc = Get-ExamDoc $match
    if (-not $doc) { Log '找不到考试页面'; return @() }
    $total = Get-TotalQuestions $doc
    if ($total -le 0) { Log '读不到题号网格，无法扫描'; return @() }

    $completion = Get-Completion $doc
    if ($completion -eq 100) {
        if (-not $Quiet) { Log '答题卡显示完成率 100%' }
        $script:missing = @()
        return @()
    }

    [void](Focus-Browser $match)
    Nap 400 800 | Out-Null

    $missing = @()
    for ($n = 1; $n -le $total; $n++) {
        if (-not $Quiet -and -not $script:scanning) { break }
        $doc = Get-ExamDoc $match
        if (-not $doc) { break }
        [void](Jump-To $doc $n)
        Nap 650 1000 | Out-Null
        $doc = Get-ExamDoc $match
        $sel = @(Get-SelectedLetters $doc $match)
        if ($sel.Count -eq 0) { $missing += $n }
        if (-not $Quiet -and ($n % 10) -eq 0) { Log ('  核对 ' + $n + '/' + $total) }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $script:missing = $missing
    if ($missing.Count -eq 0) { Log '答题卡核对：全部已答' }
    else { Log ('答题卡核对：未作答 ' + ($missing -join ', ')) }
    return $missing
}

# ---------------------------------------------------------------- 界面

$form = New-Object System.Windows.Forms.Form
$form.Text = '智慧树单元测试助手 v0.2.0'
$form.Size = New-Object System.Drawing.Size(640, 460)
$form.StartPosition = 'Manual'
$form.TopMost = $false
$form.MinimumSize = New-Object System.Drawing.Size(520, 220)

$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(($wa.Right - 660), ($wa.Bottom - 480))

function Add-Label($parent, $text, $x, $y, $w) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size($w, 20)
    $parent.Controls.Add($l)
    return $l
}

Add-Label $form 'API Base URL' 12 15 90 | Out-Null
$txtBase = New-Object System.Windows.Forms.TextBox
$txtBase.Location = New-Object System.Drawing.Point(104, 12)
$txtBase.Size = New-Object System.Drawing.Size(510, 24)
$txtBase.Text = 'https://api.deepseek.com/v1'
$form.Controls.Add($txtBase)

Add-Label $form 'Model' 12 45 90 | Out-Null
$txtModel = New-Object System.Windows.Forms.TextBox
$txtModel.Location = New-Object System.Drawing.Point(104, 42)
$txtModel.Size = New-Object System.Drawing.Size(250, 24)
$txtModel.Text = 'deepseek-chat'
$form.Controls.Add($txtModel)

Add-Label $form '窗口标题片段' 368 45 96 | Out-Null
$txtWin = New-Object System.Windows.Forms.TextBox
$txtWin.Location = New-Object System.Drawing.Point(466, 42)
$txtWin.Size = New-Object System.Drawing.Size(148, 24)
$txtWin.Text = '在线学堂'
$form.Controls.Add($txtWin)

Add-Label $form 'API Key' 12 75 90 | Out-Null
$txtKey = New-Object System.Windows.Forms.TextBox
$txtKey.Location = New-Object System.Drawing.Point(104, 72)
$txtKey.Size = New-Object System.Drawing.Size(380, 24)
$txtKey.UseSystemPasswordChar = $true
$form.Controls.Add($txtKey)

$btnShow = New-Object System.Windows.Forms.Button
$btnShow.Text = '显示'
$btnShow.Location = New-Object System.Drawing.Point(490, 71)
$btnShow.Size = New-Object System.Drawing.Size(56, 26)
$btnShow.Add_Click({ $txtKey.UseSystemPasswordChar = -not $txtKey.UseSystemPasswordChar })
$form.Controls.Add($btnShow)

$grpImage = New-Object System.Windows.Forms.GroupBox
$grpImage.Text = '图片题'
$grpImage.Location = New-Object System.Drawing.Point(12, 104)
$grpImage.Size = New-Object System.Drawing.Size(292, 52)
$form.Controls.Add($grpImage)

$rdoSkip = New-Object System.Windows.Forms.RadioButton
$rdoSkip.Text = '跳过'
$rdoSkip.Checked = $true
$rdoSkip.Location = New-Object System.Drawing.Point(12, 20)
$rdoSkip.Size = New-Object System.Drawing.Size(70, 22)
$grpImage.Controls.Add($rdoSkip)

$rdoWait = New-Object System.Windows.Forms.RadioButton
$rdoWait.Text = '停留等我做'
$rdoWait.Location = New-Object System.Drawing.Point(92, 20)
$rdoWait.Size = New-Object System.Drawing.Size(120, 22)
$grpImage.Controls.Add($rdoWait)

$chkTop = New-Object System.Windows.Forms.CheckBox
$chkTop.Text = '窗口置顶'
$chkTop.Location = New-Object System.Drawing.Point(320, 108)
$chkTop.Size = New-Object System.Drawing.Size(100, 22)
$chkTop.Add_Click({ $form.TopMost = $chkTop.Checked })
$form.Controls.Add($chkTop)

$btnRead = New-Object System.Windows.Forms.Button
$btnRead.Text = '读取当前题'
$btnRead.Location = New-Object System.Drawing.Point(428, 104)
$btnRead.Size = New-Object System.Drawing.Size(90, 26)
$form.Controls.Add($btnRead)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = '检查答题卡'
$btnScan.Location = New-Object System.Drawing.Point(12, 164)
$btnScan.Size = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnScan)

$btnJump = New-Object System.Windows.Forms.Button
$btnJump.Text = '跳到第一道未答'
$btnJump.Location = New-Object System.Drawing.Point(120, 164)
$btnJump.Size = New-Object System.Drawing.Size(120, 26)
$form.Controls.Add($btnJump)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = '停止'
$btnStop.Location = New-Object System.Drawing.Point(248, 164)
$btnStop.Size = New-Object System.Drawing.Size(64, 26)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$btnAuto = New-Object System.Windows.Forms.Button
$btnAuto.Text = '自动作答'
$btnAuto.Location = New-Object System.Drawing.Point(320, 164)
$btnAuto.Size = New-Object System.Drawing.Size(90, 26)
$form.Controls.Add($btnAuto)

$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = '提交（需确认）'
$btnSubmit.Location = New-Object System.Drawing.Point(418, 164)
$btnSubmit.Size = New-Object System.Drawing.Size(110, 26)
$form.Controls.Add($btnSubmit)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(12, 200)
$txtLog.Size = New-Object System.Drawing.Size(600, 200)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtLog.Anchor = 'Top,Left,Right,Bottom'
$form.Controls.Add($txtLog)

$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.Text = '智慧树单元测试助手'
$notify.Visible = $true
$notify.Add_MouseDoubleClick({
    $form.Show()
    $form.WindowState = 'Normal'
})

function Log([string]$msg) {
    $line = ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $msg)
    $txtLog.AppendText($line + "`r`n")
    [System.Windows.Forms.Application]::DoEvents()
}

$script:running = $false
$script:scanning = $false
$script:missing = @()
$script:answeredThisRun = @()
$script:skippedThisRun = @()

# ---------------------------------------------------------------- 事件：读取当前题

$btnRead.Add_Click({
    try {
        $doc = Get-ExamDoc $txtWin.Text
        if (-not $doc) { Log '找不到浏览器窗口，确认浏览器开着且标题含该片段'; return }
        $q = Read-Question $doc
        if (-not $q) { Log '没读到题号，可能不在考试页或页面文本不透出'; return }
        Log ('第 ' + $q.Num + ' 题 [' + $q.Type + '] ' + $q.Stem)
        foreach ($k in $q.Options.Keys) { Log ('  ' + $k + '. ' + $q.Options[$k]) }
        if (-not $q.HasText) { Log '  → 图片题（选项无文字）' }
        $sel = @(Get-SelectedLetters $doc $txtWin.Text)
        if ($sel.Count -gt 0) { Log ('  当前已选：' + ($sel -join '')) } else { Log '  当前已选：（无）' }
    }
    catch { Log ('读取失败：' + $_.Exception.Message) }
})

# ---------------------------------------------------------------- 事件：自动作答

$btnAuto.Add_Click({
    if (-not $txtKey.Text) { Log '先填 API Key'; return }
    $script:running = $true
    $script:answeredThisRun = @()
    $script:skippedThisRun = @()
    $btnAuto.Enabled = $false
    $btnStop.Enabled = $true

    $wasVisible = $form.Visible
    $form.Hide()
    [void](Focus-Browser $txtWin.Text)

    try {
        Nap 600 1000 | Out-Null
        $lastNum = 0
        $total = 0

        for ($step = 0; $step -lt 80; $step++) {
            if (-not $script:running) { Log '已停止'; break }
            if ($step -gt 0 -and ($step % 9) -eq 0) {
                [void](Focus-Browser $txtWin.Text)
                Log ('  停顿 ' + (Nap 2000 4000) + 'ms')
            }

            $doc = Get-ExamDoc $txtWin.Text
            if (-not $doc) { Log '浏览器窗口丢了，停止'; break }
            if ($total -le 0) { $total = Get-TotalQuestions $doc }

            $q = $null
            for ($try = 0; $try -lt 6; $try++) {
                $q = Read-Question $doc
                if ($q) { break }
                Nap 350 450 | Out-Null
                $doc = Get-ExamDoc $txtWin.Text
            }
            if (-not $q) { Log '读不到题目，停止'; break }
            if ($q.Num -le $lastNum) { Log ('题号没前进（' + $q.Num + '），停止'); break }
            $lastNum = $q.Num

            if (-not $q.HasText) {
                if ($rdoWait.Checked) {
                    $form.Show()
                    $script:running = $false
                    $r = [System.Windows.Forms.MessageBox]::Show(('第 ' + $q.Num + ' 题是图片题，请在浏览器里自己选好后点「确定」继续。'), '图片题', 'OKCancel', 'Information')
                    $script:running = $true
                    $form.Hide()
                    if ($r -ne 'OK') { Log '用户中止'; break }
                    [void](Focus-Browser $txtWin.Text)
                    $script:skippedThisRun += $q.Num
                    Log ('第 ' + $q.Num + ' 题：用户自行作答')
                }
                else {
                    $script:skippedThisRun += $q.Num
                    Log ('第 ' + $q.Num + ' 题：图片题，跳过（可稍后自己补）')
                }
            }
            else {
                try {
                    $ans = Ask-Model $txtBase.Text $txtModel.Text $txtKey.Text $q
                    if (-not $ans.Letters) {
                        Log ('第 ' + $q.Num + ' 题：模型没给字母，跳过。原文=' + $ans.Raw)
                    }
                    else {
                        $res = Apply-Answer $doc $ans.Letters $txtWin.Text
                        $script:answeredThisRun += $q.Num
                        $now = '空'
                        if ($res -and $res.After) { $now = $res.After }
                        $old = '空'
                        if ($res -and $res.Before) { $old = $res.Before }
                        if ($res -and $res.Ok) {
                            Log ('第 ' + $q.Num + ' 题 [' + $q.Type + '] 目标 ' + $ans.Letters + '（原 ' + $old + ' → 现 ' + $now + '）')
                        }
                        else {
                            Log ('第 ' + $q.Num + ' 题：点击后与目标不一致（目标 ' + $ans.Letters + '，实际 ' + $now + '），请人工确认')
                        }
                    }
                }
                catch { Log ('第 ' + $q.Num + ' 题：模型调用失败 - ' + $_.Exception.Message) }
            }

            $doc = Get-ExamDoc $txtWin.Text
            Nap 200 400 | Out-Null
            $clicked = Click-ByName $doc '下一题'
            if ($clicked -eq 'no-button') { $clicked = Click-ByName $doc '保存' }
            Nap 900 1700 | Out-Null
        }

        $msg = '本套填答结束：作答 ' + $script:answeredThisRun.Count + ' 题'
        if ($script:skippedThisRun.Count -gt 0) { $msg += '；图片题 ' + ($script:skippedThisRun -join ', ') }
        Log $msg
        Log '开始核对答题卡…'
        $script:scanning = $true
        try { $missing = @(Scan-Missing $txtWin.Text) } finally { $script:scanning = $false }
        if ($missing.Count -gt 0) {
            $msg += "`r`n`r`n仍未作答：" + ($missing -join ', ') + "`r`n补完再点「提交（需确认）」，提交前会再核对一次。"
        }
        else {
            $msg += "`r`n`r`n答题卡显示全部已答，可以提交。"
        }
        [void][System.Windows.Forms.MessageBox]::Show($msg, '请检查后提交')
    }
    catch { Log ('中断：' + $_.Exception.Message) }
    finally {
        $script:running = $false
        $btnAuto.Enabled = $true
        $btnStop.Enabled = $false
        if ($wasVisible) { $form.Show() }
    }
})

$btnStop.Add_Click({
    $script:running = $false
    $script:scanning = $false
    Log '收到停止指令'
})

$btnScan.Add_Click({
    $script:running = $true
    $script:scanning = $true
    $btnScan.Enabled = $false
    try { [void](Scan-Missing $txtWin.Text) }
    finally {
        $script:scanning = $false
        $script:running = $false
        $btnScan.Enabled = $true
    }
})

$btnJump.Add_Click({
    if ($script:missing.Count -eq 0) { Log '没有记录到未答题（先点「检查答题卡」）'; return }
    $target = $script:missing[0]
    $doc = Get-ExamDoc $txtWin.Text
    if (-not $doc) { Log '找不到考试页面'; return }
    [void](Focus-Browser $txtWin.Text)
    Log ('跳到第 ' + $target + ' 题 → ' + (Jump-To $doc $target))
})

$btnSubmit.Add_Click({
    $doc = Get-ExamDoc $txtWin.Text
    if (-not $doc) { Log '找不到考试页面'; return }
    $completion = Get-Completion $doc
    if ($completion -lt 100) {
        Log ('完成率 ' + $completion + '%，先核对一遍…')
        $script:running = $true
        $script:scanning = $true
        try { [void](Scan-Missing $txtWin.Text) }
        finally { $script:scanning = $false; $script:running = $false }
    }
    $tail = ''
    if ($script:missing.Count -gt 0) { $tail = "`r`n未作答：" + ($script:missing -join ', ') }
    $text = '完成率 ' + $completion + '%' + $tail + "`r`n`r`n确认提交？提交后不能再修改。"
    $yes = [System.Windows.Forms.MessageBox]::Show($text, '二次确认', 'YesNo', 'Warning')
    if ($yes -ne 'Yes') { Log '已取消提交'; return }
    try {
        [void](Focus-Browser $txtWin.Text)
        Nap 400 700 | Out-Null
        $doc = Get-ExamDoc $txtWin.Text
        Log ('提交作业 → ' + (Click-ByName $doc '提交作业'))
        Nap 1200 1800 | Out-Null
        $doc = Get-ExamDoc $txtWin.Text
        Log ('确认 → ' + (Click-ByName $doc '确定'))
    }
    catch { Log ('提交失败：' + $_.Exception.Message) }
})

$form.Add_Resize({
    if ($form.WindowState -eq 'Minimized') { $form.Hide() }
})

[void]$form.ShowDialog()
