#requires -Version 5.1
<#
  Start-ExamHelper.ps1  v0.4.0 - 智慧树/在线学堂单元测试助手（公开版）

  v0.4.0
    - 图片题支持“自动看图作答”，同时保留“跳过 / 停留等我做”两种人工选项
    - 界面重排：统一字体、按钮自适应宽度、不再出现文字被截断
    - 免责声明按用户提供的定稿替换（第二条改为三种图片题模式）

  v0.3.0
    - 启动免责声明（可按版本永久关闭）、GitHub 更新提示、配置持久化

  v0.2.0
    - 运行收托盘、只点差异项、图片题可选、答题卡核对

  用法： powershell -NoProfile -ExecutionPolicy Bypass -File .\Start-ExamHelper.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$ErrorActionPreference = 'Stop'

$script:AppVersion = '0.4.0'
$script:RepoOwner = 'a3323554475-netizen'
$script:RepoName = 'zhihuishu-exam-helper'
$script:RepoUrl = "https://github.com/$($script:RepoOwner)/$($script:RepoName)"
$script:IssuesUrl = "$($script:RepoUrl)/issues"
$script:SettingsDir = Join-Path $env:APPDATA 'ZhihuishuExamHelper'
$script:SettingsPath = Join-Path $script:SettingsDir 'settings.json'

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

# ---------------------------------------------------------------- 设置

function Load-Settings {
    $obj = [ordered]@{}
    if (Test-Path -LiteralPath $script:SettingsPath) {
        try {
            $raw = Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8
            $json = $raw | ConvertFrom-Json
            foreach ($p in $json.PSObject.Properties) { $obj[$p.Name] = $p.Value }
        }
        catch { }
    }
    return $obj
}

function Save-Settings($obj) {
    try {
        if (-not (Test-Path -LiteralPath $script:SettingsDir)) {
            New-Item -ItemType Directory -Force -Path $script:SettingsDir | Out-Null
        }
        ($obj | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
    }
    catch { }
}

$script:settings = Load-Settings

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
    return [pscustomobject]@{ Bitmap = $bmp; Left = $r.Left; Top = $r.Top; Width = $w; Height = $h }
}

function Test-BlueAt($bmp, [int]$x, [int]$y) {
    if ($x -lt 0 -or $y -lt 0 -or $x -ge $bmp.Width -or $y -ge $bmp.Height) { return $false }
    $p = $bmp.GetPixel($x, $y)
    return (($p.B - $p.R) -gt 35 -and $p.B -gt 110)
}

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

# 截取题干与四个选项（含图片）作为一张图，供视觉模型作答
function Get-QuestionImage([string]$match, $q) {
    $shot = Get-WindowShot $match
    if (-not $shot) { return $null }
    try {
        $items = Get-Items (Get-ExamDoc $match)
        $startRect = $null
        $lastRect = $null
        foreach ($it in $items) {
            if ($it.T -ne 'Text') { continue }
            if (-not $startRect -and $it.N -match '^\d+\.?$') { $startRect = $it.E.Current.BoundingRectangle }
            if ($it.N -match '^[A-D]\.$') { $lastRect = $it.E.Current.BoundingRectangle }
        }
        if (-not $startRect -or -not $lastRect) { return $null }

        $left = [int]($startRect.Left - $shot.Left - 12)
        $top = [int]($startRect.Top - $shot.Top - 8)
        $right = [int]($lastRect.Left - $shot.Left + 620)
        $bottom = [int]($lastRect.Top - $shot.Top + 90)

        if ($left -lt 0) { $left = 0 }
        if ($top -lt 0) { $top = 0 }
        if ($right -gt $shot.Width - 1) { $right = $shot.Width - 1 }
        if ($bottom -gt $shot.Height - 1) { $bottom = $shot.Height - 1 }
        $w = $right - $left
        $h = $bottom - $top
        if ($w -lt 40 -or $h -lt 40) { return $null }

        $rect = New-Object System.Drawing.Rectangle($left, $top, $w, $h)
        $crop = $shot.Bitmap.Clone($rect, $shot.Bitmap.PixelFormat)
        $ms = New-Object System.IO.MemoryStream
        $crop.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $crop.Dispose()
        $bytes = $ms.ToArray()
        $ms.Dispose()
        return [Convert]::ToBase64String($bytes)
    }
    catch { return $null }
    finally { $shot.Bitmap.Dispose() }
}

# ---------------------------------------------------------------- 模型调用

function New-ModelBody([string]$model, $messages) {
    return @{
        model       = $model
        temperature = 0
        messages    = $messages
    } | ConvertTo-Json -Depth 10
}

function ConvertTo-Letters([string]$text) {
    $letters = ''
    foreach ($m in [regex]::Matches($text.ToUpper(), '[A-D]')) {
        if ($letters -notlike ('*' + $m.Value + '*')) { $letters += $m.Value }
    }
    return $letters
}

function Ask-Model([string]$baseUrl, [string]$model, [string]$key, $q) {
    $pairs = @()
    foreach ($k in $q.Options.Keys) { $pairs += ($k + '. ' + $q.Options[$k]) }
    $sys = '你是考试答题助手。只输出选项字母，不要解释。单选题与判断题输出 1 个字母；多选题输出多个大写字母连写，例如 ABCD。'
    $user = '题型：' + $q.Type + "`n题干：" + $q.Stem + "`n选项：`n" + ($pairs -join "`n")
    $body = New-ModelBody $model @(
        @{ role = 'system'; content = $sys },
        @{ role = 'user'; content = $user }
    )
    $uri = $baseUrl.TrimEnd('/') + '/chat/completions'
    $headers = @{ Authorization = ('Bearer ' + $key); 'Content-Type' = 'application/json' }
    $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 90
    $text = [string]$resp.choices[0].message.content
    return [pscustomobject]@{ Letters = (ConvertTo-Letters $text); Raw = $text }
}

# 图片题：把题目截图交给支持图像输入的模型
function Ask-Model-Vision([string]$baseUrl, [string]$model, [string]$key, $q, [string]$pngBase64) {
    $sys = '你是考试答题助手。图片中是一道题和它的选项截图。只输出正确选项的字母，不要解释；多选题输出多个大写字母连写，例如 ABCD。'
    $userText = '题型：' + $q.Type + '。请根据图片判断正确选项字母。'
    $content = @(
        @{ type = 'text'; text = $userText },
        @{ type = 'image_url'; image_url = @{ url = ('data:image/png;base64,' + $pngBase64) } }
    )
    $body = New-ModelBody $model @(
        @{ role = 'system'; content = $sys },
        @{ role = 'user'; content = $content }
    )
    $uri = $baseUrl.TrimEnd('/') + '/chat/completions'
    $headers = @{ Authorization = ('Bearer ' + $key); 'Content-Type' = 'application/json' }
    $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 120
    $text = [string]$resp.choices[0].message.content
    return [pscustomobject]@{ Letters = (ConvertTo-Letters $text); Raw = $text }
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

# ---------------------------------------------------------------- 主界面

$form = New-Object System.Windows.Forms.Form
$form.Text = '智慧树单元测试助手 v' + $script:AppVersion
$form.ClientSize = New-Object System.Drawing.Size(780, 600)
$form.MinimumSize = New-Object System.Drawing.Size(720, 520)
$form.StartPosition = 'Manual'
$form.TopMost = $false
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$startX = $wa.Right - 820
$startY = $wa.Bottom - 660
if ($script:settings.Contains('windowX') -and $script:settings.Contains('windowY')) {
    try {
        $startX = [int]$script:settings['windowX']
        $startY = [int]$script:settings['windowY']
    } catch { }
}
$form.Location = New-Object System.Drawing.Point($startX, $startY)

function New-FlatButton([string]$text, [int]$x, [int]$y, [int]$w, [int]$h) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Location = New-Object System.Drawing.Point($x, $y)
    $b.Size = New-Object System.Drawing.Size($w, $h)
    $b.FlatStyle = 'System'
    $b.UseVisualStyleBackColor = $true
    $b.TextAlign = 'MiddleCenter'
    return $b
}

$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Location = New-Object System.Drawing.Point(8, 6)
$panelTop.Size = New-Object System.Drawing.Size(764, 176)
$panelTop.Anchor = 'Top,Left,Right'
$panelTop.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($panelTop)

function Add-Caption($parent, [string]$text, [int]$x, [int]$y, [int]$w) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.TextAlign = 'MiddleRight'
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size($w, 24)
    $l.ForeColor = [System.Drawing.Color]::FromArgb(70, 74, 82)
    $parent.Controls.Add($l)
    return $l
}

Add-Caption $panelTop 'API Base URL' 6 12 122 | Out-Null
$txtBase = New-Object System.Windows.Forms.TextBox
$txtBase.Location = New-Object System.Drawing.Point(134, 12)
$txtBase.Size = New-Object System.Drawing.Size(620, 26)
$txtBase.Text = 'https://api.deepseek.com/v1'
$panelTop.Controls.Add($txtBase)

Add-Caption $panelTop 'Model' 6 46 122 | Out-Null
$txtModel = New-Object System.Windows.Forms.TextBox
$txtModel.Location = New-Object System.Drawing.Point(134, 46)
$txtModel.Size = New-Object System.Drawing.Size(240, 26)
$txtModel.Text = 'deepseek-chat'
$panelTop.Controls.Add($txtModel)

Add-Caption $panelTop '窗口标题片段' 358 44 104 | Out-Null
$txtWin = New-Object System.Windows.Forms.TextBox
$txtWin.Location = New-Object System.Drawing.Point(468, 44)
$txtWin.Size = New-Object System.Drawing.Size(150, 26)
$txtWin.Text = '在线学堂'
$panelTop.Controls.Add($txtWin)

$chkTop = New-Object System.Windows.Forms.CheckBox
$chkTop.Text = '窗口置顶'
$chkTop.Location = New-Object System.Drawing.Point(662, 48)
$chkTop.Size = New-Object System.Drawing.Size(90, 24)
$chkTop.Add_Click({ $form.TopMost = $chkTop.Checked })
$panelTop.Controls.Add($chkTop)

Add-Caption $panelTop 'API Key' 6 80 122 | Out-Null
$txtKey = New-Object System.Windows.Forms.TextBox
$txtKey.Location = New-Object System.Drawing.Point(134, 80)
$txtKey.Size = New-Object System.Drawing.Size(460, 26)
$txtKey.UseSystemPasswordChar = $true
$panelTop.Controls.Add($txtKey)

$btnShow = New-FlatButton '显示' 600 79 70 28
$btnShow.Add_Click({ $txtKey.UseSystemPasswordChar = -not $txtKey.UseSystemPasswordChar })
$panelTop.Controls.Add($btnShow)

$grpImage = New-Object System.Windows.Forms.GroupBox
$grpImage.Text = '图片题'
$grpImage.Location = New-Object System.Drawing.Point(8, 112)
$grpImage.Size = New-Object System.Drawing.Size(480, 58)
$panelTop.Controls.Add($grpImage)

$rdoAuto = New-Object System.Windows.Forms.RadioButton
$rdoAuto.Text = '自动看图作答'
$rdoAuto.Checked = $true
$rdoAuto.Location = New-Object System.Drawing.Point(12, 24)
$rdoAuto.Size = New-Object System.Drawing.Size(120, 22)
$grpImage.Controls.Add($rdoAuto)

$rdoSkip = New-Object System.Windows.Forms.RadioButton
$rdoSkip.Text = '跳过'
$rdoSkip.Location = New-Object System.Drawing.Point(170, 24)
$rdoSkip.Size = New-Object System.Drawing.Size(64, 22)
$grpImage.Controls.Add($rdoSkip)

$rdoWait = New-Object System.Windows.Forms.RadioButton
$rdoWait.Text = '停留等我做'
$rdoWait.Location = New-Object System.Drawing.Point(246, 24)
$rdoWait.Size = New-Object System.Drawing.Size(110, 22)
$grpImage.Controls.Add($rdoWait)

$lblImageHint = New-Object System.Windows.Forms.Label
$lblImageHint.Text = '需视觉模型'
$lblImageHint.Location = New-Object System.Drawing.Point(372, 24)
$lblImageHint.Size = New-Object System.Drawing.Size(96, 20)
$lblImageHint.ForeColor = [System.Drawing.Color]::FromArgb(140, 144, 152)
$grpImage.Controls.Add($lblImageHint)

$btnRead = New-FlatButton '读取当前题' 620 120 130 28
$panelTop.Controls.Add($btnRead)

$panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$panelButtons.Location = New-Object System.Drawing.Point(8, 190)
$panelButtons.Size = New-Object System.Drawing.Size(764, 44)
$panelButtons.Anchor = 'Top,Left,Right'
$panelButtons.WrapContents = $false
$panelButtons.Padding = New-Object System.Windows.Forms.Padding(2, 2, 2, 2)
$panelButtons.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($panelButtons)

$btnAuto = New-Object System.Windows.Forms.Button
$btnAuto.Text = '自动作答'
$btnAuto.AutoSize = $true
$btnAuto.MinimumSize = New-Object System.Drawing.Size(96, 32)
$btnAuto.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
$panelButtons.Controls.Add($btnAuto)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = '停止'
$btnStop.AutoSize = $true
$btnStop.MinimumSize = New-Object System.Drawing.Size(72, 32)
$btnStop.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
$btnStop.Enabled = $false
$panelButtons.Controls.Add($btnStop)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = '检查答题卡'
$btnScan.AutoSize = $true
$btnScan.MinimumSize = New-Object System.Drawing.Size(104, 32)
$btnScan.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
$panelButtons.Controls.Add($btnScan)

$btnJump = New-Object System.Windows.Forms.Button
$btnJump.Text = '跳到第一道未答'
$btnJump.AutoSize = $true
$btnJump.MinimumSize = New-Object System.Drawing.Size(124, 32)
$btnJump.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
$panelButtons.Controls.Add($btnJump)

$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = '提交（需确认）'
$btnSubmit.AutoSize = $true
$btnSubmit.MinimumSize = New-Object System.Drawing.Size(120, 32)
$btnSubmit.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
$panelButtons.Controls.Add($btnSubmit)

$btnHelp = New-Object System.Windows.Forms.Button
$btnHelp.Text = '说明 / 反馈'
$btnHelp.AutoSize = $true
$btnHelp.MinimumSize = New-Object System.Drawing.Size(96, 32)
$btnHelp.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
$btnHelp.Add_Click({ Start-Process $script:IssuesUrl })
$panelButtons.Controls.Add($btnHelp)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(8, 240)
$txtLog.Size = New-Object System.Drawing.Size(764, 310)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtLog.Anchor = 'Top,Left,Right,Bottom'
$form.Controls.Add($txtLog)

$lnkUpdate = New-Object System.Windows.Forms.LinkLabel
$lnkUpdate.Text = '当前版本 v' + $script:AppVersion + ' · 项目主页 / 反馈'
$lnkUpdate.Location = New-Object System.Drawing.Point(8, 566)
$lnkUpdate.Size = New-Object System.Drawing.Size(640, 20)
$lnkUpdate.Anchor = 'Left,Bottom'
$lnkUpdate.Add_Click({ Start-Process $script:RepoUrl })
$form.Controls.Add($lnkUpdate)

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
$script:autoImageThisRun = @()

# ---------------------------------------------------------------- 免责声明

$disclaimerText = @'
【使用前请读】

一、正确率会有浮动
答案由你选择的大模型现场生成，不同模型、不同时间可能给出不同答案；多选题偶尔只给一个字母。
本工具只做「读题 → 点选 → 保存」。

二、图片题正确率一般
选项是图片的题目，可以交给模型看图作答（需要支持图像输入的模型），也可以选「跳过」或「停留等我做」。
这类题正确率取决于模型与你，做对与否请自行判断。

三、为爱发电
免费、开源、无广告、不收费，仅限个人辅助学习，不可替代正常学习，不得转售！随缘更新；不承诺分数，也不对成绩负责。

四、提交前请自己核对
程序不会自动提交。点「提交（需确认）」后仍会核对一遍答题卡，最终决定权在用户。

五、判色依赖环境
「已选 / 未选」靠屏幕取色判断，浅色主题 + 75% 缩放下最准。
换主题或缩放后，如果日志里出现「与目标不一致」，请手工确认那一题。

六、安全
API Key 只保存在内存里，不写盘、不上传。不要把账号密码或验证码交给任何人，包括本程序。

七、有问题欢迎反馈
GitHub Issues：请点下方「查看项目主页 / 反馈」
贴日志时记得先把 Key 抹掉。

八、遇到未作答的题目
先点「跳到第一道未答」定位，再点「自动作答」继续；已作答的题不会被重复点选。

九、保持浏览器答题界面无遮挡
作答过程中程序窗口会收进托盘，请让答题页面完整可见；被其它置顶窗口盖住时，屏幕取色判读会读错，点击也会打偏。
'@

$disclaimerText = $disclaimerText + "`r`n当前版本 v" + $script:AppVersion + "。以后有新版本，会在窗口底部提示。"

function Show-Disclaimer {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '使用前请读（v' + $script:AppVersion + '）'
    $dlg.ClientSize = New-Object System.Drawing.Size(720, 560)
    $dlg.StartPosition = 'CenterScreen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $dlg.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.ScrollBars = 'Vertical'
    $box.BorderStyle = 'FixedSingle'
    $box.Location = New-Object System.Drawing.Point(16, 14)
    $box.Size = New-Object System.Drawing.Size(688, 476)
    $box.Text = ($disclaimerText -replace "`r?`n", "`r`n")
    $dlg.Controls.Add($box)
    $dlg.Add_Shown({ $box.SelectionStart = 0; $box.SelectionLength = 0; $box.ScrollToCaret() })
    $dlg.Add_Shown({ [void][ExamWin]::ShowWindow($dlg.Handle, 5); $dlg.Activate() })

    $chkNever = New-Object System.Windows.Forms.CheckBox
    $chkNever.Text = '本版本不再提示'
    $chkNever.Location = New-Object System.Drawing.Point(16, 504)
    $chkNever.Size = New-Object System.Drawing.Size(180, 26)
    if ($script:settings.Contains('disclaimerDismissedVersion')) {
        if ($script:settings['disclaimerDismissedVersion'] -eq $script:AppVersion) { $chkNever.Checked = $true }
    }
    $dlg.Controls.Add($chkNever)

    $btnRepo = New-Object System.Windows.Forms.Button
    $btnRepo.Text = '查看项目主页 / 反馈'
    $btnRepo.Location = New-Object System.Drawing.Point(280, 500)
    $btnRepo.Size = New-Object System.Drawing.Size(180, 32)
    $btnRepo.Add_Click({ Start-Process $script:IssuesUrl })
    $dlg.Controls.Add($btnRepo)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '我知道了'
    $btnOk.Location = New-Object System.Drawing.Point(484, 500)
    $btnOk.Size = New-Object System.Drawing.Size(200, 32)
    $btnOk.DialogResult = 'OK'
    $dlg.Controls.Add($btnOk)

    $dlg.AcceptButton = $btnOk
    $res = $dlg.ShowDialog()
    $accepted = ($res -eq [System.Windows.Forms.DialogResult]::OK)
    if ($accepted -and $chkNever.Checked) {
        $script:settings['disclaimerDismissedVersion'] = $script:AppVersion
        Save-Settings $script:settings
    }
    $dlg.Dispose()
    return $accepted
}

# ---------------------------------------------------------------- 更新提示

function Start-UpdateCheck {
    $urls = @(
        "https://api.github.com/repos/$($script:RepoOwner)/$($script:RepoName)/releases/latest",
        "https://raw.githubusercontent.com/$($script:RepoOwner)/$($script:RepoName)/main/VERSION"
    )
    $script:updateRun = [powershell]::Create()
    [void]$script:updateRun.AddScript({
        param($targets)
        foreach ($u in $targets) {
            try {
                $r = Invoke-RestMethod -Uri $u -TimeoutSec 8 -Headers @{ 'User-Agent' = 'zhihuishu-exam-helper' }
                if ($r -is [string]) { $v = ($r | Out-String).Trim() }
                else { $v = [string]$r.tag_name }
                $v = $v.Trim().TrimStart('v', 'V')
                if ($v) { return $v }
            }
            catch { }
        }
        return $null
    }).AddArgument($urls)
    $script:updateAsync = $script:updateRun.BeginInvoke()
}

function Complete-UpdateCheck {
    if (-not $script:updateAsync) { return }
    if (-not $script:updateAsync.IsCompleted) { return }

    $script:updateTimer.Stop()
    $remote = $null
    try { $remote = $script:updateRun.EndInvoke($script:updateAsync) | Select-Object -First 1 } catch { $remote = $null }
    try { $script:updateRun.Dispose() } catch { }

    if (-not $remote) { Log '更新检查：没有取到远端版本（网络或仓库不可达）'; return }

    $localVer = [version]$script:AppVersion
    $remoteVer = $null
    try { $remoteVer = [version]([string]$remote) } catch { $remoteVer = $null }
    if (-not $remoteVer) { Log ('更新检查：远端版本无法解析（' + $remote + '）'); return }

    if ($remoteVer -gt $localVer) {
        $lnkUpdate.Text = ('发现新版本 v' + $remoteVer.ToString() + '（当前 v' + $script:AppVersion + '）→ 点此查看更新')
        $lnkUpdate.LinkColor = [System.Drawing.Color]::FromArgb(220, 80, 40)
        Log ('发现新版本 v' + $remoteVer.ToString() + '，当前 v' + $script:AppVersion + '，可在项目主页下载')
    }
    else {
        Log ('更新检查：已是最新（远端 v' + $remoteVer.ToString() + '）')
    }
}

$script:updateTimer = New-Object System.Windows.Forms.Timer
$script:updateTimer.Interval = 700
$script:updateStarted = $false
$script:updateTimer.Add_Tick({
    if (-not $script:updateStarted) {
        $script:updateStarted = $true
        Start-UpdateCheck
        return
    }
    Complete-UpdateCheck
})

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
    $script:autoImageThisRun = @()
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
                # 图片题：三种策略
                $mode = 'auto'
                if ($rdoSkip.Checked) { $mode = 'skip' }
                elseif ($rdoWait.Checked) { $mode = 'wait' }

                if ($mode -eq 'wait') {
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
                elseif ($mode -eq 'skip') {
                    $script:skippedThisRun += $q.Num
                    Log ('第 ' + $q.Num + ' 题：图片题，已跳过')
                }
                else {
                    $png = Get-QuestionImage $txtWin.Text $q
                    if (-not $png) {
                        $script:skippedThisRun += $q.Num
                        Log ('第 ' + $q.Num + ' 题：图片题截图失败，跳过')
                    }
                    else {
                        try {
                            $ans = Ask-Model-Vision $txtBase.Text $txtModel.Text $txtKey.Text $q $png
                            if (-not $ans.Letters) {
                                $script:skippedThisRun += $q.Num
                                Log ('第 ' + $q.Num + ' 题：模型没给出字母（' + $ans.Raw + '），跳过')
                            }
                            else {
                                $res = Apply-Answer $doc $ans.Letters $txtWin.Text
                                $script:autoImageThisRun += $q.Num
                                $script:answeredThisRun += $q.Num
                                $now = '空'
                                if ($res -and $res.After) { $now = $res.After }
                                if ($res -and $res.Ok) {
                                    Log ('第 ' + $q.Num + ' 题 [图片题] 模型选 ' + $ans.Letters + '（现 ' + $now + '）')
                                }
                                else {
                                    Log ('第 ' + $q.Num + ' 题 [图片题] 模型选 ' + $ans.Letters + '，点击后实际 ' + $now + '，请人工确认')
                                }
                            }
                        }
                        catch {
                            $script:skippedThisRun += $q.Num
                            Log ('第 ' + $q.Num + ' 题：图片问答失败（模型可能不支持图像输入）- ' + $_.Exception.Message)
                        }
                    }
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
        if ($script:autoImageThisRun.Count -gt 0) { $msg += '（其中图片题 ' + $script:autoImageThisRun.Count + ' 道自动作答）' }
        if ($script:skippedThisRun.Count -gt 0) { $msg += '；未作答 ' + ($script:skippedThisRun -join ', ') }
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

$form.Add_Shown({
    [void][ExamWin]::ShowWindow($form.Handle, 5)
    $form.Activate()
    $script:updateTimer.Start()
    if ($script:settings.Contains('baseUrl') -and $script:settings['baseUrl']) { $txtBase.Text = [string]$script:settings['baseUrl'] }
    if ($script:settings.Contains('model') -and $script:settings['model']) { $txtModel.Text = [string]$script:settings['model'] }
    if ($script:settings.Contains('winMatch') -and $script:settings['winMatch']) { $txtWin.Text = [string]$script:settings['winMatch'] }
    if ($script:settings.Contains('imageMode')) {
        switch ([string]$script:settings['imageMode']) {
            'skip' { $rdoSkip.Checked = $true }
            'wait' { $rdoWait.Checked = $true }
            default { $rdoAuto.Checked = $true }
        }
    }
    if ($script:settings.Contains('topMost') -and $script:settings['topMost']) { $chkTop.Checked = $true; $form.TopMost = $true }
    Log ('智慧树单元测试助手 v' + $script:AppVersion + ' 已启动')
    Log '提醒：本程序不会自动提交；正确率随模型浮动，图片题建议使用支持图像输入的模型。'
})

$form.Add_FormClosing({
    try {
        $script:settings['baseUrl'] = $txtBase.Text
        $script:settings['model'] = $txtModel.Text
        $script:settings['winMatch'] = $txtWin.Text
        if ($rdoSkip.Checked) { $script:settings['imageMode'] = 'skip' }
        elseif ($rdoWait.Checked) { $script:settings['imageMode'] = 'wait' }
        else { $script:settings['imageMode'] = 'auto' }
        $script:settings['topMost'] = [bool]$chkTop.Checked
        $script:settings['windowX'] = $form.Left
        $script:settings['windowY'] = $form.Top
        Save-Settings $script:settings
    }
    catch { }
    try { $notify.Visible = $false; $notify.Dispose() } catch { }
})

$form.Add_Resize({
    if ($form.WindowState -eq 'Minimized') { $form.Hide() }
})

# ---------------------------------------------------------------- 启动

$dismissed = $false
if ($script:settings.Contains('disclaimerDismissedVersion')) {
    if ($script:settings['disclaimerDismissedVersion'] -eq $script:AppVersion) { $dismissed = $true }
}

if (-not $dismissed) {
    $accepted = Show-Disclaimer
    if (-not $accepted) { exit }
}

[void]$form.ShowDialog()
