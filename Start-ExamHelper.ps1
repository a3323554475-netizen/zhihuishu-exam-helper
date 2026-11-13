#requires -Version 5.1
<#
  Start-ExamHelper.ps1 - 智慧树/在线学堂单元测试助手（公开版）

  - UI Automation 读取考试页面的题干与选项
  - 调用任意 OpenAI 兼容接口生成答案
  - 图片选项题跳过，留待人工
  - 一套题填完即停，人工检查答题卡后再手动提交

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

# ---------------------------------------------------------------- 基础动作

function Nap([int]$min, [int]$max) {
    if ($max -le $min) { Start-Sleep -Milliseconds $min; return $min }
    $ms = Get-Random -Minimum $min -Maximum $max
    Start-Sleep -Milliseconds $ms
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

function Focus-Window([string]$match) {
    $win = Get-BrowserWindow $match
    if (-not $win) { return $false }
    if (-not ('ExamFocus' -as [type])) {
        Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class ExamFocus{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);public static void Go(IntPtr h){ShowWindow(h,9);SetForegroundWindow(h);}}'
    }
    [ExamFocus]::Go([IntPtr]$win.Current.NativeWindowHandle)
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
    if (-not ('ExamMouse' -as [type])) {
        Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class ExamMouse{[DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);[DllImport("user32.dll")]public static extern void mouse_event(uint f,uint dx,uint dy,uint d,IntPtr e);public static void Click(int x,int y){SetCursorPos(x,y);System.Threading.Thread.Sleep(90);mouse_event(0x0002,0,0,0,IntPtr.Zero);System.Threading.Thread.Sleep(50);mouse_event(0x0004,0,0,0,IntPtr.Zero);}}'
    }
    $cx = [int]($r.X + $r.Width / 2)
    $cy = [int]($r.Y + $r.Height / 2)
    [ExamMouse]::Click($cx, $cy)
    return ('mouse@' + $cx + ',' + $cy)
}

function Click-ByName($doc, [string]$name) {
    foreach ($it in (Get-Items $doc)) {
        if ($it.T -eq 'Button' -and $it.N -eq $name) { return (Invoke-Element $it.E) }
    }
    return 'no-button'
}

function Click-Option($doc, [string]$letter) {
    $items = Get-Items $doc
    for ($i = 0; $i -lt $items.Count; $i++) {
        $it = $items[$i]
        if ($it.T -ne 'Text') { continue }
        if ($it.N -ne ($letter + '.')) { continue }
        $target = $it
        if ($i + 1 -lt $items.Count -and $items[$i + 1].T -eq 'Text' -and $items[$i + 1].N -notmatch '^[A-D]\.$') {
            $target = $items[$i + 1]
        }
        return (Invoke-Element $target.E)
    }
    return ('no-option-' + $letter)
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
    foreach ($m in [regex]::Matches($text.ToUpper(), '[A-D]')) { if ($letters -notlike ('*' + $m.Value + '*')) { $letters += $m.Value } }
    return [pscustomobject]@{ Letters = $letters; Raw = $text }
}

# ---------------------------------------------------------------- 界面

$form = New-Object System.Windows.Forms.Form
$form.Text = '智慧树单元测试助手'
$form.Size = New-Object System.Drawing.Size(760, 620)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true

function New-Label($text, $x, $y, $w) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size($w, 20)
    return $l
}

$form.Controls.Add((New-Label 'API Base URL' 16 18 100))
$txtBase = New-Object System.Windows.Forms.TextBox
$txtBase.Location = New-Object System.Drawing.Point(120, 15)
$txtBase.Size = New-Object System.Drawing.Size(600, 24)
$txtBase.Text = 'https://api.deepseek.com/v1'
$form.Controls.Add($txtBase)

$form.Controls.Add((New-Label 'Model' 16 52 100))
$txtModel = New-Object System.Windows.Forms.TextBox
$txtModel.Location = New-Object System.Drawing.Point(120, 49)
$txtModel.Size = New-Object System.Drawing.Size(280, 24)
$txtModel.Text = 'deepseek-chat'
$form.Controls.Add($txtModel)

$form.Controls.Add((New-Label '窗口标题片段' 420 52 120))
$txtWin = New-Object System.Windows.Forms.TextBox
$txtWin.Location = New-Object System.Drawing.Point(540, 49)
$txtWin.Size = New-Object System.Drawing.Size(180, 24)
$txtWin.Text = '在线学堂'
$form.Controls.Add($txtWin)

$form.Controls.Add((New-Label 'API Key' 16 86 100))
$txtKey = New-Object System.Windows.Forms.TextBox
$txtKey.Location = New-Object System.Drawing.Point(120, 83)
$txtKey.Size = New-Object System.Drawing.Size(460, 24)
$txtKey.UseSystemPasswordChar = $true
$form.Controls.Add($txtKey)

$chkShow = New-Object System.Windows.Forms.CheckBox
$chkShow.Text = '显示'
$chkShow.Location = New-Object System.Drawing.Point(590, 85)
$chkShow.Size = New-Object System.Drawing.Size(60, 22)
$chkShow.Add_Click({ $txtKey.UseSystemPasswordChar = -not $chkShow.Checked })
$form.Controls.Add($chkShow)

$btnRead = New-Object System.Windows.Forms.Button
$btnRead.Text = '读取当前题'
$btnRead.Location = New-Object System.Drawing.Point(16, 122)
$btnRead.Size = New-Object System.Drawing.Size(120, 32)
$form.Controls.Add($btnRead)

$btnAuto = New-Object System.Windows.Forms.Button
$btnAuto.Text = '自动作答（跳过图片题）'
$btnAuto.Location = New-Object System.Drawing.Point(146, 122)
$btnAuto.Size = New-Object System.Drawing.Size(190, 32)
$form.Controls.Add($btnAuto)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = '停止'
$btnStop.Location = New-Object System.Drawing.Point(346, 122)
$btnStop.Size = New-Object System.Drawing.Size(80, 32)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = '提交（需确认）'
$btnSubmit.Location = New-Object System.Drawing.Point(436, 122)
$btnSubmit.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($btnSubmit)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(16, 166)
$txtLog.Size = New-Object System.Drawing.Size(710, 400)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

function Log([string]$msg) {
    $line = ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $msg)
    $txtLog.AppendText($line + "`r`n")
    [System.Windows.Forms.Application]::DoEvents()
}

$script:running = $false

# ---------------------------------------------------------------- 事件：读取当前题

$btnRead.Add_Click({
    try {
        $doc = Get-ExamDoc $txtWin.Text
        if (-not $doc) { Log '找不到浏览器窗口，请确认浏览器已打开且标题包含该片段'; return }
        $q = Read-Question $doc
        if (-not $q) { Log '页面里没读到题号，可能不在考试页或页面文本不透出'; return }
        Log ('第 ' + $q.Num + ' 题 [' + $q.Type + '] ' + $q.Stem)
        foreach ($k in $q.Options.Keys) { Log ('  ' + $k + '. ' + $q.Options[$k]) }
        if (-not $q.HasText) { Log '  → 这道题的选项是图片，程序会跳过它' }
    }
    catch { Log ('读取失败：' + $_.Exception.Message) }
})

# ---------------------------------------------------------------- 事件：自动作答

$btnAuto.Add_Click({
    if (-not $txtKey.Text) { Log '先填 API Key'; return }

    $script:running = $true
    $btnAuto.Enabled = $false
    $btnStop.Enabled = $true

    $answered = 0
    $skipped = @()
    $lastNum = 0

    try {
        [void](Focus-Window $txtWin.Text)
        Nap 600 1000 | Out-Null

        $limit = 80
        for ($step = 0; $step -lt $limit; $step++) {
            if (-not $script:running) { Log '已停止'; break }
            if ($step -gt 0 -and ($step % 9) -eq 0) {
                [void](Focus-Window $txtWin.Text)
                $p = Nap 2000 4000
                Log ('  停顿 ' + $p + 'ms')
            }

            $doc = Get-ExamDoc $txtWin.Text
            if (-not $doc) { Log '窗口丢了，停止'; break }

            $q = $null
            for ($try = 0; $try -lt 6; $try++) {
                $q = Read-Question $doc
                if ($q) { break }
                Nap 350 450 | Out-Null
                $doc = Get-ExamDoc $txtWin.Text
            }
            if (-not $q) { Log '读不到题目，停止'; break }
            if ($q.Num -le $lastNum) { Log ('题号没有前进（' + $q.Num + '），停止'); break }
            $lastNum = $q.Num

            if (-not $q.HasText) {
                $skipped += $q.Num
                Log ('第 ' + $q.Num + ' 题：图片选项，跳过（请稍后自己看图作答）')
            }
            else {
                try {
                    $ans = Ask-Model $txtBase.Text $txtModel.Text $txtKey.Text $q
                    if (-not $ans.Letters) {
                        Log ('第 ' + $q.Num + ' 题：模型没给出字母，跳过。原文=' + $ans.Raw)
                    }
                    else {
                        $picked = ''
                        foreach ($c in $ans.Letters.ToCharArray()) {
                            [void](Click-Option $doc ([string]$c))
                            $picked += $c
                            Nap 260 620 | Out-Null
                        }
                        $answered++
                        Log ('第 ' + $q.Num + ' 题 [' + $q.Type + '] 选 ' + $picked)
                    }
                }
                catch {
                    Log ('第 ' + $q.Num + ' 题：模型调用失败 - ' + $_.Exception.Message)
                }
            }

            Nap 200 400 | Out-Null
            $clicked = Click-ByName $doc '下一题'
            if ($clicked -eq 'no-button') { $clicked = Click-ByName $doc '保存' }
            Nap 900 1800 | Out-Null
        }

        $msg = '本套填答结束：已作答 ' + $answered + ' 题'
        if ($skipped.Count -gt 0) { $msg += '；跳过图片题 ' + ($skipped -join ', ') }
        $msg += "`r`n`r`n现在请切到浏览器看右侧答题卡：蓝色=已答，黑色=未答。补完后回到这里点「提交（需确认）」。"
        Log $msg
        [void][System.Windows.Forms.MessageBox]::Show($msg, '请人工检查后再提交')
    }
    catch {
        Log ('中断：' + $_.Exception.Message)
    }
    finally {
        $script:running = $false
        $btnAuto.Enabled = $true
        $btnStop.Enabled = $false
    }
})

$btnStop.Add_Click({
    $script:running = $false
    Log '收到停止指令'
})

# ---------------------------------------------------------------- 事件：提交

$btnSubmit.Add_Click({
    $yes = [System.Windows.Forms.MessageBox]::Show('确认提交？提交后这份试卷不能再修改。', '二次确认', 'YesNo', 'Warning')
    if ($yes -ne 'Yes') { Log '已取消提交'; return }
    try {
        [void](Focus-Window $txtWin.Text)
        Nap 400 700 | Out-Null
        $doc = Get-ExamDoc $txtWin.Text
        if (-not $doc) { Log '找不到考试页面'; return }
        Log ('提交作业 → ' + (Click-ByName $doc '提交作业'))
        Nap 1200 1800 | Out-Null
        $doc = Get-ExamDoc $txtWin.Text
        Log ('确认 → ' + (Click-ByName $doc '确定'))
    }
    catch { Log ('提交失败：' + $_.Exception.Message) }
})

[void]$form.ShowDialog()
