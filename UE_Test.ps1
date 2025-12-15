Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ======================
# Global variables
# ======================
$global:psList      = @()   # iperf3 PowerShell window PID list
$global:PingJobUE1  = $null
$global:PingJobUE2  = $null
$global:RebootTimer      = $null
$global:RebootDetectTry  = 0
$global:PingTickBusy = $false

# ======================
# Config file path
# ======================
if ($PSCommandPath) {
    $scriptDir = Split-Path -Parent $PSCommandPath
} else {
    $scriptDir = (Get-Location).Path
}
$global:ConfigFilePath = Join-Path $scriptDir "UE_Test_Config.json"

# UE 정보 구조체: IP + Modem 정보
$global:UE1Info = [PSCustomObject]@{
    IP        = $null
    ModemName = $null
    ComPort   = $null
}
$global:UE2Info = [PSCustomObject]@{
    IP        = $null
    ModemName = $null
    ComPort   = $null
}

# ======================
# Colors
# ======================
$colorDefaultLog = [System.Drawing.Color]::White
$colorUE1        = [System.Drawing.Color]::LightGreen
$colorUE2        = [System.Drawing.Color]::LightSkyBlue

# ======================
# Main Form
# ======================
$form = New-Object System.Windows.Forms.Form
$form.Text = "my5G UE Test Tool"
$form.ClientSize = New-Object System.Drawing.Size(1000, 720)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$form.ForeColor = [System.Drawing.Color]::White
$form.AutoScroll = $true

# 공통 스타일
$labelColor    = [System.Drawing.Color]::White
$textBgColor   = [System.Drawing.Color]::FromArgb(45,45,48)
$textFgColor   = [System.Drawing.Color]::White
$buttonBgColor = [System.Drawing.Color]::FromArgb(64,64,64)
$buttonFgColor = [System.Drawing.Color]::White

function Set-DarkTextBoxStyle($tb) {
    $tb.BackColor = $textBgColor
    $tb.ForeColor = $textFgColor
    $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
}
function Set-DarkLabelStyle($lb) {
    $lb.ForeColor = $labelColor
    $lb.BackColor = [System.Drawing.Color]::Transparent
}
function Set-DarkButtonStyle($btn) {
    $btn.BackColor = $buttonBgColor
    $btn.ForeColor = $buttonFgColor
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
function Set-DarkComboBoxStyle($cb) {
    $cb.BackColor = $textBgColor
    $cb.ForeColor = $textFgColor
    $cb.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
function Set-DarkCheckBoxStyle($cb) {
    $cb.ForeColor = $labelColor
    $cb.BackColor = [System.Drawing.Color]::Transparent
}

# ======================
# 좌측 상단: Device Info + Swap
# ======================
[int]$labelX = 20
[int]$inputX = 190
[int]$widthLabel = 160
[int]$widthInput = 110
[int]$height = 23
[int]$gap = 8

[int]$y = 10

# --- Device Info + UE List 만 그룹박스로 묶기 ---
$grpDevice = New-Object System.Windows.Forms.GroupBox
$grpDevice.Text = "Device Info"
$grpDevice.Location = New-Object System.Drawing.Point(10, $y)
$grpDevice.Size = New-Object System.Drawing.Size(970, 65)   # 필요하면 높이 조절
$grpDevice.ForeColor = $labelColor
$form.Controls.Add($grpDevice)

# 그룹 내부 좌표
[int]$gY = 20
[int]$gX = 10

$lblDeviceUE1 = New-Object System.Windows.Forms.Label
$lblDeviceUE1.Location = New-Object System.Drawing.Point($gX, $gY)

# 기존: 400 → 그룹 가로(930) - 여백 조금(40) = 890 정도
$lblDeviceUE1.Size = New-Object System.Drawing.Size(890, $height)
$lblDeviceUE1.ForeColor = $colorUE1
$lblDeviceUE1.BackColor = [System.Drawing.Color]::Transparent
$grpDevice.Controls.Add($lblDeviceUE1)

$gY += $height

$lblDeviceUE2 = New-Object System.Windows.Forms.Label
$lblDeviceUE2.Location = New-Object System.Drawing.Point($gX, $gY)
$lblDeviceUE2.Size = New-Object System.Drawing.Size(890, $height)
$lblDeviceUE2.ForeColor = $colorUE2
$lblDeviceUE2.BackColor = [System.Drawing.Color]::Transparent
$grpDevice.Controls.Add($lblDeviceUE2)

# 그룹박스 아래부터는 기존처럼 폼에 직접 배치
$y = $grpDevice.Bottom + 20

# Swap 버튼
$btnSwapUE = New-Object System.Windows.Forms.Button
$btnSwapUE.Text = "Swap UE1 <-> UE2"
$btnSwapUE.Location = New-Object System.Drawing.Point($labelX, $y)
$btnSwapUE.Size = New-Object System.Drawing.Size(160, 28)
Set-DarkButtonStyle $btnSwapUE
$form.Controls.Add($btnSwapUE)

$y += $height + 20

# iPerf Server IP
$labelServerIp = New-Object System.Windows.Forms.Label
$labelServerIp.Text = "iPerf Server IP"
$labelServerIp.Location = New-Object System.Drawing.Point($labelX, $y)
$labelServerIp.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelServerIp
$form.Controls.Add($labelServerIp)

$textServerIp = New-Object System.Windows.Forms.TextBox
$textServerIp.Location = New-Object System.Drawing.Point($inputX, $y)
$textServerIp.Size = New-Object System.Drawing.Size($widthInput, $height)
$textServerIp.Text = "10.36.10.250"
Set-DarkTextBoxStyle $textServerIp
$form.Controls.Add($textServerIp)

$y += $height + (4 * $gap)

# Route Buttons
$buttonRouteAdmin = New-Object System.Windows.Forms.Button
$buttonRouteAdmin.Text = "Add Routing"
$buttonRouteAdmin.Location = New-Object System.Drawing.Point($labelX, $y)
$buttonRouteAdmin.Size = New-Object System.Drawing.Size(200, 30)
Set-DarkButtonStyle $buttonRouteAdmin
$form.Controls.Add($buttonRouteAdmin)

$buttonRouteDelete = New-Object System.Windows.Forms.Button
$buttonRouteDelete.Text = "Del Routing"
$buttonRouteDelete.Location = New-Object System.Drawing.Point(($labelX + 220), $y)
$buttonRouteDelete.Size = New-Object System.Drawing.Size(200, 30)
Set-DarkButtonStyle $buttonRouteDelete
$form.Controls.Add($buttonRouteDelete)

$y += $height + (4 * $gap)

# ============================
# DL / UL Ports (2열: UE1 / UE2)
# ============================

# UE 컬럼 타이틀을 체크박스로 사용 (Enable UE1/UE2 역할)
$checkUE1 = New-Object System.Windows.Forms.CheckBox
$checkUE1.Text = "Enable UE1"
$checkUE1.Location = New-Object System.Drawing.Point($inputX, $y)
$checkUE1.Size = New-Object System.Drawing.Size($widthInput, $height)
$checkUE1.Checked = $true
Set-DarkCheckBoxStyle $checkUE1
$form.Controls.Add($checkUE1)

$colUe2X = $inputX + $widthInput + 10

$checkUE2 = New-Object System.Windows.Forms.CheckBox
$checkUE2.Text = "Enable UE2"
$checkUE2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$checkUE2.Size = New-Object System.Drawing.Size($widthInput, $height)
$checkUE2.Checked = $true
Set-DarkCheckBoxStyle $checkUE2
$form.Controls.Add($checkUE2)

$y += $height + $gap

# ---- DL only row ----
$checkDlOnlyUe1 = New-Object System.Windows.Forms.CheckBox
$checkDlOnlyUe1.Text = "DL only"
$checkDlOnlyUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$checkDlOnlyUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
$checkDlOnlyUe1.Checked = $false
Set-DarkCheckBoxStyle $checkDlOnlyUe1
$form.Controls.Add($checkDlOnlyUe1)

$checkDlOnlyUe2 = New-Object System.Windows.Forms.CheckBox
$checkDlOnlyUe2.Text = "DL only"
$checkDlOnlyUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$checkDlOnlyUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$checkDlOnlyUe2.Checked = $false
Set-DarkCheckBoxStyle $checkDlOnlyUe2
$form.Controls.Add($checkDlOnlyUe2)

$y += $height + $gap

# ---- UL only row ----
$checkUlOnlyUe1 = New-Object System.Windows.Forms.CheckBox
$checkUlOnlyUe1.Text = "UL only"
$checkUlOnlyUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$checkUlOnlyUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
$checkUlOnlyUe1.Checked = $false
Set-DarkCheckBoxStyle $checkUlOnlyUe1
$form.Controls.Add($checkUlOnlyUe1)

$checkUlOnlyUe2 = New-Object System.Windows.Forms.CheckBox
$checkUlOnlyUe2.Text = "UL only"
$checkUlOnlyUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$checkUlOnlyUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$checkUlOnlyUe2.Checked = $false
Set-DarkCheckBoxStyle $checkUlOnlyUe2
$form.Controls.Add($checkUlOnlyUe2)

$y += $height + (2 * $gap)

# ---- DL Port row ----
$labelDlPort = New-Object System.Windows.Forms.Label
$labelDlPort.Text = "DL Port #"
$labelDlPort.Location = New-Object System.Drawing.Point($labelX, $y)
$labelDlPort.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelDlPort
$form.Controls.Add($labelDlPort)

$textDlPortUe1 = New-Object System.Windows.Forms.TextBox
$textDlPortUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$textDlPortUe1.Size = New-Object System.Drawing.Size($widthInput, $height)   # ←  widthInput 사용
$textDlPortUe1.Text = "5501"
Set-DarkTextBoxStyle $textDlPortUe1
$form.Controls.Add($textDlPortUe1)

$textDlPortUe2 = New-Object System.Windows.Forms.TextBox
$textDlPortUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)      # ← 오른쪽 열로 이동
$textDlPortUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textDlPortUe2.Text = "5601"
Set-DarkTextBoxStyle $textDlPortUe2
$form.Controls.Add($textDlPortUe2)

$y += $height + $gap
# ---- UL Port row ----
$labelUlPort = New-Object System.Windows.Forms.Label
$labelUlPort.Text = "UL Port #"
$labelUlPort.Location = New-Object System.Drawing.Point($labelX, $y)
$labelUlPort.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelUlPort
$form.Controls.Add($labelUlPort)

$textUlPortUe1 = New-Object System.Windows.Forms.TextBox
$textUlPortUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$textUlPortUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUlPortUe1.Text = "5502"
Set-DarkTextBoxStyle $textUlPortUe1
$form.Controls.Add($textUlPortUe1)

$textUlPortUe2 = New-Object System.Windows.Forms.TextBox
$textUlPortUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$textUlPortUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUlPortUe2.Text = "5602"
Set-DarkTextBoxStyle $textUlPortUe2
$form.Controls.Add($textUlPortUe2)

$y += $height + $gap

# Download BW
$labelDlBw = New-Object System.Windows.Forms.Label
$labelDlBw.Text = "Download BW"
$labelDlBw.Location = New-Object System.Drawing.Point($labelX, $y)
$labelDlBw.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelDlBw
$form.Controls.Add($labelDlBw)

$textDlBwUe1 = New-Object System.Windows.Forms.TextBox
$textDlBwUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$textDlBwUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
$textDlBwUe1.Text = "1000M"
Set-DarkTextBoxStyle $textDlBwUe1
$form.Controls.Add($textDlBwUe1)

$textDlBwUe2 = New-Object System.Windows.Forms.TextBox
$textDlBwUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$textDlBwUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textDlBwUe2.Text = "1000M"
Set-DarkTextBoxStyle $textDlBwUe2
$form.Controls.Add($textDlBwUe2)

$y += $height + $gap

# Upload BW
$labelUlBw = New-Object System.Windows.Forms.Label
$labelUlBw.Text = "Upload BW"
$labelUlBw.Location = New-Object System.Drawing.Point($labelX, $y)
$labelUlBw.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelUlBw
$form.Controls.Add($labelUlBw)

$textUlBwUe1 = New-Object System.Windows.Forms.TextBox
$textUlBwUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$textUlBwUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUlBwUe1.Text = "500M"
Set-DarkTextBoxStyle $textUlBwUe1
$form.Controls.Add($textUlBwUe1)

$textUlBwUe2 = New-Object System.Windows.Forms.TextBox
$textUlBwUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$textUlBwUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUlBwUe2.Text = "500M"
Set-DarkTextBoxStyle $textUlBwUe2
$form.Controls.Add($textUlBwUe2)

$y += $height + $gap

# Duration
$labelTime = New-Object System.Windows.Forms.Label
$labelTime.Text = "Duration (sec)"
$labelTime.Location = New-Object System.Drawing.Point($labelX, $y)
$labelTime.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelTime
$form.Controls.Add($labelTime)

$textTimeUe1 = New-Object System.Windows.Forms.TextBox
$textTimeUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$textTimeUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
$textTimeUe1.Text = "0"
Set-DarkTextBoxStyle $textTimeUe1
$form.Controls.Add($textTimeUe1)

$textTimeUe2 = New-Object System.Windows.Forms.TextBox
$textTimeUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$textTimeUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textTimeUe2.Text = "0"
Set-DarkTextBoxStyle $textTimeUe2
$form.Controls.Add($textTimeUe2)

$y += $height + $gap

# Protocol
$labelProto = New-Object System.Windows.Forms.Label
$labelProto.Text = "Protocol"
$labelProto.Location = New-Object System.Drawing.Point($labelX, $y)
$labelProto.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelProto
$form.Controls.Add($labelProto)

$comboProtoUe1 = New-Object System.Windows.Forms.ComboBox
$comboProtoUe1.Location = New-Object System.Drawing.Point($inputX, $y)
$comboProtoUe1.Size = New-Object System.Drawing.Size($widthInput, $height)
[void]$comboProtoUe1.Items.Add("UDP")
[void]$comboProtoUe1.Items.Add("TCP")
$comboProtoUe1.SelectedIndex = 0
Set-DarkComboBoxStyle $comboProtoUe1
$form.Controls.Add($comboProtoUe1)

$comboProtoUe2 = New-Object System.Windows.Forms.ComboBox
$comboProtoUe2.Location = New-Object System.Drawing.Point($colUe2X, $y)
$comboProtoUe2.Size = New-Object System.Drawing.Size($widthInput, $height)
[void]$comboProtoUe2.Items.Add("UDP")
[void]$comboProtoUe2.Items.Add("TCP")
$comboProtoUe2.SelectedIndex = 0
Set-DarkComboBoxStyle $comboProtoUe2
$form.Controls.Add($comboProtoUe2)

$y += $height + (2 * $gap)

# iperf Buttons (버튼 생성만)
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Start Selected UEs"
$buttonStart.Location = New-Object System.Drawing.Point(50, $y)
$buttonStart.Size = New-Object System.Drawing.Size(160, 30)
Set-DarkButtonStyle $buttonStart
$form.Controls.Add($buttonStart)

$buttonStop = New-Object System.Windows.Forms.Button
$buttonStop.Text = "Stop All Sessions"
$buttonStop.Location = New-Object System.Drawing.Point(250, $y)
$buttonStop.Size = New-Object System.Drawing.Size(160, 30)
Set-DarkButtonStyle $buttonStop
$form.Controls.Add($buttonStop)

# ---- 여기 아래에 Save / Load 버튼 추가 ----
$y += $height + 2*$gap + 10

$buttonSaveCfg = New-Object System.Windows.Forms.Button
$buttonSaveCfg.Text = "Save Config"
$buttonSaveCfg.Location = New-Object System.Drawing.Point(50, $y)
$buttonSaveCfg.Size = New-Object System.Drawing.Size(160, 30)
Set-DarkButtonStyle $buttonSaveCfg
$form.Controls.Add($buttonSaveCfg)

$buttonLoadCfg = New-Object System.Windows.Forms.Button
$buttonLoadCfg.Text = "Load Config"
$buttonLoadCfg.Location = New-Object System.Drawing.Point(250, $y)
$buttonLoadCfg.Size = New-Object System.Drawing.Size(160, 30)
Set-DarkButtonStyle $buttonLoadCfg
$form.Controls.Add($buttonLoadCfg)

$y += $height + 2*$gap + 30

$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Exit"
$buttonClose.Location = New-Object System.Drawing.Point(180, $y)
$buttonClose.Size = New-Object System.Drawing.Size(120, 30)
Set-DarkButtonStyle $buttonClose
$form.Controls.Add($buttonClose)

# ======================
# Helper: Remote NDIS IP list
# ======================
function Get-NdisIpList {
    $ips = @()
    try {
        $cfgs = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue |
                Where-Object { $_.IPEnabled }
        $adapters = Get-WmiObject Win32_NetworkAdapter -ErrorAction SilentlyContinue
        $adapterByIndex = @{}
        foreach ($ad in $adapters) {
            $adapterByIndex[$ad.Index] = $ad
        }

        foreach ($cfg in $cfgs) {
            $ad = $adapterByIndex[$cfg.Index]
            if (-not $ad) { continue }
            if ($ad.Description -notlike "*Remote NDIS Compatible Device*") { continue }

            $ipv4s = $cfg.IPAddress | Where-Object {
                $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notlike "169.254.*"
            }
            foreach ($ip in $ipv4s) {
                if ($ips -notcontains $ip) { $ips += $ip }
            }
        }
    } catch {}
    return $ips
}

# ======================
# Helper: Modem list
# ======================
function Get-ModemPorts {
    try {
        $modems = Get-CimInstance Win32_POTSModem -ErrorAction Stop
    } catch {
        $modems = Get-WmiObject Win32_POTSModem -ErrorAction SilentlyContinue
    }
    if (-not $modems) { return @() }
    $modems = $modems | Where-Object { $_.AttachedTo -match "^COM\d+" }
    return $modems
}

# ======================
# Helper: Route info for IP
# ======================
function Get-RouteInfoForIp {
    param([string]$IpAddress)

    $result = [PSCustomObject]@{
        Gateway = $null
        IfIndex = $null
    }
    if (-not $IpAddress) { return $result }

    try {
        $cfgs = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue |
                Where-Object { $_.IPEnabled }
        foreach ($cfg in $cfgs) {
            if ($cfg.IPAddress -and ($cfg.IPAddress -contains $IpAddress)) {
                if ($cfg.DefaultIPGateway -and $cfg.DefaultIPGateway.Length -gt 0) {
                    $result.Gateway = $cfg.DefaultIPGateway[0]
                }
                if ($cfg.InterfaceIndex) {
                    $result.IfIndex = $cfg.InterfaceIndex
                }
                return $result
            }
        }
    } catch {}
    return $result
}

# ======================
# 오른쪽: AT Command & Ping
# ======================
[int]$baseX = 480
[int]$y2 = 110   # 약간 아래에서 시작

# AT Command Group
$grpAt = New-Object System.Windows.Forms.GroupBox
$grpAt.Text = "AT Command"
$grpAt.Location = New-Object System.Drawing.Point($baseX, $y2)
$grpAt.Size = New-Object System.Drawing.Size(500, 320)
$grpAt.ForeColor = $labelColor
$form.Controls.Add($grpAt)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target UEs:"
$lblTarget.AutoSize = $true
$lblTarget.Location = New-Object System.Drawing.Point(15, 25)
$grpAt.Controls.Add($lblTarget)

$checkAtUE1 = New-Object System.Windows.Forms.CheckBox
$checkAtUE1.Text = "UE1"
$checkAtUE1.Location = New-Object System.Drawing.Point(100, 20)
$checkAtUE1.Checked = $true
Set-DarkCheckBoxStyle $checkAtUE1
$grpAt.Controls.Add($checkAtUE1)

$checkAtUE2 = New-Object System.Windows.Forms.CheckBox
$checkAtUE2.Text = "UE2"
$checkAtUE2.Location = New-Object System.Drawing.Point(220, 20)
$checkAtUE2.Checked = $true
Set-DarkCheckBoxStyle $checkAtUE2
$grpAt.Controls.Add($checkAtUE2)

$lblCmd = New-Object System.Windows.Forms.Label
$lblCmd.Text = "Command"
$lblCmd.AutoSize = $true
$lblCmd.Location = New-Object System.Drawing.Point(15, 55)
$grpAt.Controls.Add($lblCmd)

$txtCmd = New-Object System.Windows.Forms.TextBox
$txtCmd.Location = New-Object System.Drawing.Point(15, 75)
$txtCmd.Size = New-Object System.Drawing.Size(230, 23)
$txtCmd.BackColor = $textBgColor
$txtCmd.ForeColor = $textFgColor
$grpAt.Controls.Add($txtCmd)

$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text = "Send"
$btnSend.Location = New-Object System.Drawing.Point(255, 73)
$btnSend.Size = New-Object System.Drawing.Size(60, 26)
Set-DarkButtonStyle $btnSend
$grpAt.Controls.Add($btnSend)

# --- 여기: CFUN / Reboot 버튼 3개를 한 줄로 배치 ---
$btnCFUN0 = New-Object System.Windows.Forms.Button
$btnCFUN0.Text = "CFUN=0"
$btnCFUN0.Location = New-Object System.Drawing.Point(15, 105)
$btnCFUN0.Size = New-Object System.Drawing.Size(80, 26)
Set-DarkButtonStyle $btnCFUN0
$grpAt.Controls.Add($btnCFUN0)

$btnCFUN1 = New-Object System.Windows.Forms.Button
$btnCFUN1.Text = "CFUN=1"
$btnCFUN1.Location = New-Object System.Drawing.Point(105, 105)
$btnCFUN1.Size = New-Object System.Drawing.Size(80, 26)
Set-DarkButtonStyle $btnCFUN1
$grpAt.Controls.Add($btnCFUN1)

$btnReboot = New-Object System.Windows.Forms.Button
$btnReboot.Text = "Reboot"
$btnReboot.Location = New-Object System.Drawing.Point(195, 105)
$btnReboot.Size = New-Object System.Drawing.Size(80, 26)
Set-DarkButtonStyle $btnReboot
$grpAt.Controls.Add($btnReboot)

# Log 라벨/창은 아래쪽으로 조금 내림
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.AutoSize = $true
$lblLog.Location = New-Object System.Drawing.Point(15, 140)
$grpAt.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 160)
$txtLog.Size = New-Object System.Drawing.Size(450, 140)
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.BackColor = $textBgColor
$txtLog.ForeColor = $textFgColor
$grpAt.Controls.Add($txtLog)

function Add-Log {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = $colorDefaultLog
    )
    $time = Get-Date -Format "HH:mm:ss"
    $line = "$time  $Message`r`n"

    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor = $Color
    $txtLog.AppendText($line)
    $txtLog.SelectionColor = $colorDefaultLog
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()

    Trim-RichTextBox -Box $txtLog
}

# UE -> Server Ping Group
[int]$pingTopY = $y2 + 340

$grpPing = New-Object System.Windows.Forms.GroupBox
$grpPing.Text = "Ping : UE -> Server"
$grpPing.Location = New-Object System.Drawing.Point($baseX, $pingTopY)
$grpPing.Size = New-Object System.Drawing.Size(500, 220)
$grpPing.ForeColor = $labelColor
$form.Controls.Add($grpPing)

$checkPingUE1 = New-Object System.Windows.Forms.CheckBox
$checkPingUE1.Text = "UE1 -> Server"
$checkPingUE1.Location = New-Object System.Drawing.Point(15, 25)
$checkPingUE1.Size = New-Object System.Drawing.Size(150, 20)
Set-DarkCheckBoxStyle $checkPingUE1
$grpPing.Controls.Add($checkPingUE1)

$checkPingUE2 = New-Object System.Windows.Forms.CheckBox
$checkPingUE2.Text = "UE2 -> Server"
$checkPingUE2.Location = New-Object System.Drawing.Point(220, 25)
$checkPingUE2.Size = New-Object System.Drawing.Size(150, 20)
Set-DarkCheckBoxStyle $checkPingUE2
$grpPing.Controls.Add($checkPingUE2)

$lblPingLog = New-Object System.Windows.Forms.Label
$lblPingLog.Text = "Ping Log:"
$lblPingLog.AutoSize = $true
$lblPingLog.Location = New-Object System.Drawing.Point(15, 50)
$grpPing.Controls.Add($lblPingLog)

$txtPingLog = New-Object System.Windows.Forms.RichTextBox
$txtPingLog.Location = New-Object System.Drawing.Point(15, 70)
$txtPingLog.Size = New-Object System.Drawing.Size(450, 130)
$txtPingLog.ReadOnly = $true
$txtPingLog.ScrollBars = "Vertical"
$txtPingLog.BackColor = $textBgColor
$txtPingLog.ForeColor = $textFgColor
$grpPing.Controls.Add($txtPingLog)

function Add-PingLog {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = $colorDefaultLog
    )

    $time = Get-Date -Format "HH:mm:ss"

    # 🔸 여기 핵심: Message 끝에 붙어있는 CR/LF 제거
    $msg = ($Message -replace "(\r?\n)+$","")

    $line = "$time  $msg`r`n"

    $txtPingLog.SelectionStart = $txtPingLog.TextLength
    $txtPingLog.SelectionLength = 0
    $txtPingLog.SelectionColor = $Color
    $txtPingLog.AppendText($line)
    $txtPingLog.SelectionColor = $colorDefaultLog
    $txtPingLog.SelectionStart = $txtPingLog.TextLength
    $txtPingLog.ScrollToCaret()

    Trim-RichTextBox -Box $txtPingLog
}

function Trim-RichTextBox {
    param(
        [System.Windows.Forms.RichTextBox]$Box,
        [int]$MaxChars = 200000   # 필요 시 조절
    )
    if ($Box.TextLength -le $MaxChars) { return }

    # 앞부분을 잘라냄 (여유분 80%로 유지)
    $remove = $Box.TextLength - [int]($MaxChars * 0.8)
    if ($remove -gt 0) {
        $Box.Select(0, $remove)
        $Box.SelectedText = ""
    }
}

# ======================
# Config Save / Load
# ======================
function Save-Config {
    param(
        [string]$Path = $global:ConfigFilePath
    )

    $cfg = @{
        ServerIp = $textServerIp.Text

        UE1 = @{
            Enable = $checkUE1.Checked
            DlOnly = $checkDlOnlyUe1.Checked
            UlOnly = $checkUlOnlyUe1.Checked

            DlPort = $textDlPortUe1.Text
            UlPort = $textUlPortUe1.Text
            DlBw   = $textDlBwUe1.Text
            UlBw   = $textUlBwUe1.Text
            Time   = $textTimeUe1.Text
            Proto  = $comboProtoUe1.SelectedItem
        }

        UE2 = @{
            Enable = $checkUE2.Checked
            DlOnly = $checkDlOnlyUe2.Checked
            UlOnly = $checkUlOnlyUe2.Checked

            DlPort = $textDlPortUe2.Text
            UlPort = $textUlPortUe2.Text
            DlBw   = $textDlBwUe2.Text
            UlBw   = $textUlBwUe2.Text
            Time   = $textTimeUe2.Text
            Proto  = $comboProtoUe2.SelectedItem
        }
    }

    try {
        $json = $cfg | ConvertTo-Json -Depth 5
        $json | Set-Content -Path $Path -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Config saved:`n$Path", "Save Config") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to save config: $($_.Exception.Message)", "Save Config Error") | Out-Null
    }
}

function Load-Config {
    param(
        [string]$Path = $global:ConfigFilePath
    )

    if (-not (Test-Path $Path)) {
        [System.Windows.Forms.MessageBox]::Show("Config file not found:`n$Path", "Load Config") | Out-Null
        return
    }

    try {
        $json = Get-Content -Path $Path -Raw -Encoding UTF8
        $cfg  = $json | ConvertFrom-Json

        # Server IP
        if ($cfg.ServerIp) { $textServerIp.Text = $cfg.ServerIp }

        # ---- UE1 ----
        if ($cfg.UE1) {
            $checkUE1.Checked        = [bool]$cfg.UE1.Enable
            $checkDlOnlyUe1.Checked  = [bool]$cfg.UE1.DlOnly
            $checkUlOnlyUe1.Checked  = [bool]$cfg.UE1.UlOnly

            if ($cfg.UE1.DlPort) { $textDlPortUe1.Text = $cfg.UE1.DlPort }
            if ($cfg.UE1.UlPort) { $textUlPortUe1.Text = $cfg.UE1.UlPort }
            if ($cfg.UE1.DlBw)   { $textDlBwUe1.Text   = $cfg.UE1.DlBw }
            if ($cfg.UE1.UlBw)   { $textUlBwUe1.Text   = $cfg.UE1.UlBw }
            if ($cfg.UE1.Time)   { $textTimeUe1.Text   = $cfg.UE1.Time }

            if ($cfg.UE1.Proto) {
                if ($comboProtoUe1.Items.Contains($cfg.UE1.Proto)) {
                    $comboProtoUe1.SelectedItem = $cfg.UE1.Proto
                }
            }
        }

        # ---- UE2 ----
        if ($cfg.UE2) {
            $checkUE2.Checked        = [bool]$cfg.UE2.Enable
            $checkDlOnlyUe2.Checked  = [bool]$cfg.UE2.DlOnly
            $checkUlOnlyUe2.Checked  = [bool]$cfg.UE2.UlOnly

            if ($cfg.UE2.DlPort) { $textDlPortUe2.Text = $cfg.UE2.DlPort }
            if ($cfg.UE2.UlPort) { $textUlPortUe2.Text = $cfg.UE2.UlPort }
            if ($cfg.UE2.DlBw)   { $textDlBwUe2.Text   = $cfg.UE2.DlBw }
            if ($cfg.UE2.UlBw)   { $textUlBwUe2.Text   = $cfg.UE2.UlBw }
            if ($cfg.UE2.Time)   { $textTimeUe2.Text   = $cfg.UE2.Time }

            if ($cfg.UE2.Proto) {
                if ($comboProtoUe2.Items.Contains($cfg.UE2.Proto)) {
                    $comboProtoUe2.SelectedItem = $cfg.UE2.Proto
                }
            }
        }

        [System.Windows.Forms.MessageBox]::Show("Config loaded.`n$Path", "Load Config") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load config: $($_.Exception.Message)", "Load Config Error") | Out-Null
    }
}

# ======================
# Device Mapping (NDIS + Modem) → UE1/UE2
# ======================
function Update-DeviceInfoLabels {
    if ($global:UE1Info.IP -and $global:UE1Info.ComPort) {
        $lblDeviceUE1.Text = "UE1:  $($global:UE1Info.IP) - $($global:UE1Info.ModemName) ($($global:UE1Info.ComPort))"
    } elseif ($global:UE1Info.IP) {
        $lblDeviceUE1.Text = "UE1:  $($global:UE1Info.IP) - (modem not set)"
    } elseif ($global:UE1Info.ComPort) {
        $lblDeviceUE1.Text = "UE1:  (IP not set) - $($global:UE1Info.ModemName) ($($global:UE1Info.ComPort))"
    } else {
        $lblDeviceUE1.Text = "UE1:  (IP/Modem not set)"
    }

    if ($global:UE2Info.IP -and $global:UE2Info.ComPort) {
        $lblDeviceUE2.Text = "UE2:  $($global:UE2Info.IP) - $($global:UE2Info.ModemName) ($($global:UE2Info.ComPort))"
    } elseif ($global:UE2Info.IP) {
        $lblDeviceUE2.Text = "UE2:  $($global:UE2Info.IP) - (modem not set)"
    } elseif ($global:UE2Info.ComPort) {
        $lblDeviceUE2.Text = "UE2:  (IP not set) - $($global:UE2Info.ModemName) ($($global:UE2Info.ComPort))"
    } else {
        $lblDeviceUE2.Text = "UE2:  (IP/Modem not set)"
    }
}

function Detect-AndMapDevices {

    # 기존 UE IP 기억
    $oldUE1IP = $global:UE1Info.IP
    $oldUE2IP = $global:UE2Info.IP

    # 현재 상태 읽기
    $ips    = Get-NdisIpList
    $modems = Get-ModemPorts

    $newUE1IP = $null
    $newUE2IP = $null

    # ===== 1) IP 매핑 =====

    if ($ips.Count -eq 1) {
        # ★ IP 가 1개만 있으면 무조건 UE1 로 본다
        $newUE1IP = $ips[0]
        $newUE2IP = $null
    }
    else {
        # IP 가 0개 또는 2개 이상일 때는
        #   - 이전 IP 가 살아있으면 같은 UE 에 유지
        #   - 남은 IP 는 빈 슬롯에 순서대로 할당

        $ipList = @()
        $ipList += $ips

        if ($oldUE1IP -and ($ipList -contains $oldUE1IP)) {
            $newUE1IP = $oldUE1IP
            $ipList   = $ipList | Where-Object { $_ -ne $oldUE1IP }
        }
        if ($oldUE2IP -and ($ipList -contains $oldUE2IP)) {
            $newUE2IP = $oldUE2IP
            $ipList   = $ipList | Where-Object { $_ -ne $oldUE2IP }
        }

        foreach ($ip in $ipList) {
            if (-not $newUE1IP) { $newUE1IP = $ip; continue }
            if (-not $newUE2IP) { $newUE2IP = $ip; continue }
        }
    }

    $global:UE1Info.IP = $newUE1IP
    $global:UE2Info.IP = $newUE2IP

    # ===== 2) 모뎀(COM 포트) 매핑 (기존 규칙 유지) =====
    $global:UE1Info.ModemName = $null
    $global:UE1Info.ComPort   = $null
    $global:UE2Info.ModemName = $null
    $global:UE2Info.ComPort   = $null

    if ($modems.Count -eq 2) {
        # 2대일 때는 교차 매핑
        $global:UE1Info.ModemName = $modems[1].Name
        $global:UE1Info.ComPort   = $modems[1].AttachedTo

        $global:UE2Info.ModemName = $modems[0].Name
        $global:UE2Info.ComPort   = $modems[0].AttachedTo
    }
    else {
        if ($modems.Count -ge 1) {
            $global:UE1Info.ModemName = $modems[0].Name
            $global:UE1Info.ComPort   = $modems[0].AttachedTo
        }
        if ($modems.Count -ge 2) {
            $global:UE2Info.ModemName = $modems[1].Name
            $global:UE2Info.ComPort   = $modems[1].AttachedTo
        }
    }

    Update-DeviceInfoLabels

    # 디버깅용 로그
    $ipListStr    = if ($ips.Count)    { $ips -join ", " }             else { "(none)" }
    $modemPortStr = if ($modems.Count) { $modems.AttachedTo -join ", " } else { "(none)" }

    Add-Log "Detected NDIS IP(s): $ipListStr  /  Modem port(s): $modemPortStr" $colorDefaultLog
}

function Start-RebootDetect {
    # 이전 타이머 있으면 정리
    if ($global:RebootTimer) {
        try {
            $global:RebootTimer.Stop()
            $global:RebootTimer.Dispose()
        } catch {}
        $global:RebootTimer = $null
    }

    $global:RebootDetectTry = 0
    $global:RebootTimer = New-Object System.Windows.Forms.Timer
    $global:RebootTimer.Interval = 3000  # 3초마다 재시도 (원하면 2000/5000 등으로 조절)

    $global:RebootTimer.Add_Tick({
        $global:RebootDetectTry++

        Detect-AndMapDevices

        # 최소 한 개 UE라도 COM 포트가 잡히면 성공으로 간주
        if ($global:UE1Info.ComPort -or $global:UE2Info.ComPort) {
            Add-Log "Reboot completed. Device info re-detected (try $($global:RebootDetectTry))." $colorDefaultLog
            $global:RebootTimer.Stop()
            $global:RebootTimer.Dispose()
            $global:RebootTimer = $null
        }
        elseif ($global:RebootDetectTry -ge 10) {
            # 총 10번(=30초) 시도 후 포기
            Add-Log "Reboot detect timeout. Please check USB/NDIS and try manual Detect again." $colorDefaultLog
            $global:RebootTimer.Stop()
            $global:RebootTimer.Dispose()
            $global:RebootTimer = $null
        }
        else {
            Add-Log "Waiting UE device to re-enumerate... (try $($global:RebootDetectTry))" $colorDefaultLog
        }
    })

    $global:RebootTimer.Start()
}

# Swap UE1 <-> UE2
$btnSwapUE.Add_Click({
    $tmp = $global:UE1Info
    $global:UE1Info = $global:UE2Info
    $global:UE2Info = $tmp
    Update-DeviceInfoLabels
    Add-Log "Swapped UE1 and UE2 mapping." $colorDefaultLog
})

function Send-AT-ToUE {
    param(
        [string]$Command,
        [string]$UE,                               # "UE1" or "UE2"
        [System.Drawing.Color]$Color
    )

    $info = if ($UE -eq "UE1") { $global:UE1Info } else { $global:UE2Info }

    if (-not $info.ComPort) {
        Add-Log "$($UE): modem COM port not set." $Color
        return
    }

    $cmd = $Command.Trim()
    if ($cmd -eq "") {
        Add-Log "No command to send." $Color
        return
    }

    $send = $cmd + "`r`n"
    $portName = $info.ComPort

    $maxWaitMs = 10000
    $swWait = [System.Diagnostics.Stopwatch]::StartNew()
    $portOpened = $false
    $sp = $null

    while (-not $portOpened -and $swWait.ElapsedMilliseconds -lt $maxWaitMs) {
        try {
            $sp = New-Object System.IO.Ports.SerialPort $portName, 115200, "None", 8, "One"
            $sp.Handshake    = "None"
            $sp.ReadTimeout  = 1000
            $sp.WriteTimeout = 1000
            $sp.DtrEnable    = $true
            $sp.RtsEnable    = $true
            $sp.NewLine      = "`r`n"

            $sp.Open()
            $portOpened = $true
        } catch {
            Start-Sleep -Milliseconds 300
        }
    }

    if (-not $portOpened) {
        Add-Log "[$UE] Error: COM port $portName is not available (timeout / in use)." $Color
        return
    }

    try {
        Add-Log "[$UE] TX $($portName): $cmd" $Color

        $sp.DiscardInBuffer()
        $sp.Write($send)

        $builder = New-Object System.Text.StringBuilder
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        while ($sw.ElapsedMilliseconds -lt 2000) {
            try {
                $chunk = $sp.ReadExisting()
                if ($chunk -and $chunk.Length -gt 0) {
                    [void]$builder.Append($chunk)
                }
            } catch {}
            Start-Sleep -Milliseconds 100
        }

        $resp = $builder.ToString()
        if ($resp) {
            $respLines = $resp -split "`r?`n"
            foreach ($line in $respLines) {
                $clean = $line.Trim()
                if ($clean -ne "") {
                    Add-Log "[$UE] RX $($portName): $clean" $Color
                }
            }
        } else {
            Add-Log "[$UE] RX $($portName): (no data)" $Color
        }
    } catch {
        Add-Log "[$UE] Error on $($portName): $($_.Exception.Message)" $Color
    } finally {
        if ($sp) {
            try { if ($sp.IsOpen) { $sp.Close() } } catch {}
            try { $sp.Dispose() } catch {}
        }
    }
}


$btnSend.Add_Click({
    $cmd = $txtCmd.Text
    if ($checkAtUE1.Checked) { Send-AT-ToUE $cmd "UE1" $colorUE1 }
    if ($checkAtUE2.Checked) { Send-AT-ToUE $cmd "UE2" $colorUE2 }
})

$btnCFUN0.Add_Click({
    if ($checkAtUE1.Checked) { Send-AT-ToUE "AT+CFUN=0" "UE1" $colorUE1 }
    if ($checkAtUE2.Checked) { Send-AT-ToUE "AT+CFUN=0" "UE2" $colorUE2 }
})

$btnCFUN1.Add_Click({
    if ($checkAtUE1.Checked) { Send-AT-ToUE "AT+CFUN=1" "UE1" $colorUE1 }
    if ($checkAtUE2.Checked) { Send-AT-ToUE "AT+CFUN=1" "UE2" $colorUE2 }
})

$btnReboot.Add_Click({
    # 1) 선택된 UE 들에 Reboot AT 전송
    if ($checkAtUE1.Checked) { Send-AT-ToUE "AT+CFUN=1,1" "UE1" $colorUE1 }
    if ($checkAtUE2.Checked) { Send-AT-ToUE "AT+CFUN=1,1" "UE2" $colorUE2 }

    Add-Log "Reboot command sent. Waiting for UE(s) to come back..." $colorDefaultLog

    # 2) Reboot 후 주기적으로 Device Info 재검출
    Start-RebootDetect
})

# ======================
# Ping Events
# ======================
$checkPingUE1.Add_CheckedChanged({
    if ($checkPingUE1.Checked) {
        $server = $textServerIp.Text.Trim()
        $bind   = $global:UE1Info.IP

        if (-not $server -or -not $bind) {
            [System.Windows.Forms.MessageBox]::Show("Server IP or UE1 IP is empty (Device Info).", "Ping UE1 Error") | Out-Null
            $checkPingUE1.Checked = $false
            return
        }

        if ($global:PingJobUE1) {
            try { Stop-Job -Job $global:PingJobUE1 -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $global:PingJobUE1 -Force -ErrorAction SilentlyContinue } catch {}
            $global:PingJobUE1 = $null
        }

        try {
            $global:PingJobUE1 = Start-Job -ScriptBlock {
                param($Server, $Bind)
                ping.exe -t -S $Bind $Server
            } -ArgumentList $server, $bind
            Add-PingLog "[UE1] ping started (Bind=$bind, Server=$server)" $colorUE1
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to start ping UE1: $($_.Exception.Message)", "Ping UE1 Error") | Out-Null
            $checkPingUE1.Checked = $false
        }
    } else {
        if ($global:PingJobUE1) {
            try { Stop-Job -Job $global:PingJobUE1 -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $global:PingJobUE1 -Force -ErrorAction SilentlyContinue } catch {}
            Add-PingLog "[UE1] ping stopped." $colorUE1
            $global:PingJobUE1 = $null
        }
    }
})

$checkPingUE2.Add_CheckedChanged({
    if ($checkPingUE2.Checked) {
        $server = $textServerIp.Text.Trim()
        $bind   = $global:UE2Info.IP

        if (-not $server -or -not $bind) {
            [System.Windows.Forms.MessageBox]::Show("Server IP or UE2 IP is empty (Device Info).", "Ping UE2 Error") | Out-Null
            $checkPingUE2.Checked = $false
            return
        }

        if ($global:PingJobUE2) {
            try { Stop-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
            $global:PingJobUE2 = $null
        }

        try {
            $global:PingJobUE2 = Start-Job -ScriptBlock {
                param($Server, $Bind)
                ping.exe -t -S $Bind $Server
            } -ArgumentList $server, $bind
            Add-PingLog "[UE2] ping started (Bind=$bind, Server=$server)" $colorUE2
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to start ping UE2: $($_.Exception.Message)", "Ping UE2 Error") | Out-Null
            $checkPingUE2.Checked = $false
        }
    } else {
        if ($global:PingJobUE2) {
            try { Stop-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
            Add-PingLog "[UE2] ping stopped." $colorUE2
            $global:PingJobUE2 = $null
        }
    }
})

# Ping Job 출력 수집용 타이머
$pingTimer = New-Object System.Windows.Forms.Timer
$pingTimer.Interval = 1000
$pingTimer.Add_Tick({
    if ($global:PingTickBusy) { return }
    $global:PingTickBusy = $true
    try {
        # --- 여기부터 기존 코드 그대로 ---
        if ($global:PingJobUE1) {
            try {
                $out1 = Receive-Job -Job $global:PingJobUE1 -ErrorAction SilentlyContinue
                if ($out1) {
                    $sb = New-Object System.Text.StringBuilder
                    foreach ($line in $out1) {
                        if ($line -and $line.Trim() -ne "") {
                            $clean = $line.TrimEnd()
                            [void]$sb.Append("[UE1] $clean")
                        }
                    }
                    $text = $sb.ToString()
                    if ($text.Length -gt 0) { Add-PingLog $text $colorUE1 }
                }
            } catch {}
        }

        if ($global:PingJobUE2) {
            try {
                $out2 = Receive-Job -Job $global:PingJobUE2 -ErrorAction SilentlyContinue
                if ($out2) {
                    $sb = New-Object System.Text.StringBuilder
                    foreach ($line in $out2) {
                        if ($line -and $line.Trim() -ne "") {
                            $clean = $line.TrimEnd()
                            [void]$sb.Append("[UE2] $clean")
                        }
                    }
                    $text = $sb.ToString()
                    if ($text.Length -gt 0) { Add-PingLog $text $colorUE2 }
                }
            } catch {}
        }
        # --- 여기까지 기존 코드 그대로 ---
    }
    finally {
        $global:PingTickBusy = $false
    }
})

$pingTimer.Start()

# ======================
# iPerf / Route Events
# ======================
$buttonClose.Add_Click({
    
    try { $pingTimer.Stop(); $pingTimer.Dispose() } catch {}
    
    if ($global:PingJobUE1) {
        try { Stop-Job -Job $global:PingJobUE1 -Force -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job -Job $global:PingJobUE1 -Force -ErrorAction SilentlyContinue } catch {}
        $global:PingJobUE1 = $null
    }
    if ($global:PingJobUE2) {
        try { Stop-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
        $global:PingJobUE2 = $null
    }
    $form.Close()
})

$buttonStart.Add_Click({
    $server = $textServerIp.Text.Trim()

    # Ports
    $dl1p = $textDlPortUe1.Text.Trim()
    $ul1p = $textUlPortUe1.Text.Trim()
    $dl2p = $textDlPortUe2.Text.Trim()
    $ul2p = $textUlPortUe2.Text.Trim()

    # UE1 설정 값
    $dlBw1  = $textDlBwUe1.Text.Trim()
    $ulBw1  = $textUlBwUe1.Text.Trim()
    $t1     = $textTimeUe1.Text.Trim()
    $proto1 = $comboProtoUe1.SelectedItem

    # UE2 설정 값
    $dlBw2  = $textDlBwUe2.Text.Trim()
    $ulBw2  = $textUlBwUe2.Text.Trim()
    $t2     = $textTimeUe2.Text.Trim()
    $proto2 = $comboProtoUe2.SelectedItem

    $useUE1 = $checkUE1.Checked
    $useUE2 = $checkUE2.Checked

    if (-not $useUE1 -and -not $useUE2) {
        [System.Windows.Forms.MessageBox]::Show("Please enable at least one UE (UE1 or UE2).", "No UE Selected") | Out-Null
        return
    }

    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Server IP is empty.", "Input Error") | Out-Null
        return
    }

    $global:psList = @()

    function Start-TestWindow($cmd) {
        $p = Start-Process -FilePath powershell -ArgumentList "-NoExit", "-Command", $cmd -PassThru
        $global:psList += $p.Id
    }

    # ---------- UE1 ----------
    if ($useUE1) {
        $bind1 = $global:UE1Info.IP
        $udp1  = if ($proto1 -eq "UDP") { "-u" } else { "-P 5" }

        # DL only / UL only 상태
        $dlOnly1 = $checkDlOnlyUe1.Checked
        $ulOnly1 = $checkUlOnlyUe1.Checked

        # 기본값: 둘 다 실행
        $runDL1 = $true
        $runUL1 = $true

        if ($dlOnly1 -and -not $ulOnly1) {
            $runUL1 = $false     # DL만
        } elseif ($ulOnly1 -and -not $dlOnly1) {
            $runDL1 = $false     # UL만
        }
        # 둘 다 체크 or 둘 다 미체크 → 둘 다 실행

        if ($bind1 -and $dl1p -and $dlBw1 -and $ul1p -and $ulBw1 -and $t1) {
            if ($runDL1) {
                $DL1 = "iperf3.exe -c $server -p $dl1p $udp1 -b $dlBw1 -t $t1 -R --bind $bind1"
                Start-TestWindow $DL1
            }
            if ($runUL1) {
                $UL1 = "iperf3.exe -c $server -p $ul1p $udp1 -b $ulBw1 -t $t1 --bind $bind1"
                Start-TestWindow $UL1
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("UE1 fields or IP (Device Info) are incomplete.", "Input Error (UE1)") | Out-Null
        }
    }

    # ---------- UE2 ----------
    if ($useUE2) {
        $bind2 = $global:UE2Info.IP
        $udp2  = if ($proto2 -eq "UDP") { "-u" } else { "-P 5" }

        # DL only / UL only 상태
        $dlOnly2 = $checkDlOnlyUe2.Checked
        $ulOnly2 = $checkUlOnlyUe2.Checked

        $runDL2 = $true
        $runUL2 = $true

        if ($dlOnly2 -and -not $ulOnly2) {
            $runUL2 = $false
        } elseif ($ulOnly2 -and -not $dlOnly2) {
            $runDL2 = $false
        }

        if ($bind2 -and $dl2p -and $dlBw2 -and $ul2p -and $ulBw2 -and $t2) {
            if ($runDL2) {
                $DL2 = "iperf3.exe -c $server -p $dl2p $udp2 -b $dlBw2 -t $t2 -R --bind $bind2"
                Start-TestWindow $DL2
            }
            if ($runUL2) {
                $UL2 = "iperf3.exe -c $server -p $ul2p $udp2 -b $ulBw2 -t $t2 --bind $bind2"
                Start-TestWindow $UL2
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("UE2 fields or IP (Device Info) are incomplete.", "Input Error (UE2)") | Out-Null
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Selected UE sessions have been started.", "iperf3") | Out-Null
})   # ← Start 버튼 핸들러 끝!

# Stop 버튼
$buttonStop.Add_Click({
    try {
        $iperf = Get-Process iperf3 -ErrorAction SilentlyContinue
        if ($iperf) {
            $iperf | Stop-Process -Force -ErrorAction SilentlyContinue
        }

        foreach ($psId in $global:psList) {
            try { Stop-Process -Id $psId -Force -ErrorAction SilentlyContinue } catch {}
        }
        $global:psList = @()

        [System.Windows.Forms.MessageBox]::Show("All iperf3 sessions and windows have been closed.", "iperf3") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while stopping sessions: $($_.Exception.Message)", "Error") | Out-Null
    }
})

$buttonSaveCfg.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.InitialDirectory = $scriptDir
    $dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dlg.FileName = "UE_Test_Config.json"

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Save-Config -Path $dlg.FileName
        $global:ConfigFilePath = $dlg.FileName   # 이후 기본 경로를 방금 저장한 파일로 갱신
    }
})

$buttonLoadCfg.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog

    if (Test-Path $global:ConfigFilePath) {
        $dlg.InitialDirectory = Split-Path $global:ConfigFilePath -Parent
        $dlg.FileName = (Split-Path $global:ConfigFilePath -Leaf)
    } else {
        $dlg.InitialDirectory = $scriptDir
        $dlg.FileName = "UE_Test_Config.json"
    }

    $dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Load-Config -Path $dlg.FileName
        $global:ConfigFilePath = $dlg.FileName   # 이후 자동 로드용 기본 파일도 이걸로
    }
})

# Add Routing (자동: UE1/UE2 IP 사용)
$buttonRouteAdmin.Add_Click({
    $server = $textServerIp.Text.Trim()

    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Server IP is empty.", "Route Error") | Out-Null
        return
    }

    $cmdParts = @()

    if ($checkUE1.Checked -and $global:UE1Info.IP) {
        $info1 = Get-RouteInfoForIp -IpAddress $global:UE1Info.IP
        if ($info1.Gateway -and $info1.IfIndex) {
            $cmdParts += "route ADD $server MASK 255.255.255.255 $($info1.Gateway) METRIC 1 IF $($info1.IfIndex)"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Could not find gateway/IF for UE1 IP ($($global:UE1Info.IP)).", "Route Warning (UE1)") | Out-Null
        }
    }

    if ($checkUE2.Checked -and $global:UE2Info.IP) {
        $info2 = Get-RouteInfoForIp -IpAddress $global:UE2Info.IP
        if ($info2.Gateway -and $info2.IfIndex) {
            $cmdParts += "route ADD $server MASK 255.255.255.255 $($info2.Gateway) METRIC 1 IF $($info2.IfIndex)"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Could not find gateway/IF for UE2 IP ($($global:UE2Info.IP)).", "Route Warning (UE2)") | Out-Null
        }
    }

    if ($cmdParts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No route to add. Check UE checkbox and Device Info IP / gateway.", "Route Error") | Out-Null
        return
    }

    $cmdLine = $cmdParts -join " & "
    try {
        Start-Process "cmd.exe" -Verb RunAs -ArgumentList "/k $cmdLine" | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to start admin CMD: $($_.Exception.Message)", "Route Error") | Out-Null
    }
})

# Route Delete
$buttonRouteDelete.Add_Click({
    $server = $textServerIp.Text.Trim()

    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Server IP is empty.", "Route Delete Error") | Out-Null
        return
    }

    $cmdLine = "route DELETE $server"
    try {
        Start-Process "cmd.exe" -Verb RunAs -ArgumentList "/k $cmdLine" | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to delete route: $($_.Exception.Message)", "Route Delete Error") | Out-Null
    }
})

# ======================
# Form Shown: 초기 디바이스 매핑 + 설정 자동 로드
# ======================
$form.Add_Shown({
    Detect-AndMapDevices

    if (Test-Path $global:ConfigFilePath) {
        Load-Config -Path $global:ConfigFilePath
    }
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
