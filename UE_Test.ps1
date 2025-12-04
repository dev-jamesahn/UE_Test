Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ======================
# Global variables
# ======================
$global:psList      = @()   # iperf3 PowerShell window PID list
$global:PingJobUE1  = $null
$global:PingJobUE2  = $null

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
$form.ClientSize = New-Object System.Drawing.Size(950, 750)
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
[int]$widthInput = 240
[int]$height = 23
[int]$gap = 8

[int]$y = 10

$lblDeviceInfoTitle = New-Object System.Windows.Forms.Label
$lblDeviceInfoTitle.Text = "Device Info"
$lblDeviceInfoTitle.Location = New-Object System.Drawing.Point($labelX, $y)
$lblDeviceInfoTitle.AutoSize = $true
Set-DarkLabelStyle $lblDeviceInfoTitle
$form.Controls.Add($lblDeviceInfoTitle)

$y += $height

$lblDeviceUE1 = New-Object System.Windows.Forms.Label
$lblDeviceUE1.Location = New-Object System.Drawing.Point($labelX, $y)
$lblDeviceUE1.Size = New-Object System.Drawing.Size(650, $height)
$lblDeviceUE1.ForeColor = $colorUE1
$lblDeviceUE1.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblDeviceUE1)

$y += $height

$lblDeviceUE2 = New-Object System.Windows.Forms.Label
$lblDeviceUE2.Location = New-Object System.Drawing.Point($labelX, $y)
$lblDeviceUE2.Size = New-Object System.Drawing.Size(650, $height)
$lblDeviceUE2.ForeColor = $colorUE2
$lblDeviceUE2.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblDeviceUE2)

$y += $height + 5

$btnSwapUE = New-Object System.Windows.Forms.Button
$btnSwapUE.Text = "Swap UE1 <-> UE2"
$btnSwapUE.Location = New-Object System.Drawing.Point($labelX, $y)
$btnSwapUE.Size = New-Object System.Drawing.Size(160, 28)
Set-DarkButtonStyle $btnSwapUE
$form.Controls.Add($btnSwapUE)

$y += $height + 20

# ======================
# 좌측: iPerf 설정
# ======================

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

# UE Enable
$checkUE1 = New-Object System.Windows.Forms.CheckBox
$checkUE1.Text = "Enable UE1 (DL1/UL1, Bind #1)"
$checkUE1.Location = New-Object System.Drawing.Point($labelX, $y)
$checkUE1.Size = New-Object System.Drawing.Size(250, $height)
$checkUE1.Checked = $true
Set-DarkCheckBoxStyle $checkUE1
$form.Controls.Add($checkUE1)

$y += $height + $gap

$checkUE2 = New-Object System.Windows.Forms.CheckBox
$checkUE2.Text = "Enable UE2 (DL2/UL2, Bind #2)"
$checkUE2.Location = New-Object System.Drawing.Point($labelX, $y)
$checkUE2.Size = New-Object System.Drawing.Size(250, $height)
$checkUE2.Checked = $true
Set-DarkCheckBoxStyle $checkUE2
$form.Controls.Add($checkUE2)

$y += $height + (2 * $gap)

# DL/UL Ports & BW
$labelDl1Port = New-Object System.Windows.Forms.Label
$labelDl1Port.Text = "DL1 Port (UE1 / Bind #1)"
$labelDl1Port.Location = New-Object System.Drawing.Point($labelX, $y)
$labelDl1Port.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelDl1Port
$form.Controls.Add($labelDl1Port)

$textDl1Port = New-Object System.Windows.Forms.TextBox
$textDl1Port.Location = New-Object System.Drawing.Point($inputX, $y)
$textDl1Port.Size = New-Object System.Drawing.Size($widthInput, $height)
$textDl1Port.Text = "5501"
Set-DarkTextBoxStyle $textDl1Port
$form.Controls.Add($textDl1Port)

$y += $height + $gap

$labelUl1Port = New-Object System.Windows.Forms.Label
$labelUl1Port.Text = "UL1 Port (UE1 / Bind #1)"
$labelUl1Port.Location = New-Object System.Drawing.Point($labelX, $y)
$labelUl1Port.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelUl1Port
$form.Controls.Add($labelUl1Port)

$textUl1Port = New-Object System.Windows.Forms.TextBox
$textUl1Port.Location = New-Object System.Drawing.Point($inputX, $y)
$textUl1Port.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUl1Port.Text = "5502"
Set-DarkTextBoxStyle $textUl1Port
$form.Controls.Add($textUl1Port)

$y += $height + $gap

$labelDl2Port = New-Object System.Windows.Forms.Label
$labelDl2Port.Text = "DL2 Port (UE2 / Bind #2)"
$labelDl2Port.Location = New-Object System.Drawing.Point($labelX, $y)
$labelDl2Port.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelDl2Port
$form.Controls.Add($labelDl2Port)

$textDl2Port = New-Object System.Windows.Forms.TextBox
$textDl2Port.Location = New-Object System.Drawing.Point($inputX, $y)
$textDl2Port.Size = New-Object System.Drawing.Size($widthInput, $height)
$textDl2Port.Text = "5601"
Set-DarkTextBoxStyle $textDl2Port
$form.Controls.Add($textDl2Port)

$y += $height + $gap

$labelUl2Port = New-Object System.Windows.Forms.Label
$labelUl2Port.Text = "UL2 Port (UE2 / Bind #2)"
$labelUl2Port.Location = New-Object System.Drawing.Point($labelX, $y)
$labelUl2Port.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelUl2Port
$form.Controls.Add($labelUl2Port)

$textUl2Port = New-Object System.Windows.Forms.TextBox
$textUl2Port.Location = New-Object System.Drawing.Point($inputX, $y)
$textUl2Port.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUl2Port.Text = "5602"
Set-DarkTextBoxStyle $textUl2Port
$form.Controls.Add($textUl2Port)

$y += $height + (2 * $gap)

$labelDlBw = New-Object System.Windows.Forms.Label
$labelDlBw.Text = "Download BW"
$labelDlBw.Location = New-Object System.Drawing.Point($labelX, $y)
$labelDlBw.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelDlBw
$form.Controls.Add($labelDlBw)

$textDlBw = New-Object System.Windows.Forms.TextBox
$textDlBw.Location = New-Object System.Drawing.Point($inputX, $y)
$textDlBw.Size = New-Object System.Drawing.Size($widthInput, $height)
$textDlBw.Text = "1000M"
Set-DarkTextBoxStyle $textDlBw
$form.Controls.Add($textDlBw)

$y += $height + $gap

$labelUlBw = New-Object System.Windows.Forms.Label
$labelUlBw.Text = "Upload BW"
$labelUlBw.Location = New-Object System.Drawing.Point($labelX, $y)
$labelUlBw.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelUlBw
$form.Controls.Add($labelUlBw)

$textUlBw = New-Object System.Windows.Forms.TextBox
$textUlBw.Location = New-Object System.Drawing.Point($inputX, $y)
$textUlBw.Size = New-Object System.Drawing.Size($widthInput, $height)
$textUlBw.Text = "500M"
Set-DarkTextBoxStyle $textUlBw
$form.Controls.Add($textUlBw)

$y += $height + $gap

# Bind IP 텍스트박스는 제거됨!

# Duration
$labelTime = New-Object System.Windows.Forms.Label
$labelTime.Text = "Duration (sec)"
$labelTime.Location = New-Object System.Drawing.Point($labelX, $y)
$labelTime.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelTime
$form.Controls.Add($labelTime)

$textTime = New-Object System.Windows.Forms.TextBox
$textTime.Location = New-Object System.Drawing.Point($inputX, $y)
$textTime.Size = New-Object System.Drawing.Size($widthInput, $height)
$textTime.Text = "86400"
Set-DarkTextBoxStyle $textTime
$form.Controls.Add($textTime)

$y += $height + $gap

# Protocol
$labelProto = New-Object System.Windows.Forms.Label
$labelProto.Text = "Protocol"
$labelProto.Location = New-Object System.Drawing.Point($labelX, $y)
$labelProto.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelProto
$form.Controls.Add($labelProto)

$comboProto = New-Object System.Windows.Forms.ComboBox
$comboProto.Location = New-Object System.Drawing.Point($inputX, $y)
$comboProto.Size = New-Object System.Drawing.Size($widthInput, $height)
[void]$comboProto.Items.Add("UDP")
[void]$comboProto.Items.Add("TCP")
$comboProto.SelectedIndex = 0
Set-DarkComboBoxStyle $comboProto
$form.Controls.Add($comboProto)

$y += $height + (2 * $gap)

# iperf Buttons
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Start Selected UEs"
$buttonStart.Location = New-Object System.Drawing.Point(40, $y)
$buttonStart.Size = New-Object System.Drawing.Size(160, 40)
Set-DarkButtonStyle $buttonStart
$form.Controls.Add($buttonStart)

$buttonStop = New-Object System.Windows.Forms.Button
$buttonStop.Text = "Stop All Sessions"
$buttonStop.Location = New-Object System.Drawing.Point(250, $y)
$buttonStop.Size = New-Object System.Drawing.Size(160, 40)
Set-DarkButtonStyle $buttonStop
$form.Controls.Add($buttonStop)

$y += $height + 2*$gap + 30

$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Exit"
$buttonClose.Location = New-Object System.Drawing.Point(180, $y)
$buttonClose.Size = New-Object System.Drawing.Size(120, 40)
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
[int]$y2 = 160   # 약간 아래에서 시작

# AT Command Group
$grpAt = New-Object System.Windows.Forms.GroupBox
$grpAt.Text = "AT Command"
$grpAt.Location = New-Object System.Drawing.Point($baseX, $y2)
$grpAt.Size = New-Object System.Drawing.Size(430, 320)
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

$btnCFUN0 = New-Object System.Windows.Forms.Button
$btnCFUN0.Text = "CFUN=0"
$btnCFUN0.Location = New-Object System.Drawing.Point(325, 73)
$btnCFUN0.Size = New-Object System.Drawing.Size(80, 26)
Set-DarkButtonStyle $btnCFUN0
$grpAt.Controls.Add($btnCFUN0)

$btnCFUN1 = New-Object System.Windows.Forms.Button
$btnCFUN1.Text = "CFUN=1"
$btnCFUN1.Location = New-Object System.Drawing.Point(325, 105)
$btnCFUN1.Size = New-Object System.Drawing.Size(80, 26)
Set-DarkButtonStyle $btnCFUN1
$grpAt.Controls.Add($btnCFUN1)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.AutoSize = $true
$lblLog.Location = New-Object System.Drawing.Point(15, 110)
$grpAt.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 130)
$txtLog.Size = New-Object System.Drawing.Size(390, 160)
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
}

# UE -> Server Ping Group
[int]$pingTopY = $y2 + 340

$grpPing = New-Object System.Windows.Forms.GroupBox
$grpPing.Text = "UE -> Server Ping"
$grpPing.Location = New-Object System.Drawing.Point($baseX, $pingTopY)
$grpPing.Size = New-Object System.Drawing.Size(430, 220)
$grpPing.ForeColor = $labelColor
$form.Controls.Add($grpPing)

$checkPingUE1 = New-Object System.Windows.Forms.CheckBox
$checkPingUE1.Text = "Ping UE1 -> Server"
$checkPingUE1.Location = New-Object System.Drawing.Point(15, 25)
$checkPingUE1.Size = New-Object System.Drawing.Size(150, 20)
Set-DarkCheckBoxStyle $checkPingUE1
$grpPing.Controls.Add($checkPingUE1)

$checkPingUE2 = New-Object System.Windows.Forms.CheckBox
$checkPingUE2.Text = "Ping UE2 -> Server"
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
$txtPingLog.Size = New-Object System.Drawing.Size(390, 130)
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
    $line = "$time  $Message`r`n"

    $txtPingLog.SelectionStart = $txtPingLog.TextLength
    $txtPingLog.SelectionLength = 0
    $txtPingLog.SelectionColor = $Color
    $txtPingLog.AppendText($line)
    $txtPingLog.SelectionColor = $colorDefaultLog
    $txtPingLog.SelectionStart = $txtPingLog.TextLength
    $txtPingLog.ScrollToCaret()
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
    $ips    = Get-NdisIpList
    $modems = Get-ModemPorts

    # 초기화
    $global:UE1Info.IP        = $null
    $global:UE1Info.ModemName = $null
    $global:UE1Info.ComPort   = $null
    $global:UE2Info.IP        = $null
    $global:UE2Info.ModemName = $null
    $global:UE2Info.ComPort   = $null

    # 기본적으로 IP 순서는 그대로 사용
    if ($ips.Count -ge 1) { $global:UE1Info.IP = $ips[0] }
    if ($ips.Count -ge 2) { $global:UE2Info.IP = $ips[1] }

    # --- 여기에서 모뎀 순서를 교환해서 매핑 ---
    if ($ips.Count -eq 2 -and $modems.Count -eq 2) {
        # UE1 IP -> 두 번째 모뎀, UE2 IP -> 첫 번째 모뎀
        $global:UE1Info.ModemName = $modems[1].Name
        $global:UE1Info.ComPort   = $modems[1].AttachedTo

        $global:UE2Info.ModemName = $modems[0].Name
        $global:UE2Info.ComPort   = $modems[0].AttachedTo
    }
    else {
        # 2대가 아닌 경우는 기존 방식 유지
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
    Add-Log "Detected $($modems.Count) modem port(s)." $colorDefaultLog
}

# Swap UE1 <-> UE2
$btnSwapUE.Add_Click({
    $tmp = $global:UE1Info
    $global:UE1Info = $global:UE2Info
    $global:UE2Info = $tmp
    Update-DeviceInfoLabels
    Add-Log "Swapped UE1 and UE2 mapping." $colorDefaultLog
})

# ======================
# AT Command Send
# ======================
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

    try {
        $sp = New-Object System.IO.Ports.SerialPort $portName, 115200, "None", 8, "One"
        $sp.Handshake   = "None"
        $sp.ReadTimeout = 1000
        $sp.WriteTimeout= 1000
        $sp.DtrEnable   = $true
        $sp.RtsEnable   = $true
        $sp.NewLine     = "`r`n"

        $sp.Open()

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
            $respLines = $resp -replace "`r`n","`n" -split "`n"
            foreach ($line in $respLines) {
                if ($line.Trim() -ne "") {
                    Add-Log "[$UE] RX $($portName): $line" $Color
                }
            }
        } else {
            Add-Log "[$UE] RX $($portName): (no data)" $Color
        }

        $sp.Close()
    } catch {
        Add-Log "[$UE] Error on $($portName): $($_.Exception.Message)" $Color
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
    if ($global:PingJobUE1) {
        try {
            $out1 = Receive-Job -Job $global:PingJobUE1 -ErrorAction SilentlyContinue
            if ($out1) {
                foreach ($line in $out1) {
                    if ($line -and $line.Trim() -ne "") {
                        Add-PingLog "[UE1] $line" $colorUE1
                    }
                }
            }
        } catch {}
    }
    if ($global:PingJobUE2) {
        try {
            $out2 = Receive-Job -Job $global:PingJobUE2 -ErrorAction SilentlyContinue
            if ($out2) {
                foreach ($line in $out2) {
                    if ($line -and $line.Trim() -ne "") {
                        Add-PingLog "[UE2] $line" $colorUE2
                    }
                }
            }
        } catch {}
    }
})
$pingTimer.Start()

# ======================
# iPerf / Route Events
# ======================
$buttonClose.Add_Click({
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
    $dl1p = $textDl1Port.Text.Trim()
    $ul1p = $textUl1Port.Text.Trim()
    $dl2p = $textDl2Port.Text.Trim()
    $ul2p = $textUl2Port.Text.Trim()
    $dlBw = $textDlBw.Text.Trim()
    $ulBw = $textUlBw.Text.Trim()
    $t    = $textTime.Text.Trim()
    $proto= $comboProto.SelectedItem
    $useUE1 = $checkUE1.Checked
    $useUE2 = $checkUE2.Checked

    if (-not $useUE1 -and -not $useUE2) {
        [System.Windows.Forms.MessageBox]::Show("Please enable at least one UE (UE1 or UE2).", "No UE Selected") | Out-Null
        return
    }

    if (-not $server -or -not $dlBw -or -not $ulBw -or -not $t) {
        [System.Windows.Forms.MessageBox]::Show("Some required fields are empty.", "Input Error") | Out-Null
        return
    }

    $udp = ""
    if ($proto -eq "UDP") { $udp = "-u" }

    $global:psList = @()

    function Start-TestWindow($cmd) {
        $p = Start-Process -FilePath powershell -ArgumentList "-NoExit", "-Command", $cmd -PassThru
        $global:psList += $p.Id
    }

    if ($useUE1) {
        $bind1 = $global:UE1Info.IP
        if (-not $dl1p -or -not $ul1p -or -not $bind1) {
            [System.Windows.Forms.MessageBox]::Show("UE1 fields or IP (Device Info) are incomplete.", "Input Error (UE1)") | Out-Null
        } else {
            $DL1 = "iperf3.exe -c $server -p $dl1p $udp -b $dlBw -t $t -R --bind $bind1"
            $UL1 = "iperf3.exe -c $server -p $ul1p $udp -b $ulBw -t $t --bind $bind1"
            Start-TestWindow $DL1
            Start-TestWindow $UL1
        }
    }

    if ($useUE2) {
        $bind2 = $global:UE2Info.IP
        if (-not $dl2p -or -not $ul2p -or -not $bind2) {
            [System.Windows.Forms.MessageBox]::Show("UE2 fields or IP (Device Info) are incomplete.", "Input Error (UE2)") | Out-Null
        } else {
            $DL2 = "iperf3.exe -c $server -p $dl2p $udp -b $dlBw -t $t -R --bind $bind2"
            $UL2 = "iperf3.exe -c $server -p $ul2p $udp -b $ulBw -t $t --bind $bind2"
            Start-TestWindow $DL2
            Start-TestWindow $UL2
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Selected UE sessions have been started.", "iperf3") | Out-Null
})

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
# Form Shown: 초기 디바이스 매핑
# ======================
$form.Add_Shown({
    Detect-AndMapDevices
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
