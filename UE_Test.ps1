Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================
# Global variables
# ==========================
$global:psList        = @()     # iperf3 PowerShell 윈도우 PID 리스트
$global:ModemList     = @()     # AT 모뎀 리스트
$global:PingJobUE1    = $null   # Ping Job 핸들
$global:PingJobUE2    = $null
$global:Ue1ModemIndex = $null   # 모뎀 인덱스 (COM)
$global:Ue2ModemIndex = $null

# ===== 색상 정의 =====
$colorDefaultLog = [System.Drawing.Color]::White
$colorUE1        = [System.Drawing.Color]::LightGreen
$colorUE2        = [System.Drawing.Color]::LightSkyBlue

# ==========================
# Main Form
# ==========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "my5G UE Test Tool"
$form.ClientSize = New-Object System.Drawing.Size(1000, 850)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$form.ForeColor = [System.Drawing.Color]::White
$form.AutoScroll = $true

# ----- 공통 스타일 함수 -----
$labelColor    = [System.Drawing.Color]::White
$textBgColor   = [System.Drawing.Color]::FromArgb(45,45,48)
$textFgColor   = [System.Drawing.Color]::White
$buttonBgColor = [System.Drawing.Color]::FromArgb(64,64,64)
$buttonFgColor = [System.Drawing.Color]::White

function Set-DarkTextBoxStyle($tb) {
    $tb.BackColor  = $textBgColor
    $tb.ForeColor  = $textFgColor
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

# =====================================================================
# 0) 상단 Device Info 영역 (UE1/UE2 IP + Modem)
# =====================================================================
$lblDeviceInfo = New-Object System.Windows.Forms.Label
$lblDeviceInfo.Text = "Device Info"
$lblDeviceInfo.Location = New-Object System.Drawing.Point(20, 15)
$lblDeviceInfo.Size = New-Object System.Drawing.Size(200, 20)
Set-DarkLabelStyle $lblDeviceInfo
$form.Controls.Add($lblDeviceInfo)

$lblUE1Modem = New-Object System.Windows.Forms.Label
$lblUE1Modem.AutoSize = $false
$lblUE1Modem.Location = New-Object System.Drawing.Point(20, 35)
$lblUE1Modem.Size = New-Object System.Drawing.Size(900, 20)
$lblUE1Modem.ForeColor = $colorUE1
$lblUE1Modem.Text = "UE1: (not set)"
$form.Controls.Add($lblUE1Modem)

$lblUE2Modem = New-Object System.Windows.Forms.Label
$lblUE2Modem.AutoSize = $false
$lblUE2Modem.Location = New-Object System.Drawing.Point(20, 55)
$lblUE2Modem.Size = New-Object System.Drawing.Size(900, 20)
$lblUE2Modem.ForeColor = $colorUE2
$lblUE2Modem.Text = "UE2: (not set)"
$form.Controls.Add($lblUE2Modem)

# UE1 <-> UE2 Swap 버튼 (리스트 아래)
$btnSwapUE = New-Object System.Windows.Forms.Button
$btnSwapUE.Text = "Swap UE1 <-> UE2"
$btnSwapUE.Location = New-Object System.Drawing.Point(20, 80)
$btnSwapUE.Size = New-Object System.Drawing.Size(150, 25)
Set-DarkButtonStyle $btnSwapUE
$form.Controls.Add($btnSwapUE)

# (숨겨둔) 수동 모뎀 매핑 버튼 - 필요시 보이게 할 수 있음
$btnSetUE1 = New-Object System.Windows.Forms.Button
$btnSetUE1.Text = "Set as UE1"
$btnSetUE1.Location = New-Object System.Drawing.Point(800, 70)
$btnSetUE1.Size = New-Object System.Drawing.Size(100, 25)
Set-DarkButtonStyle $btnSetUE1
$btnSetUE1.Visible = $false
$form.Controls.Add($btnSetUE1)

$btnSetUE2 = New-Object System.Windows.Forms.Button
$btnSetUE2.Text = "Set as UE2"
$btnSetUE2.Location = New-Object System.Drawing.Point(800, 100)
$btnSetUE2.Size = New-Object System.Drawing.Size(100, 25)
Set-DarkButtonStyle $btnSetUE2
$btnSetUE2.Visible = $false
$form.Controls.Add($btnSetUE2)

# =====================================================================
# 1) 왼쪽 영역: iperf3 2-UE Bidirectional Throughput
# =====================================================================
[int]$labelX = 20
[int]$inputX = 190
[int]$widthLabel = 160
[int]$widthInput = 240
[int]$height = 23
[int]$gap = 8
[int]$y = 150      # 상단 Device Info 를 위해 여유를 두고 시작

# ----- Server IP -----
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

$y += $height + (3 * $gap)

# ----- Route Buttons -----
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

$y += $height + (3 * $gap)

# ----- UE Enable Checkboxes -----
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

# ----- Port / BW / IP / Time / Protocol -----
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

$labelBindIp1 = New-Object System.Windows.Forms.Label
$labelBindIp1.Text = "Bind IP #1 (UE1)"
$labelBindIp1.Location = New-Object System.Drawing.Point($labelX, $y)
$labelBindIp1.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelBindIp1
$form.Controls.Add($labelBindIp1)

$textBindIp1 = New-Object System.Windows.Forms.TextBox
$textBindIp1.Location = New-Object System.Drawing.Point($inputX, $y)
$textBindIp1.Size = New-Object System.Drawing.Size($widthInput, $height)
$textBindIp1.Text = "192.168.225.82"
Set-DarkTextBoxStyle $textBindIp1
$form.Controls.Add($textBindIp1)

$y += $height + $gap

$labelBindIp2 = New-Object System.Windows.Forms.Label
$labelBindIp2.Text = "Bind IP #2 (UE2)"
$labelBindIp2.Location = New-Object System.Drawing.Point($labelX, $y)
$labelBindIp2.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelBindIp2
$form.Controls.Add($labelBindIp2)

$textBindIp2 = New-Object System.Windows.Forms.TextBox
$textBindIp2.Location = New-Object System.Drawing.Point($inputX, $y)
$textBindIp2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textBindIp2.Text = "192.168.225.30"
Set-DarkTextBoxStyle $textBindIp2
$form.Controls.Add($textBindIp2)

$y += $height + $gap

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

$y += $height + (2 * $gap) + 30

$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Exit"
$buttonClose.Location = New-Object System.Drawing.Point(180, $y)
$buttonClose.Size = New-Object System.Drawing.Size(120, 40)
Set-DarkButtonStyle $buttonClose
$form.Controls.Add($buttonClose)

# =====================================================================
# 2) 오른쪽 영역: AT Command + Ping
# =====================================================================
[int]$baseX = 520
[int]$y2   = 140

$grpAt = New-Object System.Windows.Forms.GroupBox
$grpAt.Text = "AT Command"
$grpAt.Location = New-Object System.Drawing.Point($baseX, $y2)
$grpAt.Size = New-Object System.Drawing.Size(430, 310)
$grpAt.ForeColor = $labelColor
$form.Controls.Add($grpAt)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target UEs:"
$lblTarget.AutoSize = $true
$lblTarget.Location = New-Object System.Drawing.Point(15, 25)
$grpAt.Controls.Add($lblTarget)

$chkTargetUE1 = New-Object System.Windows.Forms.CheckBox
$chkTargetUE1.Text = "UE1"
$chkTargetUE1.Location = New-Object System.Drawing.Point(90, 23)
$chkTargetUE1.Size = New-Object System.Drawing.Size(60, 20)
$chkTargetUE1.Checked = $true
Set-DarkCheckBoxStyle $chkTargetUE1
$grpAt.Controls.Add($chkTargetUE1)

$chkTargetUE2 = New-Object System.Windows.Forms.CheckBox
$chkTargetUE2.Text = "UE2"
$chkTargetUE2.Location = New-Object System.Drawing.Point(150, 23)
$chkTargetUE2.Size = New-Object System.Drawing.Size(60, 20)
$chkTargetUE2.Checked = $true
Set-DarkCheckBoxStyle $chkTargetUE2
$grpAt.Controls.Add($chkTargetUE2)

$lblCmd = New-Object System.Windows.Forms.Label
$lblCmd.Text = "Command"
$lblCmd.AutoSize = $true
$lblCmd.Location = New-Object System.Drawing.Point(15, 55)
$grpAt.Controls.Add($lblCmd)

$txtCmd = New-Object System.Windows.Forms.TextBox
$txtCmd.Location = New-Object System.Drawing.Point(15, 75)
$txtCmd.Size = New-Object System.Drawing.Size(220, 23)
$txtCmd.BackColor = $textBgColor
$txtCmd.ForeColor = $textFgColor
$grpAt.Controls.Add($txtCmd)

$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text = "Send"
$btnSend.Location = New-Object System.Drawing.Point(245, 73)
$btnSend.Size = New-Object System.Drawing.Size(60, 26)
Set-DarkButtonStyle $btnSend
$grpAt.Controls.Add($btnSend)

$btnCFUN0 = New-Object System.Windows.Forms.Button
$btnCFUN0.Text = "CFUN=0"
$btnCFUN0.Location = New-Object System.Drawing.Point(315, 73)
$btnCFUN0.Size = New-Object System.Drawing.Size(80, 26)
Set-DarkButtonStyle $btnCFUN0
$grpAt.Controls.Add($btnCFUN0)

$btnCFUN1 = New-Object System.Windows.Forms.Button
$btnCFUN1.Text = "CFUN=1"
$btnCFUN1.Location = New-Object System.Drawing.Point(315, 105)
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
$txtLog.Size = New-Object System.Drawing.Size(400, 160)
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.BackColor = $textBgColor
$txtLog.ForeColor = $textFgColor
$grpAt.Controls.Add($txtLog)

# AT Log 출력 함수
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

# ----- Ping 영역 -----
[int]$pingTopY = $y2 + 320

$grpPing = New-Object System.Windows.Forms.GroupBox
$grpPing.Text = "UE -> Server Ping"
$grpPing.Location = New-Object System.Drawing.Point($baseX, $pingTopY)
$grpPing.Size = New-Object System.Drawing.Size(430, 210)
$grpPing.ForeColor = $labelColor
$form.Controls.Add($grpPing)

$checkPingUE1 = New-Object System.Windows.Forms.CheckBox
$checkPingUE1.Text = "Ping UE1 -> Server"
$checkPingUE1.Location = New-Object System.Drawing.Point(15, 25)
$checkPingUE1.Size = New-Object System.Drawing.Size(160, 20)
Set-DarkCheckBoxStyle $checkPingUE1
$grpPing.Controls.Add($checkPingUE1)

$checkPingUE2 = New-Object System.Windows.Forms.CheckBox
$checkPingUE2.Text = "Ping UE2 -> Server"
$checkPingUE2.Location = New-Object System.Drawing.Point(200, 25)
$checkPingUE2.Size = New-Object System.Drawing.Size(160, 20)
Set-DarkCheckBoxStyle $checkPingUE2
$grpPing.Controls.Add($checkPingUE2)

$lblPingLog = New-Object System.Windows.Forms.Label
$lblPingLog.Text = "Ping Log:"
$lblPingLog.AutoSize = $true
$lblPingLog.Location = New-Object System.Drawing.Point(15, 50)
$grpPing.Controls.Add($lblPingLog)

$txtPingLog = New-Object System.Windows.Forms.RichTextBox
$txtPingLog.Location = New-Object System.Drawing.Point(15, 70)
$txtPingLog.Size = New-Object System.Drawing.Size(400, 120)
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

# =====================================================================
# 공통 유틸 함수: Remote NDIS / Route / Modem / Summary
# =====================================================================

# NDIS IP 리스트 검색
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

            if ($ad.Description -notlike "*Remote NDIS Compatible Device*") {
                continue
            }

            $ipv4s = $cfg.IPAddress | Where-Object {
                $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notlike "169.254.*"
            }

            foreach ($ip in $ipv4s) {
                if ($ips -notcontains $ip) {
                    $ips += $ip
                }
            }
        }
    }
    catch { }

    return $ips
}

function Refresh-NdisBindIps {
    $ips = Get-NdisIpList

    if ($ips.Count -ge 1) { $textBindIp1.Text = $ips[0] }
    if ($ips.Count -ge 2) { $textBindIp2.Text = $ips[1] }

    Update-UeSummaryLabels
}

# 특 IP가 붙은 NIC 의 Gateway/IF Index
function Get-RouteInfoForIp {
    param(
        [string]$IpAddress
    )

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
    }
    catch { }

    return $result
}

# AT용 모뎀 포트 검색 (Win32_POTSModem)
function Detect-ModemPorts {
    $global:ModemList = @()
    $global:Ue1ModemIndex = $null
    $global:Ue2ModemIndex = $null

    try {
        $modems = Get-CimInstance Win32_POTSModem -ErrorAction Stop
    }
    catch {
        $modems = Get-WmiObject Win32_POTSModem -ErrorAction SilentlyContinue
    }

    if ($modems) {
        $modems = $modems | Where-Object { $_.AttachedTo -match "^COM\d+" }
    }

    if (-not $modems -or $modems.Count -eq 0) {
        Add-Log "No modem ports found (Win32_POTSModem empty)"
        Update-UeSummaryLabels
        return
    }

    foreach ($m in $modems) {
        $global:ModemList += $m
    }

    # 기본 매핑: 0 -> UE1, 1 -> UE2
    if ($global:ModemList.Count -ge 1) { $global:Ue1ModemIndex = 0 }
    if ($global:ModemList.Count -ge 2) { $global:Ue2ModemIndex = 1 }

    Add-Log "Detected $($global:ModemList.Count) modem port(s)."

    Update-UeSummaryLabels
}

# 상단 Device Info 라벨 업데이트
function Update-UeSummaryLabels {
    # UE1
    $ip1  = $textBindIp1.Text.Trim()
    $mod1 = "(not set)"

    if ($global:Ue1ModemIndex -ne $null -and
        $global:Ue1ModemIndex -ge 0 -and
        $global:Ue1ModemIndex -lt $global:ModemList.Count) {
        $m1 = $global:ModemList[$global:Ue1ModemIndex]
        $mod1 = "$($m1.Name) ($($m1.AttachedTo))"
    }

    if ($ip1 -and $mod1 -ne "(not set)") {
        $lblUE1Modem.Text = "UE1:  $ip1 - $mod1"
    }
    elseif ($ip1) {
        $lblUE1Modem.Text = "UE1:  $ip1 - (modem not set)"
    }
    elseif ($mod1 -ne "(not set)") {
        $lblUE1Modem.Text = "UE1:  (no IP) - $mod1"
    }
    else {
        $lblUE1Modem.Text = "UE1:  (not set)"
    }

    # UE2
    $ip2  = $textBindIp2.Text.Trim()
    $mod2 = "(not set)"

    if ($global:Ue2ModemIndex -ne $null -and
        $global:Ue2ModemIndex -ge 0 -and
        $global:Ue2ModemIndex -lt $global:ModemList.Count) {
        $m2 = $global:ModemList[$global:Ue2ModemIndex]
        $mod2 = "$($m2.Name) ($($m2.AttachedTo))"
    }

    if ($ip2 -and $mod2 -ne "(not set)") {
        $lblUE2Modem.Text = "UE2:  $ip2 - $mod2"
    }
    elseif ($ip2) {
        $lblUE2Modem.Text = "UE2:  $ip2 - (modem not set)"
    }
    elseif ($mod2 -ne "(not set)") {
        $lblUE2Modem.Text = "UE2:  (no IP) - $mod2"
    }
    else {
        $lblUE2Modem.Text = "UE2:  (not set)"
    }
}

# =====================================================================
# 이벤트 핸들러
# =====================================================================

# Exit 버튼
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

# iperf Start
$buttonStart.Add_Click({
    $server = $textServerIp.Text.Trim()
    $dl1p   = $textDl1Port.Text.Trim()
    $ul1p   = $textUl1Port.Text.Trim()
    $dl2p   = $textDl2Port.Text.Trim()
    $ul2p   = $textUl2Port.Text.Trim()
    $dlBw   = $textDlBw.Text.Trim()
    $ulBw   = $textUlBw.Text.Trim()
    $bind1  = $textBindIp1.Text.Trim()
    $bind2  = $textBindIp2.Text.Trim()
    $t      = $textTime.Text.Trim()
    $proto  = $comboProto.SelectedItem
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
        if (-not $dl1p -or -not $ul1p -or -not $bind1) {
            [System.Windows.Forms.MessageBox]::Show("UE1 fields are incomplete.", "Input Error (UE1)") | Out-Null
        }
        else {
            $DL1 = "iperf3.exe -c $server -p $dl1p $udp -b $dlBw -t $t -R --bind $bind1"
            $UL1 = "iperf3.exe -c $server -p $ul1p $udp -b $ulBw -t $t --bind $bind1"
            Start-TestWindow $DL1
            Start-TestWindow $UL1
        }
    }

    if ($useUE2) {
        if (-not $dl2p -or -not $ul2p -or -not $bind2) {
            [System.Windows.Forms.MessageBox]::Show("UE2 fields are incomplete.", "Input Error (UE2)") | Out-Null
        }
        else {
            $DL2 = "iperf3.exe -c $server -p $dl2p $udp -b $dlBw -t $t -R --bind $bind2"
            $UL2 = "iperf3.exe -c $server -p $ul2p $udp -b $ulBw -t $t --bind $bind2"
            Start-TestWindow $DL2
            Start-TestWindow $UL2
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Selected UE sessions have been started.", "iperf3") | Out-Null
})

# iperf Stop
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
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error while stopping sessions: $($_.Exception.Message)", "Error") | Out-Null
    }
})

# Routing Add
$buttonRouteAdmin.Add_Click({
    $server = $textServerIp.Text.Trim()
    $bind1  = $textBindIp1.Text.Trim()
    $bind2  = $textBindIp2.Text.Trim()

    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Server IP is empty.", "Route Error") | Out-Null
        return
    }

    $cmdParts = @()

    if ($checkUE1.Checked -and $bind1) {
        $info1 = Get-RouteInfoForIp -IpAddress $bind1
        if ($info1.Gateway -and $info1.IfIndex) {
            $cmdParts += "route ADD $server MASK 255.255.255.255 $($info1.Gateway) METRIC 1 IF $($info1.IfIndex)"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Could not find gateway/IF for UE1 Bind IP ($bind1).", "Route Warning (UE1)") | Out-Null
        }
    }

    if ($checkUE2.Checked -and $bind2) {
        $info2 = Get-RouteInfoForIp -IpAddress $bind2
        if ($info2.Gateway -and $info2.IfIndex) {
            $cmdParts += "route ADD $server MASK 255.255.255.255 $($info2.Gateway) METRIC 1 IF $($info2.IfIndex)"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Could not find gateway/IF for UE2 Bind IP ($bind2).", "Route Warning (UE2)") | Out-Null
        }
    }

    if ($cmdParts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No route to add. Check UE checkbox, Bind IP and gateway.", "Route Error") | Out-Null
        return
    }

    $cmdLine = $cmdParts -join " & "
    try {
        Start-Process "cmd.exe" -Verb RunAs -ArgumentList "/k $cmdLine" | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to start admin CMD: $($_.Exception.Message)", "Route Error") | Out-Null
    }
})

# Routing Delete
$buttonRouteDelete.Add_Click({
    $server = $textServerIp.Text.Trim()
    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Server IP is empty.", "Route Delete Error") | Out-Null
        return
    }

    $cmdLine = "route DELETE $server"
    try {
        Start-Process "cmd.exe" -Verb RunAs -ArgumentList "/k $cmdLine" | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to delete route: $($_.Exception.Message)", "Route Delete Error") | Out-Null
    }
})

# AT Command 전송
function Send-ATToTargets {
    param([string]$Command)

    $cmd = $Command.Trim()
    if (-not $cmd) {
        Add-Log "No command to send"
        return
    }

    # 대상 UE -> 모뎀 인덱스 리스트 구성
    $indices = @()

    if ($chkTargetUE1.Checked -and $global:Ue1ModemIndex -ne $null) {
        $indices += $global:Ue1ModemIndex
    }
    if ($chkTargetUE2.Checked -and $global:Ue2ModemIndex -ne $null) {
        $indices += $global:Ue2ModemIndex
    }

    if ($indices.Count -eq 0) {
        Add-Log "No target UE selected / modem not mapped"
        return
    }

    $send = $cmd + "`r`n"

    foreach ($idx in $indices) {
        if ($idx -lt 0 -or $idx -ge $global:ModemList.Count) { continue }

        $modem = $global:ModemList[$idx]
        $portName = $modem.AttachedTo

        # 인덱스에 따라 색 결정
        $logColor = $colorDefaultLog
        if ($idx -eq $global:Ue1ModemIndex) { $logColor = $colorUE1 }
        elseif ($idx -eq $global:Ue2ModemIndex) { $logColor = $colorUE2 }

        try {
            $sp = New-Object System.IO.Ports.SerialPort $portName, 115200, "None", 8, "One"
            $sp.Handshake = "None"
            $sp.ReadTimeout = 1000
            $sp.WriteTimeout = 1000
            $sp.DtrEnable = $true
            $sp.RtsEnable = $true
            $sp.NewLine = "`r`n"

            $sp.Open()

            Add-Log "[$portName] TX: $cmd" $logColor

            $sp.DiscardInBuffer()
            $sp.Write($send)

            $builder = New-Object System.Text.StringBuilder
            $sw = [System.Diagnostics.Stopwatch]::StartNew()

            while ($sw.ElapsedMilliseconds -lt 2000) {
                try {
                    $chunk = $sp.ReadExisting()
                    if ($chunk -and $chunk.Length -gt 0) { [void]$builder.Append($chunk) }
                } catch { }
                Start-Sleep -Milliseconds 100
            }

            $resp = $builder.ToString()
            if ($resp) {
                $respLines = $resp -replace "`r`n", "`n" -split "`n"
                foreach ($line in $respLines) {
                    if ($line.Trim()) {
                        Add-Log "[$portName] RX: $line" $logColor
                    }
                }
            }
            else {
                Add-Log "[$portName] RX: (no data)" $logColor
            }

            $sp.Close()
        }
        catch {
            Add-Log "[$portName] Error: $($_.Exception.Message)" $logColor
        }
    }
}

$btnSend.Add_Click({
    Send-ATToTargets $txtCmd.Text
})

$btnCFUN0.Add_Click({
    Send-ATToTargets "AT+CFUN=0"
})

$btnCFUN1.Add_Click({
    Send-ATToTargets "AT+CFUN=1"
})

# Bind IP 변경 시 상단 라벨 갱신
$textBindIp1.Add_TextChanged({ Update-UeSummaryLabels })
$textBindIp2.Add_TextChanged({ Update-UeSummaryLabels })

# UE Swap 버튼
$btnSwapUE.Add_Click({
    # Bind IP 스왑
    $tmpIp = $textBindIp1.Text
    $textBindIp1.Text = $textBindIp2.Text
    $textBindIp2.Text = $tmpIp

    # 모뎀 인덱스 스왑
    $tmpIdx = $global:Ue1ModemIndex
    $global:Ue1ModemIndex = $global:Ue2ModemIndex
    $global:Ue2ModemIndex = $tmpIdx

    Update-UeSummaryLabels
    Add-Log "UE1 / UE2 mapping swapped."
})

# Ping 체크박스 이벤트
$checkPingUE1.Add_CheckedChanged({
    if ($checkPingUE1.Checked) {
        $server = $textServerIp.Text.Trim()
        $bind1  = $textBindIp1.Text.Trim()

        if (-not $server -or -not $bind1) {
            [System.Windows.Forms.MessageBox]::Show("Server IP or UE1 Bind IP is empty.", "Ping UE1 Error") | Out-Null
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
            } -ArgumentList $server, $bind1
            Add-PingLog "[UE1] ping started (Bind=$bind1, Server=$server)" $colorUE1
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to start ping UE1: $($_.Exception.Message)", "Ping UE1 Error") | Out-Null
            $checkPingUE1.Checked = $false
        }
    }
    else {
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
        $bind2  = $textBindIp2.Text.Trim()

        if (-not $server -or -not $bind2) {
            [System.Windows.Forms.MessageBox]::Show("Server IP or UE2 Bind IP is empty.", "Ping UE2 Error") | Out-Null
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
            } -ArgumentList $server, $bind2
            Add-PingLog "[UE2] ping started (Bind=$bind2, Server=$server)" $colorUE2
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to start ping UE2: $($_.Exception.Message)", "Ping UE2 Error") | Out-Null
            $checkPingUE2.Checked = $false
        }
    }
    else {
        if ($global:PingJobUE2) {
            try { Stop-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $global:PingJobUE2 -Force -ErrorAction SilentlyContinue } catch {}
            Add-PingLog "[UE2] ping stopped." $colorUE2
            $global:PingJobUE2 = $null
        }
    }
})

# Ping Job 출력 수집용 Timer
$pingTimer = New-Object System.Windows.Forms.Timer
$pingTimer.Interval = 1000
$pingTimer.Add_Tick({
    if ($global:PingJobUE1) {
        try {
            $out1 = Receive-Job -Job $global:PingJobUE1 -ErrorAction SilentlyContinue
            if ($out1) {
                foreach ($line in $out1) {
                    if ($line -and $line.Trim()) {
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
                    if ($line -and $line.Trim()) {
                        Add-PingLog "[UE2] $line" $colorUE2
                    }
                }
            }
        } catch {}
    }
})
$pingTimer.Start()

# 폼 Shown 이벤트
$form.Add_Shown({
    Refresh-NdisBindIps
    Detect-ModemPorts
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
