Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global lists
$global:psList = @()       # iperf3 PowerShell 윈도우 PID 리스트
$global:ModemList = @()    # 모뎀 리스트

# ----- Main Form -----
$form = New-Object System.Windows.Forms.Form
$form.Text = "2UE iperf3 + Modem AT All-in-One"
$form.ClientSize = New-Object System.Drawing.Size(1000, 750)   # 넉넉하게
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$form.ForeColor = [System.Drawing.Color]::White
$form.AutoScroll = $true                                       # 스크롤 가능

# ----- 공통 스타일 -----
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

# =========================================================
# 1) 왼쪽 영역: iperf3 2-UE Bidirectional Throughput
# =========================================================
[int]$labelX = 20
[int]$inputX = 190
[int]$widthLabel = 160
[int]$widthInput = 240
[int]$height = 23
[int]$gap = 8
[int]$y = 20

# ----- Server IP -----
$labelServerIp = New-Object System.Windows.Forms.Label
$labelServerIp.Text = "Server IP"
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

$y += $height + $gap

# ----- UE1 / UE2 Enable Checkboxes -----
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

# ----- DL1 Port (Bind1) -----
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

# ----- UL1 Port (Bind1) -----
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

# ----- DL2 Port (Bind2) -----
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

# ----- UL2 Port (Bind2) -----
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

# ----- Download BW -----
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

# ----- Upload BW -----
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

# ----- Bind IP #1 -----
$labelBindIp1 = New-Object System.Windows.Forms.Label
$labelBindIp1.Text = "Bind IP #1 (UE1)"
$labelBindIp1.Location = New-Object System.Drawing.Point($labelX, $y)
$labelBindIp1.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelBindIp1
$form.Controls.Add($labelBindIp1)

$textBindIp1 = New-Object System.Windows.Forms.TextBox
$textBindIp1.Location = New-Object System.Drawing.Point($inputX, $y)
$textBindIp1.Size = New-Object System.Drawing.Size($widthInput, $height)
$textBindIp1.Text = "192.168.225.30"
Set-DarkTextBoxStyle $textBindIp1
$form.Controls.Add($textBindIp1)

$y += $height + $gap

# ----- Bind IP #2 -----
$labelBindIp2 = New-Object System.Windows.Forms.Label
$labelBindIp2.Text = "Bind IP #2 (UE2)"
$labelBindIp2.Location = New-Object System.Drawing.Point($labelX, $y)
$labelBindIp2.Size = New-Object System.Drawing.Size($widthLabel, $height)
Set-DarkLabelStyle $labelBindIp2
$form.Controls.Add($labelBindIp2)

$textBindIp2 = New-Object System.Windows.Forms.TextBox
$textBindIp2.Location = New-Object System.Drawing.Point($inputX, $y)
$textBindIp2.Size = New-Object System.Drawing.Size($widthInput, $height)
$textBindIp2.Text = "192.168.225.82"
Set-DarkTextBoxStyle $textBindIp2
$form.Controls.Add($textBindIp2)

$y += $height + $gap

# =========================================================
#   Remote NDIS IP List (자동 검색 + 선택 적용)
# =========================================================
[int]$ndisTopY = $y

$labelNdis = New-Object System.Windows.Forms.Label
$labelNdis.Text = "Remote NDIS Compatible Device IPv4:"
$labelNdis.Location = New-Object System.Drawing.Point($labelX, $ndisTopY)
$labelNdis.Size = New-Object System.Drawing.Size(260, $height)
Set-DarkLabelStyle $labelNdis
$form.Controls.Add($labelNdis)

$listNdis = New-Object System.Windows.Forms.ListBox
$listNdis.Location = New-Object System.Drawing.Point($labelX, ($ndisTopY + $height))
$listNdis.Size = New-Object System.Drawing.Size(370, 100)
$listNdis.BackColor = $textBgColor
$listNdis.ForeColor = $textFgColor
$form.Controls.Add($listNdis)

$btnNdisRefresh = New-Object System.Windows.Forms.Button
$btnNdisRefresh.Text = "Refresh NDIS IPs"
$btnNdisRefresh.Location = New-Object System.Drawing.Point($labelX, ($ndisTopY + $height + 110))
$btnNdisRefresh.Size = New-Object System.Drawing.Size(120, 30)
Set-DarkButtonStyle $btnNdisRefresh
$form.Controls.Add($btnNdisRefresh)

$btnApplyUe1 = New-Object System.Windows.Forms.Button
$btnApplyUe1.Text = "Apply to UE1"
$btnApplyUe1.Location = New-Object System.Drawing.Point(($labelX + 130), ($ndisTopY + $height + 110))
$btnApplyUe1.Size = New-Object System.Drawing.Size(100, 30)
Set-DarkButtonStyle $btnApplyUe1
$form.Controls.Add($btnApplyUe1)

$btnApplyUe2 = New-Object System.Windows.Forms.Button
$btnApplyUe2.Text = "Apply to UE2"
$btnApplyUe2.Location = New-Object System.Drawing.Point(($labelX + 240), ($ndisTopY + $height + 110))
$btnApplyUe2.Size = New-Object System.Drawing.Size(100, 30)
Set-DarkButtonStyle $btnApplyUe2
$form.Controls.Add($btnApplyUe2)

$y = $ndisTopY + $height + 110 + 40 + (2 * $gap)

# ----- Duration -----
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

# ----- Protocol -----
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

# ----- Route Button (Admin CMD) -----
$buttonRouteAdmin = New-Object System.Windows.Forms.Button
$buttonRouteAdmin.Text = "Add Route in Admin CMD (use Gateway of Bind IP)"
$buttonRouteAdmin.Location = New-Object System.Drawing.Point(40, $y)
$buttonRouteAdmin.Size = New-Object System.Drawing.Size(370, 30)
Set-DarkButtonStyle $buttonRouteAdmin
$form.Controls.Add($buttonRouteAdmin)

$y += 30 + $gap

# ----- iperf Buttons -----
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

# =========================================================
#   Remote NDIS IP 탐색 함수
# =========================================================
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

            $ipv4s = $cfg.IPAddress |
                     Where-Object {
                         $_ -match '^\d+\.\d+\.\d+\.\d+$' -and
                         $_ -notlike "169.254.*"
                     }

            foreach ($ip in $ipv4s) {
                if ($ips -notcontains $ip) {
                    $ips += $ip
                }
            }
        }
    }
    catch {
    }

    return $ips
}

function Refresh-NdisList {
    $listNdis.Items.Clear()
    $ips = Get-NdisIpList

    if (-not $ips -or $ips.Count -eq 0) {
        [void]$listNdis.Items.Add("No Remote NDIS IP found")
        return
    }

    foreach ($ip in $ips) {
        [void]$listNdis.Items.Add($ip)
    }

    if ($ips.Count -ge 1) { $textBindIp1.Text = $ips[0] }
    if ($ips.Count -ge 2) { $textBindIp2.Text = $ips[1] }
}

# =========================================================
#   특정 IP가 붙어있는 NIC의 Gateway 찾기 함수
# =========================================================
function Get-GatewayForIp {
    param(
        [string]$IpAddress
    )

    if (-not $IpAddress) { return $null }

    try {
        $cfgs = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue |
                Where-Object { $_.IPEnabled }

        foreach ($cfg in $cfgs) {
            if ($cfg.IPAddress -and ($cfg.IPAddress -contains $IpAddress)) {
                if ($cfg.DefaultIPGateway -and $cfg.DefaultIPGateway.Length -gt 0) {
                    return $cfg.DefaultIPGateway[0]
                }
            }
        }
    }
    catch {
    }

    return $null
}

# ----- iperf Events -----
$buttonClose.Add_Click({
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
    $bind1 = $textBindIp1.Text.Trim()
    $bind2 = $textBindIp2.Text.Trim()
    $t = $textTime.Text.Trim()
    $proto = $comboProto.SelectedItem
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

$buttonStop.Add_Click({
    try {
        $iperf = Get-Process iperf3 -ErrorAction SilentlyContinue
        if ($iperf) {
            $iperf | Stop-Process -Force -ErrorAction SilentlyContinue
        }

        foreach ($psId in $global:psList) {
            try {
                Stop-Process -Id $psId -Force -ErrorAction SilentlyContinue
            } catch {}
        }

        $global:psList = @()

        [System.Windows.Forms.MessageBox]::Show("All iperf3 sessions and windows have been closed.", "iperf3") | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error while stopping sessions: $($_.Exception.Message)", "Error") | Out-Null
    }
})

# ----- Route (Gateway 자동 적용, 관리자 CMD) -----
$buttonRouteAdmin.Add_Click({
    $server = $textServerIp.Text.Trim()
    $bind1  = $textBindIp1.Text.Trim()
    $bind2  = $textBindIp2.Text.Trim()

    if (-not $server) {
        [System.Windows.Forms.MessageBox]::Show("Server IP is empty.", "Route Error") | Out-Null
        return
    }

    $cmdParts = @()

    # UE1용
    if ($checkUE1.Checked -and $bind1) {
        $gw1 = Get-GatewayForIp -IpAddress $bind1
        if ($gw1) {
            $cmdParts += "route ADD $server MASK 255.255.255.255 $gw1"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Could not find gateway for UE1 Bind IP ($bind1).", "Route Warning (UE1)") | Out-Null
        }
    }

    # UE2용
    if ($checkUE2.Checked -and $bind2) {
        $gw2 = Get-GatewayForIp -IpAddress $bind2
        if ($gw2) {
            $cmdParts += "route ADD $server MASK 255.255.255.255 $gw2"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Could not find gateway for UE2 Bind IP ($bind2).", "Route Warning (UE2)") | Out-Null
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

# =========================================================
# 2) 오른쪽 영역: Modem AT Command Tool (Multi-port)
# =========================================================
[int]$baseX = 520
[int]$y2 = 20

$labelList = New-Object System.Windows.Forms.Label
$labelList.Text = "Detected Modem Ports (Win32_POTSModem):"
$labelList.AutoSize = $true
$labelList.Location = New-Object System.Drawing.Point($baseX, $y2)
Set-DarkLabelStyle $labelList
$form.Controls.Add($labelList)

$listPorts = New-Object System.Windows.Forms.ListBox
$listPorts.Location = New-Object System.Drawing.Point($baseX, ($y2 + 25))
$listPorts.Size = New-Object System.Drawing.Size(380, 150)
$listPorts.BackColor = $textBgColor
$listPorts.ForeColor = $textFgColor
$listPorts.SelectionMode = "MultiExtended"
$form.Controls.Add($listPorts)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point($baseX, ($y2 + 185))
$statusLabel.ForeColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($statusLabel)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Location = New-Object System.Drawing.Point($baseX, ($y2 + 215))
$btnRefresh.Size = New-Object System.Drawing.Size(100, 30)
Set-DarkButtonStyle $btnRefresh
$form.Controls.Add($btnRefresh)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(($baseX + 120), ($y2 + 215))
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
Set-DarkButtonStyle $btnSelectAll
$form.Controls.Add($btnSelectAll)

$btnCloseModem = New-Object System.Windows.Forms.Button
$btnCloseModem.Text = "Exit"
$btnCloseModem.Location = New-Object System.Drawing.Point(($baseX + 240), ($y2 + 215))
$btnCloseModem.Size = New-Object System.Drawing.Size(100, 30)
Set-DarkButtonStyle $btnCloseModem
$form.Controls.Add($btnCloseModem)

$grpAt = New-Object System.Windows.Forms.GroupBox
$grpAt.Text = "AT Command (broadcast to all selected modems)"
$grpAt.Location = New-Object System.Drawing.Point($baseX, ($y2 + 260))
$grpAt.Size = New-Object System.Drawing.Size(380, 300)
$grpAt.ForeColor = $labelColor
$form.Controls.Add($grpAt)

$lblSelected = New-Object System.Windows.Forms.Label
$lblSelected.Text = "Selected modems: 0"
$lblSelected.AutoSize = $true
$lblSelected.Location = New-Object System.Drawing.Point(15, 25)
$grpAt.Controls.Add($lblSelected)

$lblCmd = New-Object System.Windows.Forms.Label
$lblCmd.Text = "AT command:"
$lblCmd.AutoSize = $true
$lblCmd.Location = New-Object System.Drawing.Point(15, 55)
$grpAt.Controls.Add($lblCmd)

$txtCmd = New-Object System.Windows.Forms.TextBox
$txtCmd.Location = New-Object System.Drawing.Point(15, 75)
$txtCmd.Size = New-Object System.Drawing.Size(200, 23)
$txtCmd.BackColor = $textBgColor
$txtCmd.ForeColor = $textFgColor
$grpAt.Controls.Add($txtCmd)

$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text = "Send"
$btnSend.Location = New-Object System.Drawing.Point(225, 73)
$btnSend.Size = New-Object System.Drawing.Size(60, 26)
Set-DarkButtonStyle $btnSend
$grpAt.Controls.Add($btnSend)

$btnCFUN0 = New-Object System.Windows.Forms.Button
$btnCFUN0.Text = "CFUN=0"
$btnCFUN0.Location = New-Object System.Drawing.Point(295, 73)
$btnCFUN0.Size = New-Object System.Drawing.Size(70, 26)
Set-DarkButtonStyle $btnCFUN0
$grpAt.Controls.Add($btnCFUN0)

$btnCFUN1 = New-Object System.Windows.Forms.Button
$btnCFUN1.Text = "CFUN=1"
$btnCFUN1.Location = New-Object System.Drawing.Point(295, 105)
$btnCFUN1.Size = New-Object System.Drawing.Size(70, 26)
Set-DarkButtonStyle $btnCFUN1
$grpAt.Controls.Add($btnCFUN1)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.AutoSize = $true
$lblLog.Location = New-Object System.Drawing.Point(15, 110)
$grpAt.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 130)
$txtLog.Size = New-Object System.Drawing.Size(350, 150)
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.BackColor = $textBgColor
$txtLog.ForeColor = $textFgColor
$grpAt.Controls.Add($txtLog)

$grpAt.Enabled = $false

function Add-Log {
    param([string]$Message)
    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("$time  $Message`r`n")
}

function Update-SelectedCount {
    $count = $listPorts.SelectedIndices.Count
    $lblSelected.Text = "Selected modems: $count"
}

function Detect-ModemPorts {
    $listPorts.Items.Clear()
    $txtLog.Clear()
    $txtCmd.Clear()
    $global:ModemList = @()
    $lblSelected.Text = "Selected modems: 0"
    $statusLabel.Text = ""
    $grpAt.Enabled = $false

    try {
        $modems = Get-CimInstance Win32_POTSModem -ErrorAction Stop
    }
    catch {
        $modems = Get-WmiObject Win32_POTSModem -ErrorAction SilentlyContinue
    }

    if (-not $modems) {
        $listPorts.Items.Add("No modem ports found (Win32_POTSModem empty)")
        $statusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - No modem ports found"
        return
    }

    $modems = $modems | Where-Object { $_.AttachedTo -match "^COM\d+" }

    if (-not $modems -or $modems.Count -eq 0) {
        $listPorts.Items.Add("No modem ports with COM attached")
        $statusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - No modem COM ports found"
        return
    }

    foreach ($m in $modems) {
        $display = "$($m.Name) ($($m.AttachedTo))"
        $listPorts.Items.Add($display)
        $global:ModemList += $m
    }

    $statusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - $($modems.Count) modem port(s) detected"
    $grpAt.Enabled = $true
}

function Send-ATCommand {
    param([string]$Command)

    $selectedIndices = $listPorts.SelectedIndices
    if ($selectedIndices.Count -eq 0) {
        Add-Log "No modem selected"
        return
    }

    $cmd = $Command.Trim()
    if ($cmd -eq "") {
        Add-Log "No command to send"
        return
    }

    $send = $cmd + "`r`n"

    foreach ($idx in $selectedIndices) {
        if ($idx -lt 0 -or $idx -ge $global:ModemList.Count) { continue }

        $modem = $global:ModemList[$idx]
        $portName = $modem.AttachedTo

        try {
            $sp = New-Object System.IO.Ports.SerialPort $portName, 115200, "None", 8, "One"
            $sp.Handshake = "None"
            $sp.ReadTimeout = 1000
            $sp.WriteTimeout = 1000
            $sp.DtrEnable = $true
            $sp.RtsEnable = $true
            $sp.NewLine = "`r`n"

            $sp.Open()

            Add-Log "TX ${portName}: $cmd"

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
                } catch { }
                Start-Sleep -Milliseconds 100
            }

            $resp = $builder.ToString()

            if ($resp) {
                $respLines = $resp -replace "`r`n", "`n" -split "`n"
                foreach ($line in $respLines) {
                    if ($line.Trim() -ne "") {
                        Add-Log "RX ${portName}: $line"
                    }
                }
            }
            else {
                Add-Log "RX ${portName}: (no data)"
            }

            $sp.Close()
        }
        catch {
            Add-Log "Error on ${portName}: $($_.Exception.Message)"
        }
    }
}

$btnRefresh.Add_Click({ Detect-ModemPorts })
$btnCloseModem.Add_Click({ $form.Close() })

$btnSelectAll.Add_Click({
    for ($i = 0; $i -lt $listPorts.Items.Count; $i++) {
        $listPorts.SetSelected($i, $true)
    }
    Update-SelectedCount
})

$listPorts.Add_SelectedIndexChanged({ Update-SelectedCount })

$btnSend.Add_Click({
    Send-ATCommand $txtCmd.Text
})

$btnCFUN0.Add_Click({
    Send-ATCommand "AT+CFUN=0"
})

$btnCFUN1.Add_Click({
    Send-ATCommand "AT+CFUN=1"
})

$btnNdisRefresh.Add_Click({ Refresh-NdisList })

$btnApplyUe1.Add_Click({
    $sel = $listNdis.SelectedItem
    if ($sel -and ($sel -notlike "No Remote NDIS IP*")) {
        $textBindIp1.Text = $sel
    }
})

$btnApplyUe2.Add_Click({
    $sel = $listNdis.SelectedItem
    if ($sel -and ($sel -notlike "No Remote NDIS IP*")) {
        $textBindIp2.Text = $sel
    }
})

$form.Add_Shown({
    Detect-ModemPorts
    Refresh-NdisList
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
