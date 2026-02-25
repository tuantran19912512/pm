# ==============================================================================
# [ TAB 4 ] SUA LOI MAY IN & MANG LAN (BAN FULL KET HOP IP SCANNER)
# ==============================================================================
$pnlPrint = New-Object System.Windows.Forms.Panel; $pnlPrint.Dock = "Fill"; $pnlPrint.Visible = $false
$lblPrint = New-Object System.Windows.Forms.Label; $lblPrint.Text = "SUA LOI MAY IN & MANG LAN"; $lblPrint.Font = $PhongChu.TieuDe; $lblPrint.ForeColor = $MauNen.XanhLa; $lblPrint.Size = New-Object System.Drawing.Size(600, 40); $lblPrint.Location = New-Object System.Drawing.Point(10, 10)

# GROUP 1: FIX LAN
$grLan = New-Object System.Windows.Forms.GroupBox; $grLan.Text = " 1. Sua Loi Ket Noi (Loi 0x0000011b, 0x00000709...) "; $grLan.ForeColor = [System.Drawing.Color]::Gray; $grLan.Size = New-Object System.Drawing.Size(740, 100); $grLan.Location = New-Object System.Drawing.Point(10, 60)
$btnFixLan = New-Object System.Windows.Forms.Button; $btnFixLan.Text = "FIX ALL (SMB 1.0, RPC, SMB, PRINTNIGHTMARE)"; $btnFixLan.Location = New-Object System.Drawing.Point(20, 35); $btnFixLan.Size = New-Object System.Drawing.Size(700, 45); ThietKeNut $btnFixLan $MauNen.Do
$grLan.Controls.Add($btnFixLan)

# GROUP 2: FIX DRIVER EPSON/BROTHER
$grDriver = New-Object System.Windows.Forms.GroupBox; $grDriver.Text = " 2. Sua Loi Driver & May Scan (Epson/Brother) "; $grDriver.ForeColor = [System.Drawing.Color]::Gray; $grDriver.Size = New-Object System.Drawing.Size(740, 100); $grDriver.Location = New-Object System.Drawing.Point(10, 170)
$btnFix3e3 = New-Object System.Windows.Forms.Button; $btnFix3e3.Text = "FIX LOI 3E3 (EPSON LQ-310)"; $btnFix3e3.Location = New-Object System.Drawing.Point(20, 35); $btnFix3e3.Size = New-Object System.Drawing.Size(340, 45); ThietKeNut $btnFix3e3 $MauNen.Cam
$btnFixBrother = New-Object System.Windows.Forms.Button; $btnFixBrother.Text = "FIX LOI SCAN BROTHER (BI VANG)"; $btnFixBrother.Location = New-Object System.Drawing.Point(380, 35); $btnFixBrother.Size = New-Object System.Drawing.Size(340, 45); ThietKeNut $btnFixBrother $MauNen.Do
$grDriver.Controls.AddRange(@($btnFix3e3, $btnFixBrother))

# GROUP 3: TIEN ICH MO RONG & QUET IP
$grUtil = New-Object System.Windows.Forms.GroupBox; $grUtil.Text = " 3. Tien Ich Mo Rong & Quet Mang LAN "; $grUtil.ForeColor = [System.Drawing.Color]::Gray; $grUtil.Size = New-Object System.Drawing.Size(740, 200); $grUtil.Location = New-Object System.Drawing.Point(10, 280)

$btnFwOn = New-Object System.Windows.Forms.Button; $btnFwOn.Text = "BAT TUONG LUA"; $btnFwOn.Location = New-Object System.Drawing.Point(20, 35); $btnFwOn.Size = New-Object System.Drawing.Size(220, 40); ThietKeNut $btnFwOn $MauNen.XanhLa
$btnFwOff = New-Object System.Windows.Forms.Button; $btnFwOff.Text = "TAT TUONG LUA"; $btnFwOff.Location = New-Object System.Drawing.Point(260, 35); $btnFwOff.Size = New-Object System.Drawing.Size(220, 40); ThietKeNut $btnFwOff $MauNen.Do
$btnResetSpool = New-Object System.Windows.Forms.Button; $btnResetSpool.Text = "RESET SPOOLER"; $btnResetSpool.Location = New-Object System.Drawing.Point(500, 35); $btnResetSpool.Size = New-Object System.Drawing.Size(220, 40); ThietKeNut $btnResetSpool $MauNen.NutMacDinh

$btnAddCred = New-Object System.Windows.Forms.Button; $btnAddCred.Text = "THEM CREDENTIAL (GUEST)"; $btnAddCred.Location = New-Object System.Drawing.Point(20, 85); $btnAddCred.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnAddCred $MauNen.NutMacDinh
$btnSysRes = New-Object System.Windows.Forms.Button; $btnSysRes.Text = "MO SYSTEM RESTORE"; $btnSysRes.Location = New-Object System.Drawing.Point(380, 85); $btnSysRes.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnSysRes $MauNen.XanhDuong

$btnQuetIP = New-Object System.Windows.Forms.Button; $btnQuetIP.Text = "QUET IP MANG LAN & DIA CHI MAC (ADVANCED IP SCANNER)"; $btnQuetIP.Location = New-Object System.Drawing.Point(20, 135); $btnQuetIP.Size = New-Object System.Drawing.Size(700, 45); ThietKeNut $btnQuetIP $MauNen.XanhDuong

$grUtil.Controls.AddRange(@($btnFwOn, $btnFwOff, $btnResetSpool, $btnAddCred, $btnSysRes, $btnQuetIP))

$pnlPrint.Controls.AddRange(@($lblPrint, $grLan, $grDriver, $grUtil))
$khungChinh.Controls.Add($pnlPrint)

# ==============================================================================
# LOGIC XU LY MAY IN & LAN
# ==============================================================================
$backupDir = "C:\Tool_Backups"
if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

$btnFixLan.Add_Click({
    ChayTacVu "Dang Fix loi LAN 11b/0x709 & SMB1..." {
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> BAT DAU FIX LOI LAN..."

        GhiLog "-> Dang bat SMB 1.0 (Can thiet cho may in cu/WinXP)..."
        try {
            Enable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart -ErrorAction Stop | Out-Null
            GhiLog "   + Da bat SMB 1.0 thanh cong."
        } catch {
            GhiLog "   ! Loi bat SMB 1.0: $($_.Exception.Message)"
        }

        GhiLog "-> Dang cau hinh AllowInsecureGuestAuth..."
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" -Name "AllowInsecureGuestAuth" -Value 1 -Type DWord -Force | Out-Null
        
        GhiLog "-> Dang cau hinh Registry chong loi PrintNightmare..."
        reg export "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers" "$backupDir\Printers_Policies.reg" /y 2>&1 | Out-Null
        reg export "HKLM\SYSTEM\CurrentControlSet\Control\Print" "$backupDir\Print.reg" /y 2>&1 | Out-Null
        
        reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC" /v RpcUseNamedPipeProtocol /t REG_DWORD /d 1 /f | Out-Null
        reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC" /v RpcTcpPort /t REG_DWORD /d 0 /f | Out-Null
        reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC" /v RpcAuthentication /t REG_DWORD /d 0 /f | Out-Null
        reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint" /v RestrictDriverInstallationToAdministrators /t REG_DWORD /d 1 /f | Out-Null
        reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print" /v RpcAuthnLevelPrivacyEnabled /t REG_DWORD /d 0 /f | Out-Null
        reg add "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" /v LimitBlankPasswordUse /t REG_DWORD /d 0 /f | Out-Null
        reg add "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" /v EveryoneIncludesAnonymous /t REG_DWORD /d 1 /f | Out-Null
        
        Set-SmbClientConfiguration -RequireSecuritySignature $false -Force -Confirm:$false 2>&1 | Out-Null
        Set-SmbServerConfiguration -RequireSecuritySignature $false -Force -Confirm:$false 2>&1 | Out-Null
        
        GhiLog "-> Dang khoi dong lai Spooler..."
        Restart-Service -Name Spooler -Force -ErrorAction SilentlyContinue
        
        GhiLog ">>> HOAN TAT! Vui long khoi dong lai may."
        [System.Windows.Forms.MessageBox]::Show("Da Fix xong (SMB1 + LAN)! Hay khoi dong lai may de ap dung.", "Thong bao")
    }
})

$btnFix3e3.Add_Click({ 
    ChayTacVu "Dang tim Driver Epson..." { 
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> BAT DAU FIX LOI 3E3 (EPSON LQ-310)..."
        $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
        $driverFound = $null
        
        if (Test-Path "$basePath\Epson LQ-310 ESC") { $driverFound = "Epson LQ-310 ESC" } 
        elseif (Test-Path "$basePath\Epson LQ-310 ESC/P2") { $driverFound = "Epson LQ-310 ESC/P2" }
        
        if ($driverFound) { 
            GhiLog "- Tim thay driver: $driverFound"
            reg export "$basePath\$driverFound" "$backupDir\Epson_Backup.reg" /y 2>&1 | Out-Null
            Set-ItemProperty -Path "$basePath\$driverFound" -Name "PrinterDriverAttributes" -Value 1 -Force
            Restart-Service -Name Spooler -Force -ErrorAction SilentlyContinue
            GhiLog ">>> DA FIX XONG! Hay in thu."
            [System.Windows.Forms.MessageBox]::Show("Da sua loi Epson 3e3 thanh cong!", "Thong bao") 
        } else { 
            GhiLog ">>> LOI: Khong tim thay Driver Epson LQ-310 tren may nay."
            [System.Windows.Forms.MessageBox]::Show("Khong tim thay Driver Epson LQ-310. Vui long cai driver truoc.", "Loi") 
        } 
    } 
})

$btnFixBrother.Add_Click({ 
    ChayTacVu "Fix loi Scan Brother" { 
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog "--- FIX LOI SCAN BROTHER (CRASH / BI VANG) ---"
        $key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        Set-ItemProperty -Path $key -Name "Personal" -Value "%USERPROFILE%\Documents" -Type ExpandString -ErrorAction SilentlyContinue
        GhiLog "1. Da Reset duong dan 'Personal' ve mac dinh."
        GhiLog ">>> Hoan tat."
        [System.Windows.Forms.MessageBox]::Show("Da Reset duong dan User Shell Folders.`nVui long Log out (Dang xuat) hoac khoi dong lai may de ap dung.", "Thanh cong") 
    } 
})

$btnResetSpool.Add_Click({ 
    ChayTacVu "Reset Print Spooler" { 
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog "--- RESET PRINT SPOOLER ---"
        Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
        GhiLog "1. Da dung dich vu Spooler"
        $pathSpool = "$env:SystemRoot\System32\spool\PRINTERS\*"
        Remove-Item $pathSpool -Force -Recurse -ErrorAction SilentlyContinue
        GhiLog "2. Da xoa cache/lenh in bi ket"
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        GhiLog "3. Da khoi dong lai Spooler"
        GhiLog ">>> Hoan tat."
        [System.Windows.Forms.MessageBox]::Show("Da Reset Spooler & Xoa lenh in ket!", "Thanh cong") 
    } 
})

$btnFwOn.Add_Click({ 
    ChayTacVu "Bat Tuong Lua" { 
        netsh advfirewall set allprofiles state on | Out-Null
        [System.Windows.Forms.MessageBox]::Show("Da BAT Tuong lua (Firewall).", "Thong bao") 
    } 
})

$btnFwOff.Add_Click({ 
    ChayTacVu "Tat Tuong Lua" { 
        netsh advfirewall set allprofiles state off | Out-Null
        [System.Windows.Forms.MessageBox]::Show("Da TAT Tuong lua (Firewall).", "Thong bao") 
    } 
})

$btnAddCred.Add_Click({ 
    $ip = HienThiInputBox "Them Credential" "Nhap IP hoac Ten May Chu (VD: 192.168.1.10):"
    if (-not [string]::IsNullOrWhiteSpace($ip)) { 
        cmdkey /add:$ip /user:Guest | Out-Null
        [System.Windows.Forms.MessageBox]::Show("Da them thong tin truy cap (User Guest) cho IP: $ip", "Thanh cong") 
    } 
})

$btnSysRes.Add_Click({ Start-Process rstrui.exe })

# ==============================================================================
# LOGIC: QUET IP MANG LAN (BAN PRO: HO TRO QUET NHIEU LOP MANG CUNG LUC)
# ==============================================================================
$btnQuetIP.Add_Click({
    ChayTacVu "Dang quet mang LAN..." {
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> BAT DAU QUET MANG LAN (MULTI-SUBNET) <<<"

        # 1. Tu dong nhan dien TAT CA cac lop mang may dang ket noi
        $MyIPs = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch "Loopback|Pseudo" } | Select-Object -ExpandProperty IPAddress
        if (-not $MyIPs) { 
            GhiLog "!!! Loi: Khong tim thay Card mang nao dang ket noi."
            return 
        }
        
        $DetectedSubnets = ($MyIPs | ForEach-Object { $_.Substring(0, $_.LastIndexOf('.')) }) | Select-Object -Unique
        $DefaultText = $DetectedSubnets -join ", "

        # 2. Hien thi bang Input giong thanh cong cu cua Advanced IP Scanner
        Add-Type -AssemblyName Microsoft.VisualBasic
        $UserInput = [Microsoft.VisualBasic.Interaction]::InputBox("Phat hien may ban dang o cac lop mang ben duoi.`nBan co the them hoac xoa lop mang (Cach nhau bang dau phay).`n`nVD: 192.168.1, 10.0.0, 172.16.0", "Nhap lop mang can quet", $DefaultText)

        if ([string]::IsNullOrWhiteSpace($UserInput)) { 
            GhiLog ">>> DA HUY QUET MANG <<<"
            return 
        }

        # Tach cac lop mang va kiem tra dinh dang (Chi lay phan dau VD: 192.168.1)
        $SubnetsToScan = $UserInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d{1,3}\.\d{1,3}\.\d{1,3}$" }

        if ($SubnetsToScan.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Dinh dang lop mang khong hop le! Vui long nhap 3 cum so dau, VD: 192.168.1", "Loi")
            return
        }

        GhiLog "-> IP cua ban dang co: $($MyIPs -join ' | ')"
        GhiLog "-> Dang quet $($SubnetsToScan.Count) dải mang: $($SubnetsToScan -join ', ')"
        GhiLog "-> [1/4] Dang ban dong loat $($SubnetsToScan.Count * 254) goi tin Ping..."

        # 3. PING NHANH GOM CHUNG TAT CA CAC LOP MANG
        $Pingers = @()
        foreach ($Subnet in $SubnetsToScan) {
            for ($i = 1; $i -le 254; $i++) {
                $ip = "$Subnet.$i"
                $ping = New-Object System.Net.NetworkInformation.Ping
                $Pingers += [PSCustomObject]@{ IP = $ip; Task = $ping.SendPingAsync($ip, 500); Ping = $ping }
            }
        }

        $ThoiGianChoPing = [DateTime]::Now.AddSeconds(3.5)
        $PingXong = $false
        while ((-not $PingXong) -and ([DateTime]::Now -lt $ThoiGianChoPing)) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
            $PingXong = $true
            foreach ($p in $Pingers) { if (-not $p.Task.IsCompleted) { $PingXong = $false; break } }
        }

        $ActiveIPs = @()
        foreach ($p in $Pingers) {
            if ($p.Task.IsCompleted -and $p.Task.Status -eq 'RanToCompletion') {
                if ($p.Task.Result.Status -eq 'Success') { $ActiveIPs += $p.IP }
            }
            $p.Ping.Dispose()
        }

        # 4. DOC TROM BANG ARP DE BAT MAY CHAN PING (Ap dung cho tat ca dải)
        GhiLog "-> [2/4] Dang quet bang ARP tim may an minh..."
        $arpOutput = arp -a
        foreach ($line in $arpOutput) {
            foreach ($Subnet in $SubnetsToScan) {
                if ($line -match "^\s*($Subnet\.\d+)\s+([0-9a-fA-F-]{17})\s+(dynamic|static)") {
                    $arpIp = $Matches[1]
                    $arpMac = $Matches[2]
                    if ($arpMac -notmatch "ff-ff-ff-ff-ff-ff" -and $arpIp -ne "$Subnet.255") {
                        if ($ActiveIPs -notcontains $arpIp) { $ActiveIPs += $arpIp }
                    }
                }
            }
        }
        
        foreach ($ip in $MyIPs) { if ($ActiveIPs -notcontains $ip) { $ActiveIPs += $ip } }

        GhiLog "-> Phat hien TONG CONG $($ActiveIPs.Count) thiet bi."
        GhiLog "-> [3/4] Dang tra cuu Ten may qua DNS Server..."

        # 5. RESOLVE HOSTNAME (DNS)
        $DnsTasks = @()
        foreach ($ip in $ActiveIPs) {
            $DnsTasks += [PSCustomObject]@{ IP = $ip; Task = [System.Net.Dns]::GetHostEntryAsync($ip) }
        }

        $ThoiGianChoDns = [DateTime]::Now.AddSeconds(4)
        $DnsXong = $false
        while ((-not $DnsXong) -and ([DateTime]::Now -lt $ThoiGianChoDns)) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
            $DnsXong = $true
            foreach ($t in $DnsTasks) { if (-not $t.Task.IsCompleted) { $DnsXong = $false; break } }
        }

        $HostNames = @{}
        $ThieuTenIPs = @()
        foreach ($t in $DnsTasks) {
            if ($t.Task.IsCompleted -and $t.Task.Status -eq 'RanToCompletion') {
                $HostNames[$t.IP] = $t.Task.Result.HostName.Split('.')[0]
            } else {
                $HostNames[$t.IP] = $t.IP 
                $ThieuTenIPs += $t.IP
            }
        }

        # 6. QUET NETBIOS NHUNG MAY BI AN TEN
        if ($ThieuTenIPs.Count -gt 0) {
            GhiLog "-> [4/4] Dang ep tim Ten may (NetBIOS) cho $($ThieuTenIPs.Count) thiet bi an..."
            $NbtTasks = @()
            foreach ($ip in $ThieuTenIPs) {
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo.FileName = "nbtstat"
                $p.StartInfo.Arguments = "-A $ip"
                $p.StartInfo.UseShellExecute = $false
                $p.StartInfo.RedirectStandardOutput = $true
                $p.StartInfo.CreateNoWindow = $true
                $p.Start() | Out-Null
                $NbtTasks += [PSCustomObject]@{ IP = $ip; Process = $p }
            }

            $NbtTimeout = [DateTime]::Now.AddSeconds(3)
            $NbtDone = $false
            while ((-not $NbtDone) -and ([DateTime]::Now -lt $NbtTimeout)) {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 100
                $NbtDone = $true
                foreach ($t in $NbtTasks) { if (-not $t.Process.HasExited) { $NbtDone = $false; break } }
            }

            foreach ($t in $NbtTasks) {
                if ($t.Process.HasExited) {
                    $out = $t.Process.StandardOutput.ReadToEnd()
                    if ($out -match "(?m)^\s*(\S+)\s*<00>\s*UNIQUE") {
                        $TenNetBios = $Matches[1].Trim()
                        if ($TenNetBios -notmatch "ISATAP") { $HostNames[$t.IP] = $TenNetBios }
                    }
                } else {
                    $t.Process.Kill()
                }
                $t.Process.Dispose()
            }
        }
        
        # 7. TU DIEN MAC ADDRESS (OUI) NHAN DIEN HANG SAN XUAT
        $MacVendors = @{
            "00:1D:AA" = "DrayTek Corp"; "2C:FD:A1" = "ASUSTek"; "B0:6E:BF" = "ASUSTek"
            "00:15:65" = "YEALINK"; "80:5E:C0" = "YEALINK"; "48:BA:4E" = "HP Printer"
            "00:1E:8F" = "Canon Printer"; "3C:2A:F4" = "Brother Printer"; "E0:50:8B" = "Hikvision Camera"
            "18:D6:C7" = "TP-Link"; "F4:F2:6D" = "TP-Link"; "00:1A:A0" = "Dell"
            "00:23:DF" = "Apple (iPhone/Mac)"; "00:26:73" = "Samsung"; "10:AE:60" = "Cisco"
        }

        # 8. DOC BANG ARP VA IN KET QUA RA BANG (SAP XEP THEO IP)
        $arpOutput = arp -a
        GhiLog " "
        GhiLog "=========================================================================================="
        GhiLog " IP ADDRESS        MAC ADDRESS         HOSTNAME                  HANG SAN XUAT (NOTE)     "
        GhiLog "=========================================================================================="
        
        # Sap xep danh sach IP (theo tung lop mang tang dan)
        $SortedIPs = $ActiveIPs | Sort-Object { [version]($_) }

        foreach ($ip in $SortedIPs) {
            $mac = "Khong xac dinh   "
            $vendor = "Unknown / Khac Subnet"
            
            $arpLine = $arpOutput | Select-String -Pattern "\s$ip\s"
            if ($arpLine) {
                if ($arpLine.Line -match "([0-9a-fA-F]{2}[:-]){5}([0-9a-fA-F]{2})") {
                    $mac = $Matches[0].ToUpper().Replace("-",":")
                    $macPrefix = $mac.Substring(0,8)
                    if ($MacVendors.ContainsKey($macPrefix)) { $vendor = $MacVendors[$macPrefix] }
                    else { $vendor = "Unknown Vendor" }
                }
            }
            
            $GhiChu = "ON"
            if ($MyIPs -contains $ip) { 
                $mac = "(May cua ban)    "
                $vendor = "This PC (Local)" 
            }
            
            $ipPad = $ip.PadRight(16)
            $macPad = $mac.PadRight(19)
            
            $hostStr = $HostNames[$ip]
            if ($hostStr.Length -gt 23) { $hostStr = $hostStr.Substring(0, 20) + "..." }
            $hostPad = $hostStr.PadRight(25)
            
            $vendorPad = "$vendor"
            
            GhiLog " $ipPad $macPad $hostPad $vendorPad"
        }
        GhiLog "=========================================================================================="
        GhiLog ">>> QUET MANG HOAN TAT! <<<"
        [System.Windows.Forms.MessageBox]::Show("Da quet xong mang LAN! Tim thay $($ActiveIPs.Count) thiet bi.", "Thanh cong")
    }
})