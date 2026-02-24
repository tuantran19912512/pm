# ==============================================================================
# [ TAB 4 ] SỬA LỖI MÁY IN & LAN
# ==============================================================================
$pnlPrint = New-Object System.Windows.Forms.Panel; $pnlPrint.Dock = "Fill"; $pnlPrint.Visible = $false
$lblPrint = New-Object System.Windows.Forms.Label; $lblPrint.Text = "SỬA LỖI MÁY IN & MẠNG LAN"; $lblPrint.Font = $PhongChu.TieuDe; $lblPrint.ForeColor = $MauNen.XanhLa; $lblPrint.Size = New-Object System.Drawing.Size(600, 40); $lblPrint.Location = New-Object System.Drawing.Point(10, 10)

$grLan = New-Object System.Windows.Forms.GroupBox; $grLan.Text = " 1. Sửa Lỗi Kết Nối (Lỗi 0x0000011b, 0x00000709...) "; $grLan.ForeColor = [System.Drawing.Color]::Gray; $grLan.Size = New-Object System.Drawing.Size(740, 100); $grLan.Location = New-Object System.Drawing.Point(10, 60)
$btnFixLan = New-Object System.Windows.Forms.Button; $btnFixLan.Text = "FIX ALL (SMB 1.0, RPC, SMB, PRINTNIGHTMARE)"; $btnFixLan.Location = New-Object System.Drawing.Point(20, 35); $btnFixLan.Size = New-Object System.Drawing.Size(700, 45); ThietKeNut $btnFixLan $MauNen.Do
$grLan.Controls.Add($btnFixLan)

$grDriver = New-Object System.Windows.Forms.GroupBox; $grDriver.Text = " 2. Sửa Lỗi Driver & Máy Scan (Epson/Brother) "; $grDriver.ForeColor = [System.Drawing.Color]::Gray; $grDriver.Size = New-Object System.Drawing.Size(740, 150); $grDriver.Location = New-Object System.Drawing.Point(10, 170)
$btnFix3e3 = New-Object System.Windows.Forms.Button; $btnFix3e3.Text = "FIX LỖI 3E3 (EPSON LQ-310)"; $btnFix3e3.Location = New-Object System.Drawing.Point(20, 35); $btnFix3e3.Size = New-Object System.Drawing.Size(340, 45); ThietKeNut $btnFix3e3 $MauNen.Cam
$btnFixBrother = New-Object System.Windows.Forms.Button; $btnFixBrother.Text = "FIX LỖI SCAN BROTHER (BỊ VĂNG)"; $btnFixBrother.Location = New-Object System.Drawing.Point(380, 35); $btnFixBrother.Size = New-Object System.Drawing.Size(340, 45); ThietKeNut $btnFixBrother $MauNen.Do
$grDriver.Controls.AddRange(@($btnFix3e3, $btnFixBrother))

$grUtil = New-Object System.Windows.Forms.GroupBox; $grUtil.Text = " 3. Tiện Ích Mở Rộng "; $grUtil.ForeColor = [System.Drawing.Color]::Gray; $grUtil.Size = New-Object System.Drawing.Size(740, 180); $grUtil.Location = New-Object System.Drawing.Point(10, 330)
$btnFwOn = New-Object System.Windows.Forms.Button; $btnFwOn.Text = "BẬT TƯỜNG LỬA"; $btnFwOn.Location = New-Object System.Drawing.Point(20, 35); $btnFwOn.Size = New-Object System.Drawing.Size(220, 40); ThietKeNut $btnFwOn $MauNen.XanhLa
$btnFwOff = New-Object System.Windows.Forms.Button; $btnFwOff.Text = "TẮT TƯỜNG LỬA"; $btnFwOff.Location = New-Object System.Drawing.Point(260, 35); $btnFwOff.Size = New-Object System.Drawing.Size(220, 40); ThietKeNut $btnFwOff $MauNen.Do
$btnResetSpool = New-Object System.Windows.Forms.Button; $btnResetSpool.Text = "RESET SPOOLER"; $btnResetSpool.Location = New-Object System.Drawing.Point(500, 35); $btnResetSpool.Size = New-Object System.Drawing.Size(220, 40); ThietKeNut $btnResetSpool $MauNen.NutMacDinh
$btnAddCred = New-Object System.Windows.Forms.Button; $btnAddCred.Text = "THÊM CREDENTIAL (GUEST)"; $btnAddCred.Location = New-Object System.Drawing.Point(20, 95); $btnAddCred.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnAddCred $MauNen.NutMacDinh
$btnSysRes = New-Object System.Windows.Forms.Button; $btnSysRes.Text = "MỞ SYSTEM RESTORE"; $btnSysRes.Location = New-Object System.Drawing.Point(380, 95); $btnSysRes.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnSysRes $MauNen.XanhDuong
$grUtil.Controls.AddRange(@($btnFwOn, $btnFwOff, $btnResetSpool, $btnAddCred, $btnSysRes))

$pnlPrint.Controls.AddRange(@($lblPrint, $grLan, $grDriver, $grUtil))
$khungChinh.Controls.Add($pnlPrint)

# LOGIC MÁY IN
$backupDir = "C:\Tool_Backups"
if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }

$btnFixLan.Add_Click({
    ChayTacVu "Đang Fix lỗi LAN 11b/0x709 & SMB1..." {
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> BẮT ĐẦU FIX LỖI LAN..."

        GhiLog "-> Đang bật SMB 1.0 (Cần thiết cho máy in cũ)..."
        try {
            Enable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart -ErrorAction Stop | Out-Null
            GhiLog "   + Đã bật SMB 1.0 thành công."
        } catch {
            GhiLog "   ! Lỗi bật SMB 1.0: $($_.Exception.Message)"
        }

        GhiLog "-> Đang cấu hình AllowInsecureGuestAuth..."
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" -Name "AllowInsecureGuestAuth" -Value 1 -Type DWord -Force | Out-Null
        
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
        
        Restart-Service -Name Spooler -Force
        GhiLog ">>> HOÀN TẤT! Vui lòng khởi động lại máy."
        [System.Windows.Forms.MessageBox]::Show("Đã Fix xong (SMB1 + LAN)! Hãy khởi động lại máy để áp dụng.", "Thông báo")
    }
})

$btnFix3e3.Add_Click({ ChayTacVu "Đang tìm Driver Epson..." { ChuyenTab $pnlLog $btnMenuLog; GhiLog ">>> BẮT ĐẦU FIX LỖI 3E3..."; $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"; $driverFound = $null; if (Test-Path "$basePath\Epson LQ-310 ESC") { $driverFound = "Epson LQ-310 ESC" } elseif (Test-Path "$basePath\Epson LQ-310 ESC/P2") { $driverFound = "Epson LQ-310 ESC/P2" }; if ($driverFound) { GhiLog "- Tìm thấy driver: $driverFound"; reg export "$basePath\$driverFound" "$backupDir\Epson_Backup.reg" /y 2>&1 | Out-Null; Set-ItemProperty -Path "$basePath\$driverFound" -Name "PrinterDriverAttributes" -Value 1; Restart-Service -Name Spooler -Force; GhiLog ">>> ĐÃ FIX XONG! Hãy in thử."; [System.Windows.Forms.MessageBox]::Show("Đã sửa lỗi Epson 3e3 thành công!", "Thông báo") } else { GhiLog ">>> LỖI: Không tìm thấy Driver Epson LQ-310."; [System.Windows.Forms.MessageBox]::Show("Không tìm thấy Driver Epson LQ-310.", "Lỗi") } } })
$btnFixBrother.Add_Click({ ChayTacVu "Fix lỗi Scan Brother" { ChuyenTab $pnlLog $btnMenuLog; GhiLog "--- FIX LỖI SCAN BROTHER (CRASH) ---"; $key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"; Set-ItemProperty -Path $key -Name "Personal" -Value "%USERPROFILE%\Documents" -Type ExpandString -ErrorAction SilentlyContinue; GhiLog "1. Đã Reset đường dẫn 'Personal' về mặc định."; GhiLog ">>> Hoàn tất."; [System.Windows.Forms.MessageBox]::Show("Đã Reset đường dẫn User Shell Folders.`nVui lòng Log out (Đăng xuất) hoặc khởi động lại máy để áp dụng.", "Thành công") } })
$btnResetSpool.Add_Click({ ChayTacVu "Reset Print Spooler" { ChuyenTab $pnlLog $btnMenuLog; GhiLog "--- RESET PRINT SPOOLER ---"; Stop-Service -Name Spooler -Force; GhiLog "1. Đã dừng dịch vụ Spooler"; $pathSpool = "$env:SystemRoot\System32\spool\PRINTERS\*"; Remove-Item $pathSpool -Force -Recurse -ErrorAction SilentlyContinue; GhiLog "2. Đã xóa cache lệnh in"; Start-Service -Name Spooler; GhiLog "3. Đã khởi động lại Spooler"; GhiLog ">>> Hoàn tất."; [System.Windows.Forms.MessageBox]::Show("Đã Reset Spooler & Xóa lệnh in kẹt!", "Thành công") } })
$btnFwOn.Add_Click({ ChayTacVu "Bật Tường Lửa" { netsh advfirewall set allprofiles state on; [System.Windows.Forms.MessageBox]::Show("Đã BẬT Tường lửa", "Info") } })
$btnFwOff.Add_Click({ ChayTacVu "Tắt Tường Lửa" { netsh advfirewall set allprofiles state off; [System.Windows.Forms.MessageBox]::Show("Đã TẮT Tường lửa", "Info") } })
$btnAddCred.Add_Click({ $ip = HienThiInputBox "Thêm Credential" "Nhập IP hoặc Tên Máy Chủ (VD: 192.168.1.10):"; if (-not [string]::IsNullOrWhiteSpace($ip)) { cmdkey /add:$ip /user:Guest; [System.Windows.Forms.MessageBox]::Show("Đã thêm User Guest cho IP: $ip", "Thành công") } })
$btnSysRes.Add_Click({ Start-Process rstrui.exe })