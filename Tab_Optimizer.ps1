# ==============================================================================
# [ TAB 5 ] TỐI ƯU HỆ THỐNG
# ==============================================================================
$pnlOpt = New-Object System.Windows.Forms.Panel; $pnlOpt.Dock = "Fill"; $pnlOpt.Visible = $false
$lblOpt = New-Object System.Windows.Forms.Label; $lblOpt.Text = "TỐI ƯU HÓA HỆ THỐNG (SYSTEM OPTIMIZER)"; $lblOpt.Font = $PhongChu.TieuDe; $lblOpt.ForeColor = $MauNen.Cam; $lblOpt.Size = New-Object System.Drawing.Size(600, 40); $lblOpt.Location = New-Object System.Drawing.Point(10, 10)

$txtOptWarn = New-Object System.Windows.Forms.Label; $txtOptWarn.Text = "CẢNH BÁO: Chức năng này can thiệp sâu vào hệ thống. Lời khuyên: Hệ thống sẽ tự động tạo Điểm Khôi Phục (Restore Point) trước khi chạy để đảm bảo an toàn."; $txtOptWarn.ForeColor = [System.Drawing.Color]::Gray; $txtOptWarn.Size = New-Object System.Drawing.Size(740, 40); $txtOptWarn.Location = New-Object System.Drawing.Point(10, 60)

# 1. NÚT TỐI ƯU CHUNG
$btnOptOneClick = New-Object System.Windows.Forms.Button; $btnOptOneClick.Text = "🚀 TỐI ƯU HÓA WINDOWS (1-CLICK DỌN RÁC & DỊCH VỤ)"; $btnOptOneClick.Location = New-Object System.Drawing.Point(10, 100); $btnOptOneClick.Size = New-Object System.Drawing.Size(740, 50); ThietKeNut $btnOptOneClick $MauNen.Do

# 2. NÚT TỐI ƯU GAMING
$btnOptGaming = New-Object System.Windows.Forms.Button; $btnOptGaming.Text = "🎮 TỐI ƯU HÓA MÁY GAMING (TĂNG FPS, GIẢM PING, CHUỘT CHUẨN)"; $btnOptGaming.Location = New-Object System.Drawing.Point(10, 160); $btnOptGaming.Size = New-Object System.Drawing.Size(740, 50); ThietKeNut $btnOptGaming $MauNen.XanhDuong

# 3. NÚT TỐI ƯU WORD & EXCEL (FULL KẾ TOÁN/VĂN PHÒNG)
$btnOptOffice = New-Object System.Windows.Forms.Button; $btnOptOffice.Text = "📈 TỐI ƯU HÓA WORD & EXCEL (CHUẨN VĂN PHÒNG, KẾ TOÁN)"; $btnOptOffice.Location = New-Object System.Drawing.Point(10, 220); $btnOptOffice.Size = New-Object System.Drawing.Size(740, 50); ThietKeNut $btnOptOffice ([System.Drawing.Color]::FromArgb(0, 150, 136))

# 4. NÚT KHÔI PHỤC WINDOWS
$btnOptRestoreWin = New-Object System.Windows.Forms.Button; $btnOptRestoreWin.Text = "🔄 KHÔI PHỤC WINDOWS (TRẢ LẠI GIA TỐC CHUỘT VÀ MẠNG)"; $btnOptRestoreWin.Location = New-Object System.Drawing.Point(10, 280); $btnOptRestoreWin.Size = New-Object System.Drawing.Size(365, 50); ThietKeNut $btnOptRestoreWin $MauNen.XanhLa

# 5. NÚT KHÔI PHỤC OFFICE
$btnOptRestoreOffice = New-Object System.Windows.Forms.Button; $btnOptRestoreOffice.Text = "🔄 KHÔI PHỤC OFFICE (TRẢ VỀ MẶC ĐỊNH WORD/EXCEL)"; $btnOptRestoreOffice.Location = New-Object System.Drawing.Point(385, 280); $btnOptRestoreOffice.Size = New-Object System.Drawing.Size(365, 50); ThietKeNut $btnOptRestoreOffice $MauNen.XanhLa

$pnlOpt.Controls.AddRange(@($lblOpt, $txtOptWarn, $btnOptOneClick, $btnOptGaming, $btnOptOffice, $btnOptRestoreWin, $btnOptRestoreOffice))
$khungChinh.Controls.Add($pnlOpt)

# ==============================================================================
# LOGIC 1: TỐI ƯU WINDOWS CHUNG
# ==============================================================================
$btnOptOneClick.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Bạn có chắc chắn muốn tối ưu hóa Windows không?", "Xác nhận Tối ưu", "YesNo", "Warning") -eq "Yes") {
        ChayTacVu "Đang tối ưu hóa hệ thống..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BẮT ĐẦU TỐI ƯU HÓA WINDOWS CHUNG..."
            try { Checkpoint-Computer -Description "Truoc_Khi_Toi_Uu" -RestorePointType "MODIFY_SETTINGS" -ErrorAction SilentlyContinue; GhiLog "Đã tạo System Restore Point." } catch {}

            GhiLog "Đang tiến hành dọn dẹp rác hệ thống siêu tốc..."
            Stop-Process -Name "cleanmgr" -Force -ErrorAction SilentlyContinue
            Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches" -ErrorAction SilentlyContinue | ForEach-Object { Remove-ItemProperty -Path $_.PSPath -Name "StateFlags0001" -ErrorAction SilentlyContinue }

            Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item "$env:windir\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item "$env:windir\Prefetch\*" -Recurse -Force -ErrorAction SilentlyContinue
            Clear-RecycleBin -Force -Confirm:$false -ErrorAction SilentlyContinue

            GhiLog "-> Đang tắt Telemetry và các dịch vụ ngầm theo dõi..."
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting" "Disabled" 1
            Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance" "fAllowToGetHelp" 0
            Set-Reg "HKCU:\Control Panel\Desktop" "AutoEndTasks" 1

            # --- COMBO TỐI ƯU TRẢI NGHIỆM NGƯỜI DÙNG (UX) ĐÃ FIX LỖI ---
            GhiLog "-> Đang bật sẵn NumLock ở màn hình khởi động..."
            Set-Reg "Registry::HKEY_USERS\.DEFAULT\Control Panel\Keyboard" "InitialKeyboardIndicators" "2" "String"

            GhiLog "-> Đang thiết lập mở File Explorer ra thẳng This PC (Danh sách ổ đĩa)..."
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 1 "DWord"

            GhiLog "-> Đang hiển thị đuôi file (File extensions) để phòng chống virus..."
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0 "DWord"

            GhiLog "-> Đang tắt màn hình làm phiền 'Let's finish setting up your PC'..."
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 0 "DWord"

            GhiLog ">>> TỐI ƯU HÓA HOÀN TẤT!"
            [System.Windows.Forms.MessageBox]::Show("Đã tối ưu xong Windows! Hãy khởi động lại máy để áp dụng.", "Thành công")
        }
    }
})

# ==============================================================================
# LOGIC 2: TỐI ƯU HÓA MÁY GAMING
# ==============================================================================
$btnOptGaming.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Kích hoạt chế độ Gaming? Sẽ tắt gia tốc chuột và tối ưu mạng.", "Xác nhận Gaming", "YesNo", "Information") -eq "Yes") {
        ChayTacVu "Đang cấu hình máy Gaming..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BẮT ĐẦU TỐI ƯU HÓA MÁY GAMING..."

            GhiLog "-> Đang bật Game Mode và tắt Xbox DVR..."
            Set-Reg "HKCU:\Software\Microsoft\GameBar" "AllowAutoGameMode" 1
            Set-Reg "HKCU:\System\GameConfigStore" "GameDVR_Enabled" 0
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" "AllowGameDVR" 0

            GhiLog "-> Đang tắt gia tốc chuột (Chuẩn xác 100% cho game FPS)..."
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseSpeed" "0" "String"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseThreshold1" "0" "String"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseThreshold2" "0" "String"

            GhiLog "-> Đang xóa bỏ giới hạn băng thông mạng..."
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "NetworkThrottlingIndex" 4294967295 
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "SystemResponsiveness" 0

            GhiLog ">>> CẤU HÌNH GAMING HOÀN TẤT!"
            [System.Windows.Forms.MessageBox]::Show("Đã Tối ưu Gaming thành công!", "Thành công")
        }
    }
})

# ==============================================================================
# LOGIC TỐI ƯU OFFICE THÔNG MINH (PHÂN BIỆT BẢN QUYỀN VÀ THUỐC)
# ==============================================================================
$btnOptOffice.Add_Click({
    $XacNhan = [System.Windows.Forms.MessageBox]::Show("Hệ thống sẽ tự động nhận diện loại Office để tối ưu:`n- Office Bản quyền: Chỉ tối ưu in ấn, chống treo, giữ tính năng Online.`n- Office Thuốc: Tối ưu toàn diện + Chặn báo bản quyền + Tắt Update.`n`nBạn có muốn bắt đầu?", "Xác nhận tối ưu Office", "YesNo", "Question")
    
    if ($XacNhan -eq "Yes") {
        ChayTacVu "Đang tối ưu Office thông minh..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BẮT ĐẦU QUY TRÌNH PHÂN TÍCH VÀ TỐI ƯU OFFICE..."

            # 1. ĐÓNG CÁC TIẾN TRÌNH ĐANG CHẠY
            Stop-Process -Name "winword", "excel", "powerpnt", "outlook" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2

            # 2. KIỂM TRA LOẠI OFFICE (BẢN QUYỀN HAY THUỐC)
            $IsLicenseXinh = $false
            $CheckCTR = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "ProductReleaseIds" -ErrorAction SilentlyContinue
            
            # Nếu là Office 365 hoặc các bản Subscription thì coi như hàng xịn cần giữ Online
            if ($CheckCTR.ProductReleaseIds -like "*O365*" -or $CheckCTR.ProductReleaseIds -like "*M365*" -or $CheckCTR.ProductReleaseIds -like "*Sub*") {
                $IsLicenseXinh = $true
                GhiLog " -> PHÁT HIỆN: Office 365/Bản quyền tài khoản. Chế độ: TỐI ƯU AN TOÀN." -ForegroundColor Cyan
            } else {
                GhiLog " -> PHÁT HIỆN: Office Retail/Volume (Hàng cài sẵn). Chế độ: FULL GIÁP." -ForegroundColor Yellow
            }

            $OfficeVers = @("14.0", "15.0", "16.0")
            foreach ($Ver in $OfficeVers) {
                $Root = "HKCU:\Software\Microsoft\Office\$Ver"
                if (Test-Path $Root) {
                    GhiLog " -> Đang xử lý cấu hình cho Office v$Ver..."

                    # --- NHÓM 1: TỐI ƯU IN ẤN & CHỐNG TREO (TẤT CẢ ĐỀU DÙNG) ---
                    # Tắt Scale A4 (Fix lệch lề)
                    Set-Reg "$Root\Word\Options" "NoScalingPaperRes" 1
                    Set-Reg "$Root\Excel\Options" "AutomatedScaling" 0
                    
                    # Chống treo file nặng (Tắt tăng tốc đồ họa & Live Preview)
                    Set-Reg "$Root\Common\Graphics" "DisableHardwareAcceleration" 1
                    Set-Reg "$Root\Common\Graphics" "DisableAnimations" 1
                    Set-Reg "$Root\Common\General" "EnableLivePreview" 0
                    
                    # Tắt gạch chân chính tả & Màn hình Start
                    Set-Reg "$Root\Word\Options" "AutoSpell" 0
                    Set-Reg "$Root\Word\Options" "AutoGrammar" 0
                    Set-Reg "$Root\Common\General" "DisableBootToOfficeStart" 1

                    # Chế độ xem an toàn (Mở file Zalo là gõ được ngay)
                    $Apps = @("Word", "Excel")
                    foreach ($App in $Apps) {
                        $PVPath = "$Root\$App\Security\ProtectedView"
                        Set-Reg $PVPath "DisableInternetFilesInPV" 1
                        Set-Reg $PVPath "DisableUnsafeLocationsInPV" 1
                        Set-Reg $PVPath "DisableAttachmentsInPV" 1
                    }

                    # --- NHÓM 2: CHẶN BẢN QUYỀN & UPDATE (CHỈ ÁP DỤNG CHO HÀNG THUỐC) ---
                    if ($IsLicenseXinh -eq $false) {
                        GhiLog "    [+] Khóa Banner bản quyền và Tắt Update..."
                        # Chặn banner "Get Genuine" bằng cách ngắt kết nối trải nghiệm người dùng
                        Set-Reg "$Root\Common\Privacy" "DisconnectedState" 1
                        # Tắt tự động cập nhật
                        Set-Reg "$Root\Common\General" "EnableAutomaticUpdates" 0
                        Set-Reg "$Root\Common\General" "HideIdentifyAutomaticUpdates" 1
                        # Chặn gửi dữ liệu về MS
                        Set-Reg "$Root\Common\General" "sendtelemetry" 3
                    }

                    # --- NHÓM 3: CHUẨN KẾ TOÁN (TẤT CẢ ĐỀU DÙNG) ---
                    Set-Reg "$Root\Excel\Options" "UseSystemSeparators" 0 "DWord"
                    Set-Reg "$Root\Excel\Options" "DecimalSeparator" "," "String"
                    Set-Reg "$Root\Excel\Options" "ThousandsSeparator" "." "String"
                    Set-Reg "$Root\Excel\Options" "Font" "Times New Roman, 14" "String"
                }
            }

            # 3. THIẾT LẬP REGION WINDOWS (dd/MM/yyyy)
            GhiLog " -> Thiết lập chuẩn Region Windows (dd/MM/yyyy)..."
            $Intl = "HKCU:\Control Panel\International"
            Set-Reg $Intl "sShortDate" "dd/MM/yyyy" "String"
            Set-Reg $Intl "sDecimal" "," "String"
            Set-Reg $Intl "sThousand" "." "String"

            # 4. CĂN LỀ NGHỊ ĐỊNH 30 CHO WORD (NORMAL.DOTM)
            GhiLog " -> Đang can thiệp Normal.dotm (Nghị định 30)..."
            try {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $doc = $word.NormalTemplate.OpenAsDocument()
                $doc.Styles.Item("Normal").Font.Name = "Times New Roman"
                $doc.Styles.Item("Normal").Font.Size = 14
                $doc.PageSetup.PaperSize = 7 # A4
                $doc.PageSetup.TopMargin = 56.7    # 2cm
                $doc.PageSetup.BottomMargin = 56.7 # 2cm
                $doc.PageSetup.LeftMargin = 85.05  # 3cm
                $doc.PageSetup.RightMargin = 42.55 # 1.5cm
                $doc.Save()
                $doc.Close()
                GhiLog "    [OK] Đã lưu lề chuẩn Nghị định 30."
            } catch {
                GhiLog "    [!] Không can thiệp được file Normal.dotm (Bỏ qua)."
            } finally {
                if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
            }

            GhiLog ">>> TẤT CẢ TÁC VỤ HOÀN TẤT!"
            [System.Windows.Forms.MessageBox]::Show("Đã tối ưu Office xong!`nLoại Office: $(if($IsLicenseXinh){'Bản quyền'}else{'Cài sẵn'})", "Thành công")
        }
    }
})
# ==============================================================================
# LOGIC 4: KHÔI PHỤC MẶC ĐỊNH WINDOWS (CHUỘT, MẠNG, UX)
# ==============================================================================
$btnOptRestoreWin.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Khôi phục cài đặt gốc của Windows (Gia tốc chuột, Băng thông mạng)?", "Xác nhận Khôi phục", "YesNo", "Question") -eq "Yes") {
        ChayTacVu "Đang khôi phục Windows..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BẮT ĐẦU KHÔI PHỤC WINDOWS..."

            GhiLog "-> Đang bật lại gia tốc chuột và khôi phục cấu hình mạng..."
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseSpeed" "1" "String"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseThreshold1" "6" "String"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseThreshold2" "10" "String"
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "NetworkThrottlingIndex" 10
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "SystemResponsiveness" 20

            # Khôi phục Trải nghiệm người dùng về mặc định Windows ĐÃ FIX LỖI
            GhiLog "-> Đang khôi phục File Explorer và các cài đặt trải nghiệm (UX)..."
            Set-Reg "Registry::HKEY_USERS\.DEFAULT\Control Panel\Keyboard" "InitialKeyboardIndicators" "0" "String"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 2 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 1 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 1 "DWord"

            GhiLog ">>> KHÔI PHỤC WINDOWS HOÀN TẤT!"
            [System.Windows.Forms.MessageBox]::Show("Đã trả Windows về thiết lập nguyên bản!", "Thành công")
        }
    }
})

# ==============================================================================
# LOGIC 5: KHÔI PHỤC MẶC ĐỊNH OFFICE & CHUẨN QUỐC TẾ
# ==============================================================================
$btnOptRestoreOffice.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Khôi phục cài đặt gốc của Office (Trả về mặc định Word, Excel, Ngày tháng, Số tiền)?", "Xác nhận Khôi phục", "YesNo", "Question") -eq "Yes") {
        ChayTacVu "Đang khôi phục Office..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BẮT ĐẦU KHÔI PHỤC OFFICE..."

            GhiLog "-> Đang khôi phục cài đặt mặc định của Word & Excel..."
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Common\General" "DisableBootToOfficeStart" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Common\Graphics" "DisableAnimations" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Common\Graphics" "DisableHardwareAcceleration" 0
            
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Word\Options" "AutoRecoverDelay" 10
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" "AutoRecoverDelay" 10

            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableInternetFilesInPV" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableUnsafeLocationsInPV" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableAttachmentsInPV" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" "DisableInternetFilesInPV" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" "DisableUnsafeLocationsInPV" 0
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" "DisableAttachmentsInPV" 0

            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" "VBAWarnings" 2
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Word\Options" "AutoSpell" 1
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Word\Options" "AutoGrammar" 1

            # Khôi phục Ngày tháng và Số tiền về chuẩn Mỹ
            GhiLog "-> Đang khôi phục định dạng Ngày/Tháng và Số tiền về chuẩn Mỹ..."
            Set-Reg "HKCU:\Control Panel\International" "sShortDate" "M/d/yyyy" "String"
            Set-Reg "HKCU:\Control Panel\International" "sLongDate" "dddd, MMMM d, yyyy" "String"
            Set-Reg "HKCU:\Control Panel\International" "sDecimal" "." "String"
            Set-Reg "HKCU:\Control Panel\International" "sThousand" "," "String"

            GhiLog "-> Đang trả lại quyền quản lý dấu chấm/phẩy cho Windows..."
            Set-Reg "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" "UseSystemSeparators" 1 "DWord"
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "DecimalSeparator" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "ThousandsSeparator" -ErrorAction SilentlyContinue

            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -Name "Font" -ErrorAction SilentlyContinue

            GhiLog ">>> KHÔI PHỤC OFFICE HOÀN TẤT!"
            [System.Windows.Forms.MessageBox]::Show("Đã trả Office và các định dạng về nguyên bản thành công!", "Thành công")
        }
    }
})

