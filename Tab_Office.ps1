# ==============================================================================
# [ TAB 2 ] CÔNG CỤ OFFICE
# ==============================================================================
$pnlOff = New-Object System.Windows.Forms.Panel; $pnlOff.Dock = "Fill"; $pnlOff.Visible = $false
$lblTieuDeOff = New-Object System.Windows.Forms.Label; $lblTieuDeOff.Text = "QUẢN LÝ BẢN QUYỀN OFFICE"; $lblTieuDeOff.Font = $PhongChu.TieuDe; $lblTieuDeOff.ForeColor = $MauNen.Do; $lblTieuDeOff.Size = New-Object System.Drawing.Size(600, 40); $lblTieuDeOff.Location = New-Object System.Drawing.Point(10, 10)

$grOff1 = New-Object System.Windows.Forms.GroupBox; $grOff1.Text = " 1. Chẩn đoán & Dùng thử "; $grOff1.ForeColor = [System.Drawing.Color]::Gray; $grOff1.Size = New-Object System.Drawing.Size(740, 140); $grOff1.Location = New-Object System.Drawing.Point(10, 50)
$btnCheckStatus = New-Object System.Windows.Forms.Button; $btnCheckStatus.Text = "XEM CHI TIẾT LICENSE"; $btnCheckStatus.Location = New-Object System.Drawing.Point(20, 30); $btnCheckStatus.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnCheckStatus $MauNen.NutMacDinh
$btnCheckEdition = New-Object System.Windows.Forms.Button; $btnCheckEdition.Text = "KIỂM TRA PHIÊN BẢN (RETAIL/VL)"; $btnCheckEdition.Location = New-Object System.Drawing.Point(380, 30); $btnCheckEdition.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnCheckEdition $MauNen.XanhDuong
$btnResetTrial = New-Object System.Windows.Forms.Button; $btnResetTrial.Text = "RESET DÙNG THỬ (REARM 30 NGÀY)"; $btnResetTrial.Location = New-Object System.Drawing.Point(20, 80); $btnResetTrial.Size = New-Object System.Drawing.Size(700, 40); ThietKeNut $btnResetTrial $MauNen.Cam
$grOff1.Controls.AddRange(@($btnCheckStatus, $btnCheckEdition, $btnResetTrial))

$grOffPhone = New-Object System.Windows.Forms.GroupBox; $grOffPhone.Text = " 2. Kích hoạt qua Điện thoại (IID/CID) "; $grOffPhone.ForeColor = [System.Drawing.Color]::Gray; $grOffPhone.Size = New-Object System.Drawing.Size(740, 170); $grOffPhone.Location = New-Object System.Drawing.Point(10, 200)
$txtOffIID = New-Object System.Windows.Forms.TextBox; $txtOffIID.Location = New-Object System.Drawing.Point(20, 35); $txtOffIID.Size = New-Object System.Drawing.Size(550, 25); $txtOffIID.BackColor=$MauNen.O_Nhap; $txtOffIID.ForeColor=[System.Drawing.Color]::Yellow; $txtOffIID.ReadOnly=$true
$btnLayOffIID = New-Object System.Windows.Forms.Button; $btnLayOffIID.Text = "LẤY IID"; $btnLayOffIID.Location = New-Object System.Drawing.Point(580, 33); $btnLayOffIID.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnLayOffIID $MauNen.NutMacDinh
$btnWebCIDOff = New-Object System.Windows.Forms.Button; $btnWebCIDOff.Text = "MỞ WEB LẤY CID (VISUAL SUPPORT)"; $btnWebCIDOff.Location = New-Object System.Drawing.Point(20, 70); $btnWebCIDOff.Size = New-Object System.Drawing.Size(700, 35); ThietKeNut $btnWebCIDOff $MauNen.XanhDuong
$txtOffCID = New-Object System.Windows.Forms.TextBox; $txtOffCID.Location = New-Object System.Drawing.Point(20, 115); $txtOffCID.Size = New-Object System.Drawing.Size(550, 25); $txtOffCID.BackColor=$MauNen.O_Nhap; $txtOffCID.ForeColor=$MauNen.Chu
$btnNapOffCID = New-Object System.Windows.Forms.Button; $btnNapOffCID.Text = "NẠP CID"; $btnNapOffCID.Location = New-Object System.Drawing.Point(580, 113); $btnNapOffCID.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnNapOffCID $MauNen.Do
$grOffPhone.Controls.AddRange(@($txtOffIID, $btnLayOffIID, $btnWebCIDOff, $txtOffCID, $btnNapOffCID))

# --- ĐÃ NỚI RỘNG GROUP NÀY ĐỂ THÊM NÚT GỠ CRACK ---
$grGoKeyOff = New-Object System.Windows.Forms.GroupBox; $grGoKeyOff.Text = " 3. Quản lý & Gỡ License "; $grGoKeyOff.ForeColor = [System.Drawing.Color]::Gray; $grGoKeyOff.Size = New-Object System.Drawing.Size(740, 130); $grGoKeyOff.Location = New-Object System.Drawing.Point(10, 380)
$txtKeyGo = New-Object System.Windows.Forms.TextBox; $txtKeyGo.Location = New-Object System.Drawing.Point(20, 35); $txtKeyGo.Size = New-Object System.Drawing.Size(400, 25); $txtKeyGo.BackColor=$MauNen.O_Nhap; $txtKeyGo.ForeColor=$MauNen.Chu
$btnGoKeyOff = New-Object System.Windows.Forms.Button; $btnGoKeyOff.Text = "GỠ KEY (5 SỐ CUỐI)"; $btnGoKeyOff.Location = New-Object System.Drawing.Point(440, 33); $btnGoKeyOff.Size = New-Object System.Drawing.Size(280, 30); ThietKeNut $btnGoKeyOff $MauNen.Do
$btnGoCrack = New-Object System.Windows.Forms.Button; $btnGoCrack.Text = "🧹 GỠ SẠCH BẢN QUYỀN CRACK & KMS (RESET ZIN)"; $btnGoCrack.Location = New-Object System.Drawing.Point(20, 75); $btnGoCrack.Size = New-Object System.Drawing.Size(700, 35); ThietKeNut $btnGoCrack $MauNen.Cam
$grGoKeyOff.Controls.AddRange(@($txtKeyGo, $btnGoKeyOff, $btnGoCrack))

# --- ĐẨY XUỐNG DƯỚI CHO KHỎI ĐÈ LÊN NHAU ---
$grOutlook = New-Object System.Windows.Forms.GroupBox; $grOutlook.Text = " 4. Mở Rộng Dung Lượng Outlook (PST/OST) "; $grOutlook.ForeColor = [System.Drawing.Color]::Gray; $grOutlook.Size = New-Object System.Drawing.Size(740, 80); $grOutlook.Location = New-Object System.Drawing.Point(10, 520)
$btnOutlookExp = New-Object System.Windows.Forms.Button; $btnOutlookExp.Text = "TĂNG GIỚI HẠN DUNG LƯỢNG (TÙY CHỈNH GB)"; $btnOutlookExp.Location = New-Object System.Drawing.Point(20, 30); $btnOutlookExp.Size = New-Object System.Drawing.Size(700, 35); ThietKeNut $btnOutlookExp $MauNen.XanhDuong
$grOutlook.Controls.Add($btnOutlookExp)

$btnOhook = New-Object System.Windows.Forms.Button; $btnOhook.Text = "KÍCH HOẠT OFFICE VĨNH VIỄN (Ohook) 🔒"; $btnOhook.Location = New-Object System.Drawing.Point(10, 610); $btnOhook.Size = New-Object System.Drawing.Size(740, 60); ThietKeNut $btnOhook $MauNen.Do

$pnlOff.Controls.AddRange(@($lblTieuDeOff, $grOff1, $grOffPhone, $grGoKeyOff, $grOutlook, $btnOhook))
$khungChinh.Controls.Add($pnlOff)

# ==============================================================================
# LOGIC OFFICE
# ==============================================================================
function Tim-OSPP { return (Get-ChildItem "C:\Program Files\Microsoft Office", "C:\Program Files (x86)\Microsoft Office" -Filter OSPP.VBS -Recurse -ErrorAction SilentlyContinue | Select -First 1).FullName }

$btnCheckStatus.Add_Click({ ChayTacVu "Đang đọc License Office" { ChuyenTab $pnlLog $btnMenuLog; $v=Tim-OSPP; if($v){ cscript //nologo "$v" /dstatus | Out-String | ForEach-Object { GhiLog $_.Trim() } } else { GhiLog "Lỗi: Không tìm thấy file OSPP.VBS." } } })

$btnCheckEdition.Add_Click({ ChayTacVu "Đang phân tích phiên bản" { ChuyenTab $pnlLog $btnMenuLog; $v=Tim-OSPP; if($v){ GhiLog "--- ĐANG PHÂN TÍCH ---"; $ketQua = cscript //nologo "$v" /dstatus | Out-String; if ($ketQua -match "RETAIL edition") { GhiLog "=> Bản Bán Lẻ (RETAIL). Có thể lấy IID." } elseif ($ketQua -match "VOLUME edition") { GhiLog "=> Bản Doanh Nghiệp (VOLUME/VL). Có thể lấy IID." } elseif ($ketQua -match "O365" -or $ketQua -match "Subscription") { GhiLog "=> Bản Thuê Bao (Office 365). KHÔNG lấy được IID." } else { GhiLog "=> Không xác định." } } } })

$btnLayOffIID.Add_Click({ ChayTacVu "Đang lấy IID" { ChuyenTab $pnlLog $btnMenuLog; $v=Tim-OSPP; if($v){ GhiLog "Đang quét IID..."; $dongIID = cscript //nologo "$v" /dti | Select-String -Pattern "\d"; if ($dongIID) { $chuoiIID = (($dongIID -join "").Replace(" ","").Replace("-","")).Trim(); if ($chuoiIID -match "^\d+$") { $txtOffIID.Text = $chuoiIID; GhiLog "Lấy IID thành công." } else { GhiLog "Lỗi: IID không đúng định dạng." } } else { GhiLog "Lỗi: Không có IID. Kiểm tra lại phiên bản." } } } })

$btnNapOffCID.Add_Click({ ChayTacVu "Đang nạp CID" { $c=$txtOffCID.Text.Trim(); if($c.Length -lt 10){return}; ChuyenTab $pnlLog $btnMenuLog; $v=Tim-OSPP; if($v){ GhiLog "Đang nạp CID..."; cscript //nologo "$v" /atp:$c | Out-String | ForEach-Object { GhiLog $_.Trim() }; cscript //nologo "$v" /act | Out-String | ForEach-Object { GhiLog $_.Trim() } } } })

$btnResetTrial.Add_Click({ ChayTacVu "Đang Reset Trial" { ChuyenTab $pnlLog $btnMenuLog; $v=Tim-OSPP; if($v){ cscript //nologo "$v" /rearm | Out-String | ForEach-Object { GhiLog $_.Trim() } } } })

$btnGoKeyOff.Add_Click({ ChayTacVu "Đang gỡ Key" { $k=$txtKeyGo.Text.Trim(); if($k.Length -ne 5){return}; ChuyenTab $pnlLog $btnMenuLog; $v=Tim-OSPP; if($v){ cscript //nologo "$v" /unpkey:$k | Out-String | ForEach-Object { GhiLog $_.Trim() } } } })

# --- LOGIC NÚT XÓA SẠCH CRACK (MỚI THÊM) ---
$btnGoCrack.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Bạn có chắc chắn muốn xóa sạch bản quyền Office hiện tại (KMS, Crack) để nhập Key mới không?", "Xác nhận", "YesNo", "Warning") -eq "Yes") {
        ChayTacVu "Đang gỡ sạch Crack Office..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BẮT ĐẦU DỌN DẸP BẢN QUYỀN CRACK (KMS)..."
            
            $v = Tim-OSPP
            if ($v) {
                GhiLog "1. Đang xóa máy chủ KMS ảo (KMS Host)..."
                cscript //nologo "$v" /remhst | Out-Null
                cscript //nologo "$v" /ckms-domain | Out-Null

                GhiLog "2. Đang xóa các tác vụ Crack chạy ngầm..."
                Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -match "AutoKMS|AutoPico|KMS|OfficeSoftwareProtection" } | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue

                GhiLog "3. Đang quét và gỡ toàn bộ Product Key hiện tại..."
                $status = cscript //nologo "$v" /dstatus | Out-String
                # Tìm tất cả các đoạn có chứa "Last 5 characters of installed product key"
                $regex = "Last 5 characters of installed product key: (.{5})"
                $keys = [regex]::Matches($status, $regex) | ForEach-Object { $_.Groups[1].Value }
                
                if ($keys) {
                    foreach ($key in $keys) {
                        GhiLog " -> Đang gỡ Key có đuôi: $key"
                        cscript //nologo "$v" /unpkey:$key | Out-Null
                    }
                } else {
                    GhiLog " -> Không tìm thấy Key nào đang cài."
                }

                GhiLog "4. Đang làm sạch Registry KMS..."
                Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" -Name "KeyManagementServiceName" -ErrorAction SilentlyContinue
                
                GhiLog ">>> DỌN DẸP HOÀN TẤT! Office đã trở về trạng thái chưa kích hoạt (Zin)."
                [System.Windows.Forms.MessageBox]::Show("Đã gỡ sạch bản quyền Crack/KMS! Office đã sẵn sàng để kích hoạt xịn.", "Thành công")
            } else {
                GhiLog "Lỗi: Không tìm thấy file OSPP.VBS để thực thi."
            }
        }
    }
})

$btnWebCIDOff.Add_Click({ [System.Diagnostics.Process]::Start("https://visualsupport.microsoft.com/") })

$btnOhook.Add_Click({ if(XacNhanMatKhau){ ChayTacVu "Đang tải Script Ohook" { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $Link_HWID = "https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/Ohook.cmd"; $UrlChuan = $Link_HWID.Trim() -replace '[^\x20-\x7E]', ''; $FinalUrl = "$UrlChuan`?t=$((Get-Date).Ticks)"; try { $noiDung = (Invoke-RestMethod -Uri $FinalUrl -ErrorAction Stop).ToString(); if ([string]::IsNullOrWhiteSpace($noiDung)) { throw "File trống" }; ChayScriptOnline $noiDung "Ohook_Activation" } catch { [System.Windows.Forms.MessageBox]::Show("Lỗi tải Ohook (Check 404): $($_.Exception.Message)") } } } })

$btnOutlookExp.Add_Click({
    $userGB = HienThiInputBox "Mở Rộng Outlook" "Nhập dung lượng tối đa mong muốn (GB):`n(Ví dụ: 100)"
    if ([string]::IsNullOrWhiteSpace($userGB)) { return }
    
    $gbInt = 0
    if (-not [int64]::TryParse($userGB, [ref]$gbInt) -or $gbInt -le 0) {
        [System.Windows.Forms.MessageBox]::Show("Vui lòng nhập số lớn hơn 0.", "Lỗi nhập liệu")
        return
    }

    ChayTacVu "Đang cấu hình Outlook..." {
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> BẮT ĐẦU CẤU HÌNH DUNG LƯỢNG OUTLOOK: $gbInt GB"
        
        $MaxLimitBytes = $gbInt * 1GB
        $WarnLimitBytes = [int64]($MaxLimitBytes * 0.95)
        GhiLog "- Max Size: $MaxLimitBytes Bytes"
        GhiLog "- Warn Size: $WarnLimitBytes Bytes"

        GhiLog "- Đang đóng Outlook..."
        Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue

        $OutlookVersions = @("16.0", "15.0", "14.0", "12.0", "11.0")
        $found = $false
        
        foreach ($version in $OutlookVersions) {
            $RegPath = "HKCU:\Software\Microsoft\Office\$version\Outlook\PST"
            $OfficePath = "HKCU:\Software\Microsoft\Office\$version"
            
            if (Test-Path $OfficePath) {
                GhiLog "-> Phát hiện Office phiên bản $version"
                if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
                
                try {
                    Set-ItemProperty -Path $RegPath -Name "MaxLargeFileSize" -Type QWord -Value $MaxLimitBytes -Force
                    Set-ItemProperty -Path $RegPath -Name "WarnLargeFileSize" -Type QWord -Value $WarnLimitBytes -Force
                    GhiLog "   + Đã cập nhật Registry thành công!"
                    $found = $true
                } catch {
                    GhiLog "   ! Lỗi ghi Registry: $($_.Exception.Message)"
                }
            }
        }

        if ($found) {
            GhiLog ">>> Đang khởi động lại Outlook..."
            Start-Process outlook.exe -ErrorAction SilentlyContinue
            [System.Windows.Forms.MessageBox]::Show("Thành công! Đã đặt giới hạn Outlook lên $gbInt GB.", "Thông báo")
        } else {
            GhiLog "!!! Không tìm thấy phiên bản Outlook nào trong Registry."
            [System.Windows.Forms.MessageBox]::Show("Không tìm thấy Outlook trên máy này.", "Lỗi")
        }
    }
})