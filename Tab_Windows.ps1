# ==============================================================================
# [ TAB 1 ] CÔNG CỤ WINDOWS
# ==============================================================================
$pnlWin = New-Object System.Windows.Forms.Panel; $pnlWin.Dock = "Fill"; $pnlWin.Visible = $false
$lblTieuDeWin = New-Object System.Windows.Forms.Label; $lblTieuDeWin.Text = "QUẢN LÝ BẢN QUYỀN WINDOWS"; $lblTieuDeWin.Font = $PhongChu.TieuDe; $lblTieuDeWin.ForeColor = $MauNen.XanhDuong; $lblTieuDeWin.Size = New-Object System.Drawing.Size(600, 40); $lblTieuDeWin.Location = New-Object System.Drawing.Point(10, 10)

$grKeyWin = New-Object System.Windows.Forms.GroupBox; $grKeyWin.Text = " 1. Nhập Key và Kiểm tra "; $grKeyWin.ForeColor = [System.Drawing.Color]::Gray; $grKeyWin.Size = New-Object System.Drawing.Size(740, 140); $grKeyWin.Location = New-Object System.Drawing.Point(10, 60)
$txtOKeyWin = New-Object System.Windows.Forms.TextBox; $txtOKeyWin.Location = New-Object System.Drawing.Point(20, 35); $txtOKeyWin.Size = New-Object System.Drawing.Size(400, 25); $txtOKeyWin.BackColor = $MauNen.O_Nhap; $txtOKeyWin.ForeColor = $MauNen.Chu
$btnNapKey = New-Object System.Windows.Forms.Button; $btnNapKey.Text = "NẠP KEY"; $btnNapKey.Location = New-Object System.Drawing.Point(430, 33); $btnNapKey.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnNapKey $MauNen.NutMacDinh
$btnKichHoatNgay = New-Object System.Windows.Forms.Button; $btnKichHoatNgay.Text = "OEM TỰ ĐỘNG"; $btnKichHoatNgay.Location = New-Object System.Drawing.Point(580, 33); $btnKichHoatNgay.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnKichHoatNgay $MauNen.XanhLa
$btnXemKeyHeThong = New-Object System.Windows.Forms.Button; $btnXemKeyHeThong.Text = "XEM KEY ĐANG CÀI & BIOS"; $btnXemKeyHeThong.Location = New-Object System.Drawing.Point(20, 80); $btnXemKeyHeThong.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnXemKeyHeThong $MauNen.XanhDuong
$btnKiemTraHan = New-Object System.Windows.Forms.Button; $btnKiemTraHan.Text = "TRẠNG THÁI BẢN QUYỀN HIỆN TẠI"; $btnKiemTraHan.Location = New-Object System.Drawing.Point(380, 80); $btnKiemTraHan.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnKiemTraHan $MauNen.NutMacDinh
$grKeyWin.Controls.AddRange(@($txtOKeyWin, $btnNapKey, $btnKichHoatNgay, $btnXemKeyHeThong, $btnKiemTraHan))

$grPhoneWin = New-Object System.Windows.Forms.GroupBox; $grPhoneWin.Text = " 2. Kích hoạt qua Điện thoại (IID/CID) "; $grPhoneWin.ForeColor = [System.Drawing.Color]::Gray; $grPhoneWin.Size = New-Object System.Drawing.Size(740, 170); $grPhoneWin.Location = New-Object System.Drawing.Point(10, 210)
$txtHienIID = New-Object System.Windows.Forms.TextBox; $txtHienIID.Location = New-Object System.Drawing.Point(20, 35); $txtHienIID.Size = New-Object System.Drawing.Size(550, 25); $txtHienIID.BackColor=$MauNen.O_Nhap; $txtHienIID.ForeColor=[System.Drawing.Color]::Yellow; $txtHienIID.ReadOnly=$true
$btnLayIID = New-Object System.Windows.Forms.Button; $btnLayIID.Text = "LẤY IID"; $btnLayIID.Location = New-Object System.Drawing.Point(580, 33); $btnLayIID.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnLayIID $MauNen.NutMacDinh
$btnMoWebCID = New-Object System.Windows.Forms.Button; $btnMoWebCID.Text = "MỞ TRANG LẤY CID (VISUAL SUPPORT)"; $btnMoWebCID.Location = New-Object System.Drawing.Point(20, 70); $btnMoWebCID.Size = New-Object System.Drawing.Size(700, 35); ThietKeNut $btnMoWebCID $MauNen.XanhDuong
$txtNhapCID = New-Object System.Windows.Forms.TextBox; $txtNhapCID.Location = New-Object System.Drawing.Point(20, 120); $txtNhapCID.Size = New-Object System.Drawing.Size(550, 25); $txtNhapCID.BackColor=$MauNen.O_Nhap; $txtNhapCID.ForeColor=$MauNen.Chu
$btnNapCID = New-Object System.Windows.Forms.Button; $btnNapCID.Text = "NẠP CID"; $btnNapCID.Location = New-Object System.Drawing.Point(580, 118); $btnNapCID.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnNapCID $MauNen.Do
$grPhoneWin.Controls.AddRange(@($txtHienIID, $btnLayIID, $btnMoWebCID, $txtNhapCID, $btnNapCID))

$grUpgradeWin = New-Object System.Windows.Forms.GroupBox; $grUpgradeWin.Text = " 3. Chuyển đổi/Nâng cấp phiên bản "; $grUpgradeWin.ForeColor = [System.Drawing.Color]::Gray; $grUpgradeWin.Size = New-Object System.Drawing.Size(740, 80); $grUpgradeWin.Location = New-Object System.Drawing.Point(10, 390)
$cmbPhienBan = New-Object System.Windows.Forms.ComboBox; $cmbPhienBan.Location = New-Object System.Drawing.Point(20, 30); $cmbPhienBan.Size = New-Object System.Drawing.Size(350, 30); $cmbPhienBan.BackColor=$MauNen.O_Nhap; $cmbPhienBan.ForeColor=$MauNen.Chu; [void]$cmbPhienBan.Items.AddRange(@("Windows 10/11 Professional", "Windows 10/11 Enterprise", "Windows 10/11 Education")); $cmbPhienBan.SelectedIndex=0
$btnNangCap = New-Object System.Windows.Forms.Button; $btnNangCap.Text = "NÂNG CẤP"; $btnNangCap.Location = New-Object System.Drawing.Point(400, 28); $btnNangCap.Size = New-Object System.Drawing.Size(320, 35); ThietKeNut $btnNangCap $MauNen.XanhLa
$grUpgradeWin.Controls.AddRange(@($cmbPhienBan, $btnNangCap))

$grUpdate = New-Object System.Windows.Forms.GroupBox; $grUpdate.Text = " 4. Quản lý Windows Update "; $grUpdate.ForeColor = [System.Drawing.Color]::Gray; $grUpdate.Size = New-Object System.Drawing.Size(740, 80); $grUpdate.Location = New-Object System.Drawing.Point(10, 480)
$btnUpOn = New-Object System.Windows.Forms.Button; $btnUpOn.Text = "BẬT UPDATE (MẶC ĐỊNH)"; $btnUpOn.Location = New-Object System.Drawing.Point(20, 30); $btnUpOn.Size = New-Object System.Drawing.Size(340, 35); ThietKeNut $btnUpOn $MauNen.XanhLa
$btnUpOff = New-Object System.Windows.Forms.Button; $btnUpOff.Text = "TẮT UPDATE (VĨNH VIỄN)"; $btnUpOff.Location = New-Object System.Drawing.Point(380, 30); $btnUpOff.Size = New-Object System.Drawing.Size(340, 35); ThietKeNut $btnUpOff $MauNen.Do
$grUpdate.Controls.AddRange(@($btnUpOn, $btnUpOff))

$btnHWID = New-Object System.Windows.Forms.Button; $btnHWID.Text = "KÍCH HOẠT BẢN QUYỀN SỐ VĨNH VIỄN (HWID) 🔒"; $btnHWID.Location = New-Object System.Drawing.Point(10, 570); $btnHWID.Size = New-Object System.Drawing.Size(740, 60); ThietKeNut $btnHWID $MauNen.XanhDuong
$pnlWin.Controls.AddRange(@($lblTieuDeWin, $grKeyWin, $grPhoneWin, $grUpgradeWin, $grUpdate, $btnHWID))
$khungChinh.Controls.Add($pnlWin)

# LOGIC WINDOWS
$btnNapKey.Add_Click({ ChayTacVu "Đang nạp Key" { $k = $txtOKeyWin.Text.Trim(); if ($k.Length -lt 5) { return }; ChuyenTab $pnlLog $btnMenuLog; GhiLog "Nạp Key: $k"; cscript //nologo $env:windir\system32\slmgr.vbs /ipk $k | Out-String | ForEach-Object { GhiLog $_.Trim() } } })
$btnKichHoatNgay.Add_Click({ ChayTacVu "Đang kích hoạt Online" { ChuyenTab $pnlLog $btnMenuLog; cscript //nologo $env:windir\system32\slmgr.vbs /ato | Out-String | ForEach-Object { GhiLog $_.Trim() } } })
$btnXemKeyHeThong.Add_Click({ ChayTacVu "Đang đọc Key" { $reg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DigitalProductId -ErrorAction SilentlyContinue; if($reg){$raw=$reg.DigitalProductId;$isWin8=($raw[66]-band 1)-ne 0;$i=24;$c="BCDFGHJKMPQRTVWXY2346789";$k="";if($isWin8){$raw[66]=($raw[66]-band 0xF7)};do{$cur=0;$x=14;do{$cur=$cur*256;$cur=$raw[$x+52]+$cur;$raw[$x+52]=[math]::Floor($cur/24);$cur=$cur%24;$x--}while($x-ge 0);$i--;$k=$c[$cur]+$k;$last=$cur}while($i-ge 0);if($isWin8){$keyp1=$k.Substring(1,$last);$keyp2=$k.Substring($last+1,$k.Length-($last+1));$k=$keyp1+"N"+$keyp2};$insKey="";for($j=0;$j-lt $k.Length;$j++){$insKey+=$k[$j];if(($j+1)%5-eq 0 -and ($j+1)-ne $k.Length){$insKey+="-"}}}else{$insKey="N/A"}; $bios=(Get-CimInstance -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService").OA3xOriginalProductKey; [System.Windows.Forms.MessageBox]::Show("Key đang cài: $insKey`nKey BIOS: $bios", "Thông tin") } })
$btnKiemTraHan.Add_Click({ ChayTacVu "Đang kiểm tra thời hạn" { $items=Get-CimInstance SoftwareLicensingProduct|Where-Object{$_.PartialProductKey};$m="";foreach($i in $items){$st=switch($i.LicenseStatus){1{"Đã kích hoạt"};0{"Chưa bản quyền"};Default{"Lỗi"}};$m+="$($i.Name)`nTT: $st`nKey: ...-$($i.PartialProductKey)`n----------------`n"};[System.Windows.Forms.MessageBox]::Show($m, "Chi tiết") } })
$btnLayIID.Add_Click({ ChayTacVu "Đang lấy IID" { $iid=(cscript //nologo $env:windir\system32\slmgr.vbs /dti).Trim(); $txtHienIID.Text=$iid } })
$btnMoWebCID.Add_Click({ [System.Diagnostics.Process]::Start("https://visualsupport.microsoft.com/") })
$btnNapCID.Add_Click({ ChayTacVu "Đang nạp CID" { $c = $txtNhapCID.Text.Trim(); if ($c.Length -lt 10) { return }; ChuyenTab $pnlLog $btnMenuLog; cscript //nologo $env:windir\system32\slmgr.vbs /atp $c | Out-String | ForEach-Object { GhiLog $_.Trim() }; cscript //nologo $env:windir\system32\slmgr.vbs /ato | Out-String | ForEach-Object { GhiLog $_.Trim() } } })
$btnNangCap.Add_Click({ ChayTacVu "Đang nâng cấp" { $ed=$cmbPhienBan.SelectedItem; $k=""; if($ed-match"Professional"){$k="VK7JG-NPHTM-C97JM-9MPGT-3V66T"}; if($ed-match"Enterprise"){$k="NPPR9-FWDCX-D2C8J-H2M7V-T6WDT"}; if($ed-match"Education"){$k="NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"}; ChuyenTab $pnlLog $btnMenuLog; cscript //nologo $env:windir\system32\slmgr.vbs /ipk $k | Out-String | ForEach-Object { GhiLog $_.Trim() } } })
$btnHWID.Add_Click({ if(XacNhanMatKhau){ ChayTacVu "Đang tải Script HWID" { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $Link_HWID = "https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/HWID.cmd"; $UrlChuan = $Link_HWID.Trim() -replace '[^\x20-\x7E]', ''; $FinalUrl = "$UrlChuan`?t=$((Get-Date).Ticks)"; try { $noiDung = (Invoke-RestMethod -Uri $FinalUrl -ErrorAction Stop).ToString(); if ([string]::IsNullOrWhiteSpace($noiDung)) { throw "File trống" }; ChayScriptOnline $noiDung "HWID_Activation" } catch { [System.Windows.Forms.MessageBox]::Show("Lỗi tải HWID: $($_.Exception.Message)") } } } })

$btnUpOn.Add_Click({
    ChayTacVu "Đang BẬT Windows Update" {
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> BẬT WINDOWS UPDATE..."
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
        reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v NoAutoUpdate /f 2>&1 | Out-Null
        Set-Service wuauserv -StartupType Manual -ErrorAction SilentlyContinue
        Start-Service wuauserv -ErrorAction SilentlyContinue
        GhiLog "-> Đã Bật dịch vụ Update & Xóa Policy chặn."
        [System.Windows.Forms.MessageBox]::Show("Đã BẬT Windows Update thành công!", "Thông báo")
    }
})

$btnUpOff.Add_Click({
    ChayTacVu "Đang TẮT Windows Update" {
        ChuyenTab $pnlLog $btnMenuLog
        GhiLog ">>> TẮT WINDOWS UPDATE..."
        Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
        Set-Service wuauserv -StartupType Disabled -ErrorAction SilentlyContinue
        if (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU")) { New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null }
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Value 1 -Type DWord -Force
        GhiLog "-> Đã Tắt dịch vụ Update & Thêm Policy chặn."
        [System.Windows.Forms.MessageBox]::Show("Đã TẮT Windows Update vĩnh viễn!", "Thông báo")
    }
})