# ==============================================================================
# [ TAB 1 ] CÔNG CỤ WINDOWS (DÙNG CHO CẤU TRÚC MULTI-TAB)
# ==============================================================================
$pnlWin = New-Object System.Windows.Forms.Panel; $pnlWin.Dock = "Fill"; $pnlWin.Visible = $false
$lblTieuDeWin = New-Object System.Windows.Forms.Label; $lblTieuDeWin.Text = "QUẢN LÝ BẢN QUYỀN WINDOWS"; $lblTieuDeWin.Font = $PhongChu.TieuDe; $lblTieuDeWin.ForeColor = $MauNen.XanhDuong; $lblTieuDeWin.Size = New-Object System.Drawing.Size(600, 40); $lblTieuDeWin.Location = New-Object System.Drawing.Point(10, 10)

# --- 1. KIỂM TRA & REACTIVE OEM ---
$grCheckWin = New-Object System.Windows.Forms.GroupBox; $grCheckWin.Text = " 1. Kiểm tra & Kích hoạt OEM "; $grCheckWin.ForeColor = [System.Drawing.Color]::Gray; $grCheckWin.Size = New-Object System.Drawing.Size(740, 100); $grCheckWin.Location = New-Object System.Drawing.Point(10, 60)
$btnShowKeys = New-Object System.Windows.Forms.Button; $btnShowKeys.Text = "HIỆN KEY: INSTALLED & OEM (BIOS)"; $btnShowKeys.Location = New-Object System.Drawing.Point(20, 35); $btnShowKeys.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnShowKeys $MauNen.XanhDuong
$btnReactiveOEM = New-Object System.Windows.Forms.Button; $btnReactiveOEM.Text = "REACTIVE LẠI THEO KEY OEM"; $btnReactiveOEM.Location = New-Object System.Drawing.Point(380, 35); $btnReactiveOEM.Size = New-Object System.Drawing.Size(340, 40); ThietKeNut $btnReactiveOEM $MauNen.XanhLa
$grCheckWin.Controls.AddRange(@($btnShowKeys, $btnReactiveOEM))

# --- 2. KÍCH HOẠT QUA ĐIỆN THOẠI (IID/CID) ---
$grPhoneWin = New-Object System.Windows.Forms.GroupBox; $grPhoneWin.Text = " 2. Kích hoạt qua Điện thoại (IID/CID) "; $grPhoneWin.ForeColor = [System.Drawing.Color]::Gray; $grPhoneWin.Size = New-Object System.Drawing.Size(740, 180); $grPhoneWin.Location = New-Object System.Drawing.Point(10, 170)
$txtOKeyManual = New-Object System.Windows.Forms.TextBox; $txtOKeyManual.Location = New-Object System.Drawing.Point(20, 30); $txtOKeyManual.Size = New-Object System.Drawing.Size(550, 25); $txtOKeyManual.BackColor = $MauNen.O_Nhap; $txtOKeyManual.ForeColor = $MauNen.Chu
$btnNapKeyManual = New-Object System.Windows.Forms.Button; $btnNapKeyManual.Text = "NẠP KEY"; $btnNapKeyManual.Location = New-Object System.Drawing.Point(580, 28); $btnNapKeyManual.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnNapKeyManual $MauNen.NutMacDinh
$txtHienIID = New-Object System.Windows.Forms.TextBox; $txtHienIID.Location = New-Object System.Drawing.Point(20, 70); $txtHienIID.Size = New-Object System.Drawing.Size(550, 25); $txtHienIID.BackColor=$MauNen.O_Nhap; $txtHienIID.ForeColor=[System.Drawing.Color]::Yellow; $txtHienIID.ReadOnly=$true
$btnLayIID = New-Object System.Windows.Forms.Button; $btnLayIID.Text = "LẤY IID"; $btnLayIID.Location = New-Object System.Drawing.Point(580, 68); $btnLayIID.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnLayIID $MauNen.XanhDuong
$btnMoWebCID = New-Object System.Windows.Forms.Button; $btnMoWebCID.Text = "MỞ TRANG LẤY CID (VISUAL SUPPORT)"; $btnMoWebCID.Location = New-Object System.Drawing.Point(20, 105); $btnMoWebCID.Size = New-Object System.Drawing.Size(700, 30); ThietKeNut $btnMoWebCID $MauNen.NutMacDinh
$txtNhapCID = New-Object System.Windows.Forms.TextBox; $txtNhapCID.Location = New-Object System.Drawing.Point(20, 145); $txtNhapCID.Size = New-Object System.Drawing.Size(550, 25); $txtNhapCID.BackColor=$MauNen.O_Nhap; $txtNhapCID.ForeColor=$MauNen.Chu
$btnNapCID = New-Object System.Windows.Forms.Button; $btnNapCID.Text = "NẠP CID"; $btnNapCID.Location = New-Object System.Drawing.Point(580, 143); $btnNapCID.Size = New-Object System.Drawing.Size(140, 30); ThietKeNut $btnNapCID $MauNen.Do
$grPhoneWin.Controls.AddRange(@($txtOKeyManual, $btnNapKeyManual, $txtHienIID, $btnLayIID, $btnMoWebCID, $txtNhapCID, $btnNapCID))

# --- 3. NÂNG CẤP PHIÊN BẢN ---
$grUpgradeWin = New-Object System.Windows.Forms.GroupBox; $grUpgradeWin.Text = " 3. Nâng cấp phiên bản Windows "; $grUpgradeWin.ForeColor = [System.Drawing.Color]::Gray; $grUpgradeWin.Size = New-Object System.Drawing.Size(740, 80); $grUpgradeWin.Location = New-Object System.Drawing.Point(10, 360)
$cmbPhienBan = New-Object System.Windows.Forms.ComboBox; $cmbPhienBan.Location = New-Object System.Drawing.Point(20, 30); $cmbPhienBan.Size = New-Object System.Drawing.Size(350, 30); $cmbPhienBan.BackColor=$MauNen.O_Nhap; $cmbPhienBan.ForeColor=$MauNen.Chu; [void]$cmbPhienBan.Items.AddRange(@("Home -> Professional", "Pro -> Enterprise", "Pro -> Education")); $cmbPhienBan.SelectedIndex=0
$btnNangCap = New-Object System.Windows.Forms.Button; $btnNangCap.Text = "TIẾN HÀNH NÂNG CẤP"; $btnNangCap.Location = New-Object System.Drawing.Point(400, 28); $btnNangCap.Size = New-Object System.Drawing.Size(320, 35); ThietKeNut $btnNangCap $MauNen.XanhLa
$grUpgradeWin.Controls.AddRange(@($cmbPhienBan, $btnNangCap))

# --- 4. QUẢN LÝ UPDATE ---
$grUpdate = New-Object System.Windows.Forms.GroupBox; $grUpdate.Text = " 4. Quản lý Windows Update "; $grUpdate.ForeColor = [System.Drawing.Color]::Gray; $grUpdate.Size = New-Object System.Drawing.Size(740, 80); $grUpdate.Location = New-Object System.Drawing.Point(10, 450)
$btnUpOn = New-Object System.Windows.Forms.Button; $btnUpOn.Text = "BẬT UPDATE (MẶC ĐỊNH)"; $btnUpOn.Location = New-Object System.Drawing.Point(20, 30); $btnUpOn.Size = New-Object System.Drawing.Size(340, 35); ThietKeNut $btnUpOn $MauNen.XanhLa
$btnUpOff = New-Object System.Windows.Forms.Button; $btnUpOff.Text = "TẮT UPDATE (VĨNH VIỄN)"; $btnUpOff.Location = New-Object System.Drawing.Point(380, 30); $btnUpOff.Size = New-Object System.Drawing.Size(340, 35); ThietKeNut $btnUpOff $MauNen.Do
$grUpdate.Controls.AddRange(@($btnUpOn, $btnUpOff))

# --- 5. HWID (GIỮ NGUYÊN) ---
$btnHWID = New-Object System.Windows.Forms.Button; $btnHWID.Text = "KÍCH HOẠT BẢN QUYỀN SỐ VĨNH VIỄN (HWID) 🔒"; $btnHWID.Location = New-Object System.Drawing.Point(10, 550); $btnHWID.Size = New-Object System.Drawing.Size(740, 60); ThietKeNut $btnHWID $MauNen.XanhDuong

$pnlWin.Controls.AddRange(@($lblTieuDeWin, $grCheckWin, $grPhoneWin, $grUpgradeWin, $grUpdate, $btnHWID))
$khungChinh.Controls.Add($pnlWin)

# ==============================================================================
# LOGIC XỬ LÝ (PHẢI TRÙNG VỚI HÀM TRONG CORE.PS1)
# ==============================================================================

# Hàm nạp key thông minh
function Global:NapKeyWin($k) {
    GhiLog "-> Đang thử nạp Key: $k"
    $res = cscript //nologo $env:windir\system32\slmgr.vbs /ipk $k 2>&1 | Out-String
    if ($res -match "0xC004F069") {
        GhiLog " -> Lỗi Edition! Đang dùng DISM để ép nhận Key..."
        dism /online /set-edition:Professional /productkey:$k /accepteula /norestart | ForEach-Object { GhiLog $_ }
    } else { GhiLog $res.Trim() }
}

$btnShowKeys.Add_Click({
    ChayTacVu "Đọc Key" {
        $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DigitalProductId -ErrorAction SilentlyContinue
        if($reg){$raw=$reg.DigitalProductId;$isWin8=($raw[66]-band 1)-ne 0;$i=24;$c="BCDFGHJKMPQRTVWXY2346789";$k="";if($isWin8){$raw[66]=($raw[66]-band 0xF7)};do{$cur=0;$x=14;do{$cur=$cur*256;$cur=$raw[$x+52]+$cur;$raw[$x+52]=[math]::Floor($cur/24);$cur=$cur%24;$x--}while($x-ge 0);$i--;$k=$c[$cur]+$k;$last=$cur}while($i-ge 0);if($isWin8){$keyp1=$k.Substring(1,$last);$keyp2=$k.Substring($last+1,$k.Length-($last+1));$k=$keyp1+"N"+$keyp2};$insKey="";for($j=0;$j-lt $k.Length;$j++){$insKey+=$k[$j];if(($j+1)%5-eq 0 -and ($j+1)-ne $k.Length){$insKey+="-"}}}else{$insKey="N/A"}
        $oem = (Get-CimInstance -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService").OA3xOriginalProductKey
        [System.Windows.Forms.MessageBox]::Show("KEY ĐANG CÀI: $insKey`nKEY BIOS: $oem", "Thông tin")
    }
})

$btnReactiveOEM.Add_Click({
    ChayTacVu "Reactive OEM" {
        $oem = (Get-CimInstance -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService").OA3xOriginalProductKey
        if ($oem) { ChuyenTab $pnlLog $btnMenuLog; NapKeyWin $oem; cscript //nologo $env:windir\system32\slmgr.vbs /ato | Out-String | ForEach-Object { GhiLog $_.Trim() } }
    }
})

$btnNapKeyManual.Add_Click({ ChayTacVu "Nạp Key" { NapKeyWin $txtOKeyManual.Text.Trim() } })
$btnLayIID.Add_Click({ ChayTacVu "Lấy IID" { $iid=(cscript //nologo $env:windir\system32\slmgr.vbs /dti).Trim(); $txtHienIID.Text=$iid; GhiLog "IID: $iid" } })
$btnMoWebCID.Add_Click({ [System.Diagnostics.Process]::Start("https://visualsupport.microsoft.com/") })
$btnNapCID.Add_Click({ 
    ChayTacVu "Nạp CID" { 
        $c = $txtNhapCID.Text.Trim(); if ($c.Length -lt 40) { return }
        cscript //nologo $env:windir\system32\slmgr.vbs /atp $c | Out-String | ForEach-Object { GhiLog $_.Trim() }
        cscript //nologo $env:windir\system32\slmgr.vbs /ato | Out-String | ForEach-Object { GhiLog $_.Trim() }
    } 
})

$btnNangCap.Add_Click({
    ChayTacVu "Nâng cấp" {
        $sel = $cmbPhienBan.SelectedItem
        $k = switch -wildcard ($sel) { "*Professional*" {"W269N-WFGWX-YVC9B-4J6C9-T83GX"} "*Enterprise*" {"NPPR9-FWDCX-D2C8J-H2M7V-T6WDT"} "*Education*" {"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"} }
        ChuyenTab $pnlLog $btnMenuLog; GhiLog ">>> NÂNG CẤP LÊN $sel..."; NapKeyWin $k
    }
})

$btnUpOn.Add_Click({ ChayTacVu "Bật Update" { ChuyenTab $pnlLog $btnMenuLog; Set-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" "NoAutoUpdate" 0 -Force; Set-Service wuauserv -StartupType Manual; Start-Service wuauserv; GhiLog "-> Đã bật Update." } })
$btnUpOff.Add_Click({ ChayTacVu "Tắt Update" { ChuyenTab $pnlLog $btnMenuLog; Stop-Service wuauserv -Force; Set-Service wuauserv -StartupType Disabled; Set-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" "NoAutoUpdate" 1 -Force; GhiLog "-> Đã tắt Update vĩnh viễn." } })

$btnHWID.Add_Click({ if(XacNhanMatKhau){ ChayTacVu "HWID" { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $Link_HWID = "https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/HWID.cmd"; try { $noiDung = (Invoke-RestMethod -Uri "$Link_HWID`?t=$((Get-Date).Ticks)" -ErrorAction Stop).ToString(); ChayScriptOnline $noiDung "HWID_Activation" } catch { [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)") } } } })
