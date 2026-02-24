# ==============================================================================
# [ ĐIỀU HƯỚNG ] MENU & CẬP NHẬT
# ==============================================================================
function TaoNutMenu($ten) {
    $btn = New-Object System.Windows.Forms.Button; $btn.Text = $ten; $btn.Size = New-Object System.Drawing.Size(240, 60); $btn.Dock = "Top"
    $btn.FlatStyle = "Flat"; $btn.FlatAppearance.BorderSize = 0; $btn.Font = $PhongChu.Dam; $btn.ForeColor = [System.Drawing.Color]::Silver; $btn.TextAlign = "MiddleLeft"; $btn.Padding = New-Object System.Windows.Forms.Padding(20, 0, 0, 0); $btn.Cursor = [System.Windows.Forms.Cursors]::Hand; return $btn
}

$btnMenuLog   = TaoNutMenu "NHẬT KÝ (LOG)"
$btnMenuSpec  = TaoNutMenu "CẤU HÌNH CHI TIẾT"
$btnMenuPrint = TaoNutMenu "SỬA LỖI MÁY IN & LAN"
$btnMenuOpt   = TaoNutMenu "TỐI ƯU HỆ THỐNG"
$btnMenuOff   = TaoNutMenu "OFFICE TOOL"
$btnMenuWin   = TaoNutMenu "WINDOWS TOOL"

# Gắn logic chuyển Tab
$btnMenuWin.Add_Click({ ChuyenTab $pnlWin $btnMenuWin })
$btnMenuOff.Add_Click({ ChuyenTab $pnlOff $btnMenuOff })
$btnMenuSpec.Add_Click({ ChuyenTab $pnlSpec $btnMenuSpec })
$btnMenuLog.Add_Click({ ChuyenTab $pnlLog $btnMenuLog })
$btnMenuPrint.Add_Click({ ChuyenTab $pnlPrint $btnMenuPrint })
$btnMenuOpt.Add_Click({ ChuyenTab $pnlOpt $btnMenuOpt })

# Nút Kiểm tra cập nhật
$btnUpdate = New-Object System.Windows.Forms.Button
$btnUpdate.Text = "🔄 Kiểm tra cập nhật"
$btnUpdate.Dock = "Bottom"  # Đổi sang Dock thay vì Location tĩnh để luôn bám sát lề dưới
$btnUpdate.Size = New-Object System.Drawing.Size(240, 40)
$btnUpdate.FlatStyle = "Flat"; $btnUpdate.FlatAppearance.BorderSize = 0; $btnUpdate.ForeColor = "White"; $btnUpdate.BackColor = "#2c3e50"; $btnUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $btnUpdate.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnUpdate.Add_MouseEnter({ if(-not $btnUpdate.IsDisposed){ $btnUpdate.BackColor = "#34495e" } })
$btnUpdate.Add_MouseLeave({ if(-not $btnUpdate.IsDisposed){ $btnUpdate.BackColor = "#2c3e50" } })

$btnUpdate.Add_Click({
    ChayTacVu "Đang kiểm tra Cloud" {
        $maTicks = (Get-Date).Ticks
        $urlCheck = "https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/phienban.txt?t=$maTicks"
        $banMoi = (Invoke-RestMethod -Uri $urlCheck).ToString().Trim()

        if ([double]$banMoi -gt [double]$Global:PhienBanHienTai) {
            $msg = "Phát hiện bản mới: V$banMoi`nBản hiện tại: V$Global:PhienBanHienTai`n`nBạn có muốn cập nhật không?"
            $ask = [System.Windows.Forms.MessageBox]::Show($msg, "Cập Nhật", "YesNo", "Question")
            if ($ask -eq "Yes") {
                $linkMoi = "https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/WLST.ps1?t=$((Get-Date).Ticks)"
                $lenhChay = "powershell -NoProfile -ExecutionPolicy Bypass -Command `"& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm '$linkMoi' | iex }`""
                Start-Process cmd.exe -ArgumentList "/c $lenhChay" -WindowStyle Hidden
                $form.Close(); $form.Dispose()
            }
        } else { [System.Windows.Forms.MessageBox]::Show("Bạn đang dùng bản mới nhất (V$Global:PhienBanHienTai).", "Thông báo") }
    }
})

# AUTO CHECK UPDATE KHI MỞ APP
$form.Add_Shown({
    [System.Windows.Forms.Application]::DoEvents()
    if ($form.IsDisposed) { return }
    try {
        $maTicks = (Get-Date).Ticks
        $urlCheck = "https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/phienban.txt?t=$maTicks"
        $req = [System.Net.WebRequest]::Create($urlCheck); $req.Timeout = 3000
        $res = $req.GetResponse(); $reader = New-Object System.IO.StreamReader($res.GetResponseStream()); $banMoi = $reader.ReadToEnd().Trim(); $reader.Close(); $res.Close()
        if ((-not $form.IsDisposed) -and [double]$banMoi -gt [double]$Global:PhienBanHienTai) {
            $btnUpdate.Text = "🚀 CÓ BẢN MỚI V$banMoi"; $btnUpdate.BackColor = [System.Drawing.Color]::Red; $btnUpdate.ForeColor = [System.Drawing.Color]::Yellow
        }
    } catch {}
})