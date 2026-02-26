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


# ==============================================================================
# NÚT DONATE (GIỮ NGUYÊN TÊN BIẾN LÀ $btnUpdate ĐỂ TOOL TỰ NHẬN DIỆN)
# ==============================================================================
$btnUpdate = New-Object System.Windows.Forms.Button
$btnUpdate.Text = "☕ Ủng hộ tác giả"
$btnUpdate.Dock = "Bottom"
$btnUpdate.Size = New-Object System.Drawing.Size(240, 40)
$btnUpdate.FlatStyle = "Flat"
$btnUpdate.FlatAppearance.BorderSize = 0
$btnUpdate.ForeColor = "White"
$btnUpdate.BackColor = "#2c3e50"
$btnUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnUpdate.Cursor = [System.Windows.Forms.Cursors]::Hand

# Hiệu ứng đổi màu khi chuột lướt qua
$btnUpdate.Add_MouseEnter({ if(-not $btnUpdate.IsDisposed){ $btnUpdate.BackColor = "#34495e" } })
$btnUpdate.Add_MouseLeave({ if(-not $btnUpdate.IsDisposed){ $btnUpdate.BackColor = "#2c3e50" } })

# Logic hiện Popup Mã QR
$btnUpdate.Add_Click({
    $frmDonate = New-Object System.Windows.Forms.Form
    $frmDonate.Text = "Ủng Hộ Tác Giả ☕"
    $frmDonate.Size = New-Object System.Drawing.Size(350, 500)
    $frmDonate.StartPosition = "CenterScreen"
    $frmDonate.FormBorderStyle = "FixedDialog"
    $frmDonate.MaximizeBox = $false
    $frmDonate.MinimizeBox = $false
    $frmDonate.BackColor = [System.Drawing.Color]::White

    $lblThongBao = New-Object System.Windows.Forms.Label
    $lblThongBao.Text = "Cảm ơn bạn đã sử dụng Tool!`nNếu thấy hữu ích, hãy mời tác giả 1 ly cafe nhé!"
    $lblThongBao.AutoSize = $false
    $lblThongBao.Size = New-Object System.Drawing.Size(330, 50)
    $lblThongBao.Location = New-Object System.Drawing.Point(0, 15)
    $lblThongBao.TextAlign = "MiddleCenter"
    $lblThongBao.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblThongBao.ForeColor = [System.Drawing.Color]::Black

    # Khung chứa ảnh QR
    $picQR = New-Object System.Windows.Forms.PictureBox
    $picQR.Size = New-Object System.Drawing.Size(250, 250)
    $picQR.Location = New-Object System.Drawing.Point(40, 75)
    $picQR.SizeMode = "Zoom"
    
    # LINK ẢNH RAW TRÊN GITHUB
    $LinkAnhQR = "https://github.com/tuantran19912512/pm/blob/main/QR.jpg?raw=true" 
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $picQR.Load($LinkAnhQR)
    } catch {
        $lblLoi = New-Object System.Windows.Forms.Label
        $lblLoi.Text = "Không tải được mã QR!`nVui lòng kiểm tra lại kết nối mạng."
        $lblLoi.Location = New-Object System.Drawing.Point(0, 100)
        $lblLoi.Size = New-Object System.Drawing.Size(250, 50)
        $lblLoi.TextAlign = "MiddleCenter"
        $picQR.Controls.Add($lblLoi)
    }

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Ngân hàng: Viettinbank`nSTK: 0938964921`nChủ TK: TRAN THAI TUAN"
    $lblInfo.AutoSize = $false
    $lblInfo.Size = New-Object System.Drawing.Size(330, 70)
    $lblInfo.Location = New-Object System.Drawing.Point(0, 335)
    $lblInfo.TextAlign = "MiddleCenter"
    $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblInfo.ForeColor = [System.Drawing.Color]::DarkBlue

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Đóng"
    $btnClose.Size = New-Object System.Drawing.Size(100, 35)
    $btnClose.Location = New-Object System.Drawing.Point(115, 410)
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.BackColor = [System.Drawing.Color]::LightGray
    $btnClose.Add_Click({ $frmDonate.Close() })

    $frmDonate.Controls.Add($lblThongBao)
    $frmDonate.Controls.Add($picQR)
    $frmDonate.Controls.Add($lblInfo)
    $frmDonate.Controls.Add($btnClose)

    $frmDonate.ShowDialog() | Out-Null
})

