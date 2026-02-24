# ==============================================================================
# [ CÀI ĐẶT ] CẤU HÌNH HỆ THỐNG
# ==============================================================================
$Global:MatKhauAdmin    = "123" 
$Global:PhienBan        = "V31.13 - Outlook Storage Expander"
$Global:PhienBanHienTai = "31.13"

# ==============================================================================
# [ GIAO DIỆN ] CẤU HÌNH MÀU SẮC & PHÔNG CHỮ
# ==============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$MauNen = @{
    Chinh      = [System.Drawing.Color]::FromArgb(30, 30, 30)
    ThanhBen   = [System.Drawing.Color]::FromArgb(45, 45, 48)
    Chu        = [System.Drawing.Color]::WhiteSmoke
    O_Nhap     = [System.Drawing.Color]::FromArgb(50, 50, 50)
    XanhDuong  = [System.Drawing.Color]::FromArgb(0, 122, 204)
    Do         = [System.Drawing.Color]::FromArgb(211, 47, 47)
    Cam        = [System.Drawing.Color]::FromArgb(255, 140, 0)
    XanhLa     = [System.Drawing.Color]::FromArgb(46, 125, 50)
    NutMacDinh = [System.Drawing.Color]::FromArgb(60, 60, 60)
}
$PhongChu = @{
    TieuDe = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    Dam    = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    Code   = New-Object System.Drawing.Font("Consolas", 10)
}

function ThietKeNut($nut, $mau) {
    if ($null -eq $nut) { return }
    $nut.FlatStyle = "Flat"; $nut.BackColor = $mau; $nut.ForeColor = [System.Drawing.Color]::White
    $nut.Font = $PhongChu.Dam; $nut.FlatAppearance.BorderSize = 0; $nut.Cursor = [System.Windows.Forms.Cursors]::Hand
}

# ==============================================================================
# [ HÀM HỖ TRỢ ] GIAO DIỆN & TÁC VỤ
# ==============================================================================
$Global:lblStatus = New-Object System.Windows.Forms.Label
$Global:lblStatus.Text = "Đang xử lý..."
$Global:lblStatus.ForeColor = [System.Drawing.Color]::Yellow
$Global:lblStatus.Font = $PhongChu.Dam
$Global:lblStatus.AutoSize = $true
$Global:lblStatus.Visible = $false
$Global:lblStatus.BackColor = [System.Drawing.Color]::Transparent

$Global:progressBar = New-Object System.Windows.Forms.ProgressBar
$Global:progressBar.Size = New-Object System.Drawing.Size(300, 5)
$Global:progressBar.Style = "Marquee" 
$Global:progressBar.MarqueeAnimationSpeed = 30
$Global:progressBar.Visible = $false

function ChayTacVu {
    param([string]$Ten, [scriptblock]$Code)
    $Global:lblStatus.Text = "⏳ $Ten..."
    $Global:lblStatus.Visible = $true
    $Global:progressBar.Visible = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    
    try { & $Code } 
    catch { [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)") }
    finally {
        if (-not $form.IsDisposed) {
            $Global:lblStatus.Visible = $false
            $Global:progressBar.Visible = $false
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
}

function Clean-Url { param([string]$Url); return $Url -replace '[^\x20-\x7E]', '' }

function ChayScriptOnline($noiDung, $tenFile) {
    if ([string]::IsNullOrWhiteSpace($noiDung)) { [System.Windows.Forms.MessageBox]::Show("Lỗi: Script rỗng!", "Lỗi"); return }
    $duongDan = Join-Path $env:TEMP "$tenFile.cmd"
    [System.IO.File]::WriteAllText($duongDan, ($noiDung -replace "`n", "`r`n"), [System.Text.Encoding]::ASCII)
    ChuyenTab $pnlLog $btnMenuLog
    GhiLog ">>> Đang khởi chạy $tenFile..."
    Start-Process $duongDan -Verb RunAs -Wait
    GhiLog ">>> Đã hoàn tất $tenFile."
}

function XacNhanMatKhau {
    $cuaSoPass = New-Object System.Windows.Forms.Form
    $cuaSoPass.Text = "Admin Check"; $cuaSoPass.Size = New-Object System.Drawing.Size(350, 180)
    $cuaSoPass.StartPosition = "CenterParent"; $cuaSoPass.FormBorderStyle = "FixedDialog"
    $cuaSoPass.BackColor = $MauNen.Chinh; $cuaSoPass.MaximizeBox = $false
    $nhanBao = New-Object System.Windows.Forms.Label; $nhanBao.Text = "Mật khẩu quản trị:"; $nhanBao.Location = New-Object System.Drawing.Point(20, 20); $nhanBao.ForeColor = $MauNen.Chu; $nhanBao.AutoSize = $true
    $oPass = New-Object System.Windows.Forms.TextBox; $oPass.Location = New-Object System.Drawing.Point(20, 50); $oPass.Size = New-Object System.Drawing.Size(290, 30); $oPass.PasswordChar = '*'; $oPass.BackColor = $MauNen.O_Nhap; $oPass.ForeColor = [System.Drawing.Color]::Yellow
    $nutXacNhan = New-Object System.Windows.Forms.Button; $nutXacNhan.Text = "OK"; $nutXacNhan.Location = New-Object System.Drawing.Point(110, 90); $nutXacNhan.Size = New-Object System.Drawing.Size(120, 35); ThietKeNut $nutXacNhan $MauNen.XanhDuong; $nutXacNhan.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cuaSoPass.Controls.AddRange(@($nhanBao, $oPass, $nutXacNhan)); $cuaSoPass.AcceptButton = $nutXacNhan
    if ($cuaSoPass.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return ($oPass.Text -eq $Global:MatKhauAdmin) }
    return $false
}

function HienThiInputBox($tieuDe, $noiDung) {
    $f = New-Object System.Windows.Forms.Form
    $f.Size = New-Object System.Drawing.Size(400, 180); $f.Text = $tieuDe; $f.StartPosition = "CenterParent"
    $f.BackColor = $MauNen.Chinh; $f.FormBorderStyle = "FixedDialog"; $f.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = $noiDung; $lbl.Location = New-Object System.Drawing.Point(20, 20); $lbl.ForeColor = $MauNen.Chu; $lbl.AutoSize = $true
    $txt = New-Object System.Windows.Forms.TextBox; $txt.Location = New-Object System.Drawing.Point(20, 50); $txt.Size = New-Object System.Drawing.Size(340, 30); $txt.BackColor = $MauNen.O_Nhap; $txt.ForeColor = [System.Drawing.Color]::Yellow
    $btn = New-Object System.Windows.Forms.Button; $btn.Text = "XÁC NHẬN"; $btn.Location = New-Object System.Drawing.Point(130, 90); $btn.Size = New-Object System.Drawing.Size(120, 35); ThietKeNut $btn $MauNen.XanhDuong; $btn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.AddRange(@($lbl, $txt, $btn)); $f.AcceptButton = $btn
    if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $txt.Text }
    return $null
}

function Set-Reg {
    param($Path, $Name, $Value, $Type = "DWord")
    if (!(Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force -ErrorAction SilentlyContinue | Out-Null
    GhiLog "Reg: $Name -> $Value"
}

function Set-Svc {
    param($Name, $Type)
    try {
        if (Get-Service $Name -ErrorAction SilentlyContinue) {
            if ($Type -eq "AutomaticDelayedStart") {
                sc.exe config $Name start= delayed-auto | Out-Null
            } else {
                Set-Service -Name $Name -StartupType $Type -ErrorAction SilentlyContinue
            }
            GhiLog "Service: $Name -> $Type"
        }
    } catch {}
}

function Disable-Task {
    param($Path, $Name)
    try {
        Get-ScheduledTask -TaskPath $Path -TaskName $Name -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
        GhiLog "Task Disabled: $Name"
    } catch {}
}

# ==============================================================================
# [ KHUNG GIAO DIỆN CHÍNH ]
# ==============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "AUTO-SOFT SYSTEM MANAGER - $Global:PhienBan"
$form.Size = New-Object System.Drawing.Size(1024, 768); $form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"; $form.MaximizeBox = $false; $form.BackColor = $MauNen.Chinh

$sidebar = New-Object System.Windows.Forms.Panel; $sidebar.Size = New-Object System.Drawing.Size(240, 768); $sidebar.Dock = "Left"; $sidebar.BackColor = $MauNen.ThanhBen
$logo = New-Object System.Windows.Forms.Label; $logo.Text = "QUẢN TRỊ`nHỆ THỐNG"; $logo.Font = $PhongChu.TieuDe; $logo.ForeColor = $MauNen.XanhDuong; $logo.Size = New-Object System.Drawing.Size(240, 80); $logo.TextAlign = "MiddleCenter"; $logo.Dock = "Top"
$lblVer = New-Object System.Windows.Forms.Label; $lblVer.Text = $Global:PhienBan; $lblVer.Dock = "Bottom"; $lblVer.ForeColor = [System.Drawing.Color]::Gray; $lblVer.TextAlign = "MiddleCenter"; $lblVer.Height = 40

$khungChinh = New-Object System.Windows.Forms.Panel; $khungChinh.Dock = "Fill"; $khungChinh.BackColor = $MauNen.Chinh; $khungChinh.Padding = New-Object System.Windows.Forms.Padding(20)
$Global:lblStatus.Location = New-Object System.Drawing.Point(20, 10)
$Global:progressBar.Location = New-Object System.Drawing.Point(20, 35) 
$khungChinh.Controls.AddRange(@($Global:lblStatus, $Global:progressBar))

# --- HÀM CHUYỂN TAB (SAFE MODE) ---
function ChuyenTab($tabHienTai, $nutHienTai) {
    if ($form.IsDisposed -or $null -eq $tabHienTai -or $tabHienTai.IsDisposed) { return }
    
    $pnlWin.Visible=$false; $pnlOff.Visible=$false; $pnlSpec.Visible=$false; $pnlLog.Visible=$false; $pnlPrint.Visible=$false; $pnlOpt.Visible=$false
    foreach($b in @($btnMenuWin,$btnMenuOff,$btnMenuSpec,$btnMenuLog,$btnMenuPrint,$btnMenuOpt)){
        if ($null -ne $b -and -not $b.IsDisposed) { $b.BackColor=$MauNen.ThanhBen;$b.ForeColor=[System.Drawing.Color]::Silver }
    }
    
    $tabHienTai.Visible = $true; 
    if ($null -ne $nutHienTai -and -not $nutHienTai.IsDisposed) { $nutHienTai.BackColor = $MauNen.Chinh; $nutHienTai.ForeColor = $MauNen.XanhDuong }
}

# --- MODULE: NHẬT KÝ ---
$pnlLog = New-Object System.Windows.Forms.Panel; $pnlLog.Dock = "Fill"; $pnlLog.Visible = $false
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.Dock="Fill"; $txtLog.BackColor=[System.Drawing.Color]::Black; $txtLog.ForeColor = $MauNen.XanhLa; $txtLog.Font=$PhongChu.Code; $pnlLog.Controls.Add($txtLog)
$khungChinh.Controls.Add($pnlLog)

function GhiLog($tinNhan) { 
    if (-not $txtLog.IsDisposed) { $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] $tinNhan `r`n") }
}