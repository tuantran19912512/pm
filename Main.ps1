# ==============================================================================
# [ KHỞI CHẠY ] FILE CHÍNH (MAIN ENTRY)
# ==============================================================================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# --- KIỂM TRA QUYỀN ADMIN ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Lấy đường dẫn thư mục hiện tại của script
$ScriptPath = $PSScriptRoot

# --- NẠP CÁC MODULE (DOT-SOURCING) ---
. "$ScriptPath\Core.ps1"
. "$ScriptPath\Menu.ps1"
. "$ScriptPath\Tab_Windows.ps1"
. "$ScriptPath\Tab_Office.ps1"
. "$ScriptPath\Tab_Hardware.ps1"
. "$ScriptPath\Tab_Printer.ps1"
. "$ScriptPath\Tab_Optimizer.ps1"

# ==============================================================================
# [ LẮP RÁP GIAO DIỆN ] SẮP XẾP Z-ORDER CHUẨN ĐỂ KHÔNG BỊ ĐÈ
# ==============================================================================

# 1. Lắp ráp Sidebar (Thứ tự Add quyết định Dock)
# Add Bottom trước
$sidebar.Controls.Add($btnUpdate)
$sidebar.Controls.Add($lblVer)
# Add Top (Lưu ý: Mảng xếp từ Log -> Logo thì Logo sẽ nằm trên cùng nhất)
$sidebar.Controls.AddRange(@($btnMenuLog, $btnMenuSpec, $btnMenuPrint, $btnMenuOpt, $btnMenuOff, $btnMenuWin, $logo))

# 2. Đưa Khung chính và Sidebar vào Form
$form.Controls.Add($sidebar)
$form.Controls.Add($khungChinh)

# 3. Ép Z-order để khung chính dạt sang phải, không chui xuống dưới Sidebar
$khungChinh.BringToFront()
# ==============================================================================
# TỰ ĐỘNG DỌN DẸP RÁC (CACHE) KHI TẮT TOOL
# ==============================================================================
$form.Add_FormClosed({
    # Đường dẫn thư mục tạm lúc KhoiChay.ps1 tải về
    $ThuMucRanh = Join-Path $env:TEMP "AutoSoftManager"
    
    # Lệnh CMD ngầm: Đợi 2 giây -> Xóa sạch thư mục (bỏ qua lỗi nếu có)
    $LenhXoa = "/c timeout /t 2 /nobreak >nul & rd /s /q `"$ThuMucRanh`""
    
    # Chạy CMD ẩn để dọn dẹp sau khi PowerShell đã thoát
    Start-Process cmd.exe -ArgumentList $LenhXoa -WindowStyle Hidden
})
# --- HIỂN THỊ GIAO DIỆN ---
ChuyenTab $pnlWin $btnMenuWin

[void]$form.ShowDialog()

