# ==============================================================================
# [ KHỞI CHẠY ] FILE CHÍNH (MAIN ENTRY)
# ==============================================================================

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

# --- HIỂN THỊ GIAO DIỆN ---
ChuyenTab $pnlWin $btnMenuWin
[void]$form.ShowDialog()