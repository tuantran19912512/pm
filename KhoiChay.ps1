# ==============================================================================
# SCRIPT MỒI: BẢN CHUẨN XÁC ĐỊA CHỈ GITHUB (FIX LỖI TẢI FILE)
# ==============================================================================

# 1. Ép dùng TLS 1.2 để thông suốt kết nối với GitHub
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. ĐỊA CHỈ GỐC CHUẨN (Bác dùng link RAW rút gọn này cho sạch)
$LinkGoc = "https://raw.githubusercontent.com/tuantran19912512/scripts/main"

$DanhSachFile = @(
    "Main.ps1", "Core.ps1", "Menu.ps1", 
    "Tab_Windows.ps1", "Tab_Office.ps1", 
    "Tab_Hardware.ps1", "Tab_Printer.ps1", "Tab_Optimizer.ps1"
)

# 3. Dọn dẹp và tạo thư mục tạm
$ThuMucTam = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $ThuMucTam) { 
    Remove-Item -Path "$ThuMucTam\*" -Force -Recurse -ErrorAction SilentlyContinue 
} else { 
    New-Item -ItemType Directory -Force -Path $ThuMucTam | Out-Null 
}

Write-Host ">>> ĐANG KÉO DỮ LIỆU TỪ GITHUB..." -ForegroundColor Cyan

# 4. Tải file (Sử dụng lệnh Invoke-RestMethod để nhẹ máy)
$MaRandom = Get-Random
foreach ($File in $DanhSachFile) {
    # Link này sẽ tự ráp thành: .../scripts/main/Main.ps1?v=12345
    $LinkTai = "$LinkGoc/$File?v=$MaRandom"
    $DuongDanLuu = Join-Path $ThuMucTam $File
    
    try {
        $NoiDung = Invoke-RestMethod -Uri $LinkTai -UseBasicParsing
        [System.IO.File]::WriteAllText($DuongDanLuu, $NoiDung, [System.Text.Encoding]::UTF8)
        Write-Host "-> Đã tải xong: $File" -ForegroundColor Green
    } catch {
        Write-Host "[!] THẤT BẠI: $File. Kiểm tra lại tên file trên GitHub!" -ForegroundColor Red
        Start-Sleep -Seconds 5
        exit
    }
}

# 5. Khởi chạy Tool ngầm
$FileMain = Join-Path $ThuMucTam "Main.ps1"
if (Test-Path $FileMain) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$FileMain`""
}
exit