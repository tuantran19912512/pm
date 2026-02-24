# ==============================================================================
# SCRIPT MỒI: BẢN FIX CỨNG ĐƯỜNG DẪN KHO [PM]
# ==============================================================================

# 1. Ép dùng chuẩn bảo mật TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. ĐỊA CHỈ GỐC - Bác dùng chính xác định dạng này (Đã bỏ /refs/heads/)
$LinkGoc = "https://raw.githubusercontent.com/tuantran19912512/pm/main"

$DanhSachFile = @(
    "Main.ps1", "Core.ps1", "Menu.ps1", 
    "Tab_Windows.ps1", "Tab_Office.ps1", 
    "Tab_Hardware.ps1", "Tab_Printer.ps1", "Tab_Optimizer.ps1"
)

# 3. Tạo/Dọn dẹp thư mục tạm
$ThuMucTam = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $ThuMucTam) { 
    Remove-Item -Path "$ThuMucTam\*" -Force -Recurse -ErrorAction SilentlyContinue 
} else { 
    New-Item -ItemType Directory -Force -Path $ThuMucTam | Out-Null 
}

Write-Host ">>> ĐANG TẢI DỮ LIỆU TỪ GITHUB (KHO PM)..." -ForegroundColor Cyan

# 4. Tải file với cơ chế ép làm mới (Bypass Cache)
$MaRandom = Get-Random
foreach ($File in $DanhSachFile) {
    # Link ráp chuẩn: https://raw.githubusercontent.com/tuantran19912512/pm/main/Main.ps1?v=123
    $LinkTai = "$LinkGoc/$File?v=$MaRandom"
    $DuongDanLuu = Join-Path $ThuMucTam $File
    
    try {
        # Sử dụng Invoke-WebRequest để kiểm soát lỗi chặt chẽ hơn
        Invoke-WebRequest -Uri $LinkTai -OutFile $DuongDanLuu -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        Write-Host "-> Đã tải xong: $File" -ForegroundColor Green
    } catch {
        Write-Host "[!] KHÔNG TẢI ĐƯỢC: $File" -ForegroundColor Red
        Write-Host "-> Kiểm tra link này trên trình duyệt: $LinkTai" -ForegroundColor Gray
        Write-Host "-> Chi tiết lỗi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Start-Sleep -Seconds 10
        exit
    }
}

# 5. Khởi chạy
$FileMain = Join-Path $ThuMucTam "Main.ps1"
if (Test-Path $FileMain) {
    Write-Host ">>> TẤT CẢ FILE ĐÃ SẴN SÀNG. ĐANG KHỞI CHẠY..." -ForegroundColor Green
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$FileMain`""
}
exit
