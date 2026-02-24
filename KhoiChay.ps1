# ==============================================================================
# SCRIPT MỒI: CHỐNG LỖI 400 INVALID REQUEST
# ==============================================================================

# 1. Ép dùng TLS 1.2 - Thiếu dòng này là GitHub báo Invalid ngay
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. ĐỊA CHỈ GỐC CHUẨN (Không để dấu / ở cuối cùng)
$LinkGoc = "https://raw.githubusercontent.com/tuantran19912512/pm/main"

$DanhSachFile = @(
    "Main.ps1", "Core.ps1", "Menu.ps1", 
    "Tab_Windows.ps1", "Tab_Office.ps1", 
    "Tab_Hardware.ps1", "Tab_Printer.ps1", "Tab_Optimizer.ps1"
)

# 3. Làm sạch thư mục tạm
$ThuMucTam = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $ThuMucTam) { 
    Remove-Item -Path "$ThuMucTam\*" -Force -Recurse -ErrorAction SilentlyContinue 
} else { 
    New-Item -ItemType Directory -Force -Path $ThuMucTam | Out-Null 
}

Write-Host ">>> DANG KEO DU LIEU TU KHO [PM]..." -ForegroundColor Cyan

# 4. Tải file - Dùng vòng lặp sạch
$MaRandom = Get-Random
foreach ($File in $DanhSachFile) {
    # Đảm bảo đường dẫn không bị dư dấu /
    $LinkTai = "$($LinkGoc.TrimEnd('/'))/$File?v=$MaRandom"
    $DuongDanLuu = Join-Path $ThuMucTam $File
    
    try {
        # Dùng Invoke-WebRequest để tránh lỗi 400 của Invoke-RestMethod
        Invoke-WebRequest -Uri $LinkTai -OutFile $DuongDanLuu -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        Write-Host "-> Thanh cong: $File" -ForegroundColor Green
    } catch {
        Write-Host "[!] LOI 400 HOAC 404: $File" -ForegroundColor Red
        Write-Host "-> Link thu nghiem: $LinkTai" -ForegroundColor Gray
        Start-Sleep -Seconds 5
        exit
    }
}

# 5. Khởi chạy
$FileMain = Join-Path $ThuMucTam "Main.ps1"
if (Test-Path $FileMain) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$FileMain`""
}
exit
