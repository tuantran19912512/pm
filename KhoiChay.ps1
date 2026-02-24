# ==============================================================================
# SCRIPT MỒI: BẢN THỦ CÔNG (CHỐNG LỖI RÁP BIẾN 100%)
# ==============================================================================

# 1. Ép bảo mật TLS 1.2 để không bị GitHub chặn
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Cấu hình đường dẫn gốc (Bác kiểm tra kỹ chữ "pm" và "main")
$Goc = "https://raw.githubusercontent.com/tuantran19912512/pm/main"
$Tam = Join-Path $env:TEMP "AutoSoftManager"

# 3. Tạo thư mục tạm và dọn rác cũ
if (Test-Path $Tam) { Remove-Item -Path "$Tam\*" -Force -Recurse -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Force -Path $Tam | Out-Null }

Write-Host ">>> ĐANG TẢI DỮ LIỆU TỪ GITHUB..." -ForegroundColor Cyan

# 4. HÀM TẢI FILE ĐÍCH DANH (Cực kỳ ổn định)
function TaiFile($TenFile) {
    $Url = "$Goc/$TenFile?v=$(Get-Random)"
    $Path = Join-Path $Tam $TenFile
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Path -UseBasicParsing -ErrorAction Stop
        Write-Host " -> OK: $TenFile" -ForegroundColor Green
    } catch {
        Write-Host " [!] LỖI: Không thấy $TenFile" -ForegroundColor Red
        Write-Host "     Link lỗi: $Url" -ForegroundColor Gray
    }
}

# 5. TẢI TỪNG FILE MỘT (Không dùng vòng lặp để tránh lỗi mất biến)
TaiFile "Core.ps1"
TaiFile "Menu.ps1"
TaiFile "Tab_Windows.ps1"
TaiFile "Tab_Office.ps1"
TaiFile "Tab_Hardware.ps1"
TaiFile "Tab_Printer.ps1"
TaiFile "Tab_Optimizer.ps1"
TaiFile "Main.ps1"

# 6. Kiểm tra và Khởi chạy
if (Test-Path "$Tam\Main.ps1") {
    Write-Host ">>> TẤT CẢ ĐÃ SẴN SÀNG! ĐANG KHỞI CHẠY..." -ForegroundColor Green
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Tam\Main.ps1`""
} else {
    Write-Host "!!! THẤT BẠI: File Main.ps1 không tồn tại trên GitHub!" -ForegroundColor Red
    Start-Sleep -Seconds 10
}
exit
