# ==============================================================================
# SCRIPT MỒI: BẢN CƯỠNG ÉP (CHỐNG LỖI BIẾN 100%)
# ==============================================================================

# 1. Ép bảo mật kết nối
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Tạo thư mục tạm
$T = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $T) { Remove-Item -Path "$T\*" -Force -Recurse -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Force -Path $T | Out-Null }

Write-Host ">>> DANG TAI DU LIEU..." -ForegroundColor Cyan

# 3. TẢI ĐÍCH DANH TỪNG FILE (Thay vì dùng biến $File dễ bị lỗi)
# Bác chú ý: Phải viết đúng tên file trên GitHub (Hoa/Thường quan trọng lắm)

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Core.ps1" -OutFile "$T\Core.ps1" -UseBasicParsing
Write-Host " -> OK: Core.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Menu.ps1" -OutFile "$T\Menu.ps1" -UseBasicParsing
Write-Host " -> OK: Menu.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Tab_Windows.ps1" -OutFile "$T\Tab_Windows.ps1" -UseBasicParsing
Write-Host " -> OK: Tab_Windows.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Tab_Office.ps1" -OutFile "$T\Tab_Office.ps1" -UseBasicParsing
Write-Host " -> OK: Tab_Office.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Tab_Hardware.ps1" -OutFile "$T\Tab_Hardware.ps1" -UseBasicParsing
Write-Host " -> OK: Tab_Hardware.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Tab_Printer.ps1" -OutFile "$T\Tab_Printer.ps1" -UseBasicParsing
Write-Host " -> OK: Tab_Printer.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Tab_Optimizer.ps1" -OutFile "$T\Tab_Optimizer.ps1" -UseBasicParsing
Write-Host " -> OK: Tab_Optimizer.ps1" -ForegroundColor Green

Invoke-WebRequest "https://raw.githubusercontent.com/tuantran19912512/pm/main/Main.ps1" -OutFile "$T\Main.ps1" -UseBasicParsing
Write-Host " -> OK: Main.ps1" -ForegroundColor Green

# 4. KHỞI CHẠY
Write-Host ">>> DANG MO GIAO DIEN..." -ForegroundColor Green
Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$T\Main.ps1`""
exit
