# ==============================================================================
# SCRIPT MỒI: BẢN CỨNG (FIX ĐÚNG KHO PM / NHÁNH MAIN)
# ==============================================================================

# 1. Ép bảo mật kết nối TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Tạo thư mục tạm và dọn rác
$T = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $T) { Remove-Item -Path "$T\*" -Force -Recurse -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Force -Path $T | Out-Null }

Write-Host ">>> DANG TAI BO CONG CU TU GITHUB..." -ForegroundColor Cyan

# 3. TẢI ĐÍCH DANH TỪNG FILE (Đã fix theo đúng link kho PM của bác)
# Mình dùng Invoke-WebRequest để đảm bảo tính ổn định cao nhất

$G = "https://raw.githubusercontent.com/tuantran19912512/pm/main"

Invoke-WebRequest "$G/Core.ps1" -OutFile "$T\Core.ps1" -UseBasicParsing; Write-Host " -> OK: Core.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Menu.ps1" -OutFile "$T\Menu.ps1" -UseBasicParsing; Write-Host " -> OK: Menu.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Windows.ps1" -OutFile "$T\Tab_Windows.ps1" -UseBasicParsing; Write-Host " -> OK: Tab_Windows.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Office.ps1" -OutFile "$T\Tab_Office.ps1" -UseBasicParsing; Write-Host " -> OK: Tab_Office.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Hardware.ps1" -OutFile "$T\Tab_Hardware.ps1" -UseBasicParsing; Write-Host " -> OK: Tab_Hardware.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Printer.ps1" -OutFile "$T\Tab_Printer.ps1" -UseBasicParsing; Write-Host " -> OK: Tab_Printer.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Optimizer.ps1" -OutFile "$T\Tab_Optimizer.ps1" -UseBasicParsing; Write-Host " -> OK: Tab_Optimizer.ps1" -ForegroundColor Green
Invoke-WebRequest "$G/Main.ps1" -OutFile "$T\Main.ps1" -UseBasicParsing; Write-Host " -> OK: Main.ps1" -ForegroundColor Green

# 4. KHỞI CHẠY TOOL
if (Test-Path "$T\Main.ps1") {
    Write-Host ">>> DANG MO GIAO DIEN..." -ForegroundColor Green
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$T\Main.ps1`""
} else {
    Write-Host "!!! LOI: Khong tim thay file Main.ps1 sau khi tai!" -ForegroundColor Red
    Start-Sleep -Seconds 10
}
exit
