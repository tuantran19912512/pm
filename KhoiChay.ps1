# ==============================================================================
# SCRIPT MỒI: BẢO MẬT HASH SHA256 (KHÔNG LỘ PASS TRÊN GITHUB)
# ==============================================================================

# 1. MÃ BĂM MẬT KHẨU (Thay dãy dưới đây bằng dãy bác vừa tạo ở Bước 1)
$HashChuan = "95C551F5E06934F3351B512BC7CC299FF41E5CC895590338991E24A9DFA8B9FC" # Đây là mã của 'Kto@2026'

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   HE THONG QUAN LY PHAN MEM - PM TOOL    " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host -NoNewline "Vui long nhap mat khau admin: " -ForegroundColor Yellow

# Nhập mật khẩu ẩn
$InputPass = ""
while($true) {
    $Key = [Console]::ReadKey($true)
    if($Key.Key -eq "Enter") { Write-Host ""; break }
    if($Key.Key -eq "Backspace") {
        if($InputPass.Length -gt 0) {
            $InputPass = $InputPass.SubString(0, $InputPass.Length - 1)
            Write-Host -NoNewline "`b `b"
        }
    } else {
        $InputPass += $Key.KeyChar
        Write-Host -NoNewline "*"
    }
}

# Chuyển pass vừa nhập sang Hash để so sánh
$Bytes = [System.Text.Encoding]::UTF8.GetBytes($InputPass)
$HashObj = [System.Security.Cryptography.SHA256]::Create().ComputeHash($Bytes)
$InputHash = [System.BitConverter]::ToString($HashObj).Replace("-", "")

if ($InputHash -ne $HashChuan) {
    Write-Host "[!] Sai mat khau! Khong the truy cap." -ForegroundColor Red
    Start-Sleep -Seconds 3
    exit
}

Write-Host "[+] Xac thuc thanh cong! Dang tai tool..." -ForegroundColor Green
# 1. Ép bảo mật kết nối TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Tạo thư mục tạm và dọn rác
$T = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $T) { Remove-Item -Path "$T\*" -Force -Recurse -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Force -Path $T | Out-Null }

Write-Host ">>> Dang tai source..." -ForegroundColor Cyan

# 3. TẢI ĐÍCH DANH TỪNG FILE (Đã fix theo đúng link kho PM của bác)
# Mình dùng Invoke-WebRequest để đảm bảo tính ổn định cao nhất

$G = "https://raw.githubusercontent.com/tuantran19912512/pm/main"

Invoke-WebRequest "$G/Core.ps1" -OutFile "$T\Core.ps1" -UseBasicParsing; Write-Host " -> Complete 1" -ForegroundColor Green
Invoke-WebRequest "$G/Menu.ps1" -OutFile "$T\Menu.ps1" -UseBasicParsing; Write-Host "  -> Complete 2" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Windows.ps1" -OutFile "$T\Tab_Windows.ps1" -UseBasicParsing; Write-Host "  -> Complete 3" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Office.ps1" -OutFile "$T\Tab_Office.ps1" -UseBasicParsing; Write-Host "  -> Complete 4" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Hardware.ps1" -OutFile "$T\Tab_Hardware.ps1" -UseBasicParsing; Write-Host "  -> Complete 5" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Printer.ps1" -OutFile "$T\Tab_Printer.ps1" -UseBasicParsing; Write-Host "  -> Complete 6" -ForegroundColor Green
Invoke-WebRequest "$G/Tab_Optimizer.ps1" -OutFile "$T\Tab_Optimizer.ps1" -UseBasicParsing; Write-Host "  -> Complete 7" -ForegroundColor Green
Invoke-WebRequest "$G/Main.ps1" -OutFile "$T\Main.ps1" -UseBasicParsing; Write-Host "  -> Complete 8" -ForegroundColor Green

# 4. KHỞI CHẠY TOOL
if (Test-Path "$T\Main.ps1") {
    Write-Host ">>> DANG MO GIAO DIEN..." -ForegroundColor Green
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$T\Main.ps1`""
} else {
    Write-Host "!!! LOI: Khong tim thay file Main.ps1 sau khi tai!" -ForegroundColor Red
    Start-Sleep -Seconds 10
}
exit



