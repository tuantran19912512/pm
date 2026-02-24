# ==============================================================================
# SCRIPT MỒI: XÁC THỰC NHIỀU NGƯỜI DÙNG QUA GITHUB (HASH SECURITY)
# ==============================================================================

# 1. ĐƯỜNG DẪN FILE TEXT CHỨA DANH SÁCH MÃ HASH (Bác thay link GitHub của bác vào)
$UrlUser = "https://raw.githubusercontent.com/tuantran19912512/pm/main/users.txt"

Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   HE THONG QUAN LY PHAN MEM - PM TOOL    " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host -NoNewline "Vui long nhap ma kich hoat de tiep tuc: " -ForegroundColor Yellow

# Nhập mật khẩu ẩn dấu *
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

# Chuyển pass sang Hash SHA256 để so sánh
$Bytes = [System.Text.Encoding]::UTF8.GetBytes($InputPass)
$HashObj = [System.Security.Cryptography.SHA256]::Create().ComputeHash($Bytes)
$InputHash = [System.BitConverter]::ToString($HashObj).Replace("-", "")

# Tải danh sách Hash từ file text online
try {
    $DanhSachKey = (Invoke-WebRequest -Uri $UrlUser -UseBasicParsing -ErrorAction Stop).Content -split "`n" | ForEach-Object { $_.Trim() }
} catch {
    Write-Host "[!] Loi ket noi server xac thuc!" -ForegroundColor Red
    Start-Sleep -Seconds 3; exit
}

# Kiểm tra mã
if ($DanhSachKey -contains $InputHash) {
    Write-Host "[+] Xac thuc thanh cong! Dang tai du lieu..." -ForegroundColor Green
} else {
    Write-Host "[!] Sai ma kich hoat! Truy cap bi tu choi." -ForegroundColor Red
    Start-Sleep -Seconds 3; exit
}

# 2. TẢI VÀ CHẠY BỘ CÔNG CỤ CHÍNH
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$T = Join-Path $env:TEMP "AutoSoftManager"
if (Test-Path $T) { Remove-Item -Path "$T\*" -Force -Recurse -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Force -Path $T | Out-Null }

$G = "https://raw.githubusercontent.com/tuantran19912512/pm/main"
$Files = @("Core.ps1", "Menu.ps1", "Tab_Windows.ps1", "Tab_Office.ps1", "Tab_Hardware.ps1", "Tab_Printer.ps1", "Tab_Optimizer.ps1", "Main.ps1")

try {
    foreach ($f in $Files) {
        Invoke-WebRequest "$G/$f" -OutFile "$T\$f" -UseBasicParsing
    }
    if (Test-Path "$T\Main.ps1") {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$T\Main.ps1`""
    }
} catch {
    Write-Host "[!] Loi tai du lieu GitHub!" -ForegroundColor Red
}
exit
