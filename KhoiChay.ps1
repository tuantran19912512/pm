# ==============================================================================
# SCRIPT MỒI: BẢN THỦ CÔNG (CHỐNG LỖI RÁP BIẾN 100%)
# ==============================================================================

# 1. Ép bảo mật TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Cấu hình đường dẫn
$Goc = "https://raw.githubusercontent.com/tuantran19912512/pm/main"
$Tam = Join-Path $env:TEMP "AutoSoftManager"

# 3. Tạo thư mục tạm
if (Test-Path $Tam) { Remove-Item -Path "$Tam\*" -Force -Recurse -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Force -Path $Tam | Out-Null }

Write-Host ">>> DANG TAI DU LIEU TU GITHUB..." -ForegroundColor Cyan

# 4. HÀM TẢI FILE SIÊU CẤP (Viết riêng để bắt lỗi từng file)
function TaiFile($TenFile) {
    $L = "$Goc/$TenFile?v=$(Get-Random)"
    $D = Join-Path $Tam $TenFile
    try {
        Invoke-WebRequest -Uri $L -OutFile $D -UseBasicParsing -ErrorAction Stop
        Write-Host " -> OK: $TenFile" -ForegroundColor Green
    } catch {
        Write-Host " [!] LOI: Khong tai duoc $TenFile" -ForegroundColor Red
        Write-Host "     Link loi: $L" -ForegroundColor Gray
    }
}

# 5. LIỆT KÊ ĐÍCH DANH TỪNG FILE (Không dùng vòng lặp biến để tránh lỗi 400)
TaiFile "Core.ps1"
TaiFile "Menu.ps1"
TaiFile "Tab_Windows.ps1"
TaiFile "Tab_Office.ps1"
TaiFile "Tab_Hardware.ps1"
TaiFile "Tab_Printer.ps1"
TaiFile "Tab_Optimizer.ps1"
TaiFile "Main.ps1"

# 6. Kiểm tra file cuối và chạy
if (Test-Path "$Tam\Main.ps1") {
    Write-Host ">>> KHOI CHAY..." -ForegroundColor Green
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Tam\Main.ps1`""
} else {
    Write-Host "!!! THAT BAI: Thieu file Main.ps1. Hay kiem tra lai ten kho GitHub!" -ForegroundColor Red
    Start-Sleep -Seconds 10
}
exit
