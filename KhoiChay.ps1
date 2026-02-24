# ==============================================================================
# SCRIPT MỒI: FIX LỖI RÁP LINK (VERSION CHUẨN 100%)
# ==============================================================================

# 1. Ép bảo mật TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. ĐỊA CHỈ GỐC (Bác giữ nguyên link này, KHÔNG thêm dấu / ở cuối)
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

Write-Host ">>> DANG TAI DU LIEU TU KHO [PM]..." -ForegroundColor Cyan

# 4. Tải file với cơ chế ráp link chính xác
$MaRandom = Get-Random
foreach ($File in $DanhSachFile) {
    # Ráp thủ công để đảm bảo: LinkGoc + / + Tên File + ?v= + Mã Random
    $LinkTai = "$LinkGoc/$($File)?v=$MaRandom"
    $DuongDanLuu = Join-Path $ThuMucTam $File
    
    try {
        Write-Host "-> Đang tải: $File" -ForegroundColor Yellow
        Invoke-WebRequest -Uri $LinkTai -OutFile $DuongDanLuu -UseBasicParsing -ErrorAction Stop
        Write-Host "   => OK!" -ForegroundColor Green
    } catch {
        Write-Host "[!] LOI: Khong tai duoc $File" -ForegroundColor Red
        Write-Host "   => Link bi loi: $LinkTai" -ForegroundColor Gray
        Start-Sleep -Seconds 10
        exit
    }
}

# 5. Khởi chạy
$FileMain = Join-Path $ThuMucTam "Main.ps1"
if (Test-Path $FileMain) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$FileMain`""
}
exit
