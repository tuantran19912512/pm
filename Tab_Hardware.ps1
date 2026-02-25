# ==============================================================================
# [ TAB 3 ] CẤU HÌNH PHẦN CỨNG
# ==============================================================================
$pnlSpec = New-Object System.Windows.Forms.Panel; $pnlSpec.Dock = "Fill"; $pnlSpec.Visible = $false
$lblTieuDeSpec = New-Object System.Windows.Forms.Label; $lblTieuDeSpec.Text = "THÔNG TIN CHI TIẾT PHẦN CỨNG"; $lblTieuDeSpec.Font = $PhongChu.TieuDe; $lblTieuDeSpec.ForeColor = $MauNen.Cam; $lblTieuDeSpec.Size = New-Object System.Drawing.Size(600, 40); $lblTieuDeSpec.Location = New-Object System.Drawing.Point(10, 10)
$btnBatDauQuet = New-Object System.Windows.Forms.Button; $btnBatDauQuet.Text = "BẮT ĐẦU QUÉT HỆ THỐNG"; $btnBatDauQuet.Location = New-Object System.Drawing.Point(10, 60); $btnBatDauQuet.Size = New-Object System.Drawing.Size(740, 50); ThietKeNut $btnBatDauQuet $MauNen.Cam
$txtHienThiSpec = New-Object System.Windows.Forms.TextBox; $txtHienThiSpec.Location = New-Object System.Drawing.Point(10, 130); $txtHienThiSpec.Size = New-Object System.Drawing.Size(740, 480); $txtHienThiSpec.Multiline = $true; $txtHienThiSpec.ReadOnly = $true; $txtHienThiSpec.ScrollBars = "Vertical"; $txtHienThiSpec.BackColor = [System.Drawing.Color]::FromArgb(40,40,40); $txtHienThiSpec.ForeColor = [System.Drawing.Color]::Cyan; $txtHienThiSpec.Font = $PhongChu.Code
$pnlSpec.Controls.AddRange(@($lblTieuDeSpec, $btnBatDauQuet, $txtHienThiSpec))
$khungChinh.Controls.Add($pnlSpec)

$btnBatDauQuet.Add_Click({
    ChayTacVu "Đang quét phần cứng" {
        $os = Get-CimInstance Win32_OperatingSystem
        $cs = Get-CimInstance Win32_ComputerSystem
        $mb = Get-CimInstance Win32_BaseBoard
        $bios = Get-CimInstance Win32_BIOS
        $cpu = Get-CimInstance Win32_Processor

        $info =  "================ HỆ THỐNG ================`r`n"
        $info += "Hệ điều hành: $($os.Caption) (Bản $($os.BuildNumber))`r`nMáy:          $($cs.Manufacturer) - $($cs.Model)`r`n"
        
        $info += "`r`n================ MAINBOARD ================`r`n"
        $info += "Hãng:     $($mb.Manufacturer)`r`nModel:    $($mb.Product)`r`nSerial:   $($mb.SerialNumber)`r`nBIOS:     $($bios.SMBIOSBIOSVersion) ($($bios.ReleaseDate.ToString('dd/MM/yyyy')))`r`n"
        
        $info += "`r`n================ CPU (VI XỬ LÝ) ================`r`n"
        $info += "Tên:      $($cpu.Name)`r`nSocket:   $($cpu.SocketDesignation)`r`nTốc độ:   $([math]::Round($cpu.MaxClockSpeed/1000, 2)) GHz`r`nNhân/L:   $($cpu.NumberOfCores) Nhân / $($cpu.NumberOfLogicalProcessors) Luồng`r`n"
        
        $memArr = Get-CimInstance Win32_PhysicalMemoryArray
        $dimms = Get-CimInstance Win32_PhysicalMemory
        function LayLoaiRam($c) { switch($c){20{"DDR"} 21{"DDR2"} 24{"DDR3"} 26{"DDR4"} 34{"DDR5"} Default{"Không rõ"}} }
        $info += "`r`n================ RAM (BỘ NHỚ) ================`r`n"
        $info += "Khe cắm:  $(@($dimms).Count) dùng / $($memArr.MemoryDevices) tổng`r`n"
        foreach ($d in $dimms) { 
            $cap = [math]::Round($d.Capacity/1GB, 0)
            $info += "- Slot $($d.DeviceLocator): ${cap}GB $(LayLoaiRam $d.SMBIOSMemoryType) ($($d.Speed)MHz) | Hãng: $($d.Manufacturer) | P/N: $($d.PartNumber.Trim())`r`n" 
        }
        
        $info += "`r`n================ Ổ CỨNG ================`r`n"
        foreach ($d in Get-CimInstance Win32_DiskDrive) {
            $ws = [math]::Round($d.Size/1GB, 0)
            $ms = [math]::Round($d.Size/1000000000, 0)
            $lblM = if($ms -ge 1000){"$([math]::Round($ms/1000,0)) TB"}else{"$ms GB"}
            $info += "● $($d.Model) ($($d.MediaType))`r`n  Windows: ${ws} GB | Niêm yết: $lblM`r`n"
        }

        # ==========================================
        # LOAD THÔNG TIN GPU (CARD MÀN HÌNH)
        # ==========================================
        $info += "`r`n================ GPU (CARD ĐỒ HỌA) ================`r`n"
        $GPUs = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
        if ($GPUs) {
            foreach ($gpu in $GPUs) {
                # Tính VRAM
                $vram = "Không xác định"
                if ($gpu.AdapterRAM) {
                    if ($gpu.AdapterRAM -ge 1GB) {
                        $vram = [math]::Round($gpu.AdapterRAM / 1GB, 2).ToString() + " GB"
                    } else {
                        $vram = [math]::Round($gpu.AdapterRAM / 1MB, 0).ToString() + " MB"
                    }
                }

                # Lấy độ phân giải
                $res = "Chưa kết nối màn hình"
                if ($gpu.CurrentHorizontalResolution -and $gpu.CurrentVerticalResolution) {
                    $res = "$($gpu.CurrentHorizontalResolution) x $($gpu.CurrentVerticalResolution) @ $($gpu.CurrentRefreshRate)Hz"
                }

                $info += "● $($gpu.Name)`r`n"
                $info += "  + Bộ nhớ VRAM : $vram`r`n"
                $info += "  + Màn hình    : $res`r`n"
                $info += "  + Bản Driver  : $($gpu.DriverVersion)`r`n`r`n"
            }
        } else {
            $info += "Không thể đọc thông tin Card đồ họa.`r`n"
        }

        # Đẩy toàn bộ text lên màn hình
        $txtHienThiSpec.Text = $info
    }
})