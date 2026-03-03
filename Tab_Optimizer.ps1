# ==============================================================================
# [ TAB 5 ] TOI UU HOA HE THONG (BAN RUT GON - 3 NUT FULL OPTION)
# ==============================================================================
$pnlOpt = New-Object System.Windows.Forms.Panel; $pnlOpt.Dock = "Fill"; $pnlOpt.Visible = $false
$lblOpt = New-Object System.Windows.Forms.Label; $lblOpt.Text = "TOI UU HOA HE THONG (FULL OPTION)"; $lblOpt.Font = $PhongChu.TieuDe; $lblOpt.ForeColor = $MauNen.Cam; $lblOpt.Size = New-Object System.Drawing.Size(600, 40); $lblOpt.Location = New-Object System.Drawing.Point(10, 10)

$txtOptWarn = New-Object System.Windows.Forms.Label; $txtOptWarn.Text = "LUU Y: Vui long luu lai cac cong viec dang lam truoc khi chay Toi uu. He thong se tu dong khoi dong lai Explorer hoac dong cac ung dung dang mo."; $txtOptWarn.ForeColor = [System.Drawing.Color]::Gray; $txtOptWarn.Size = New-Object System.Drawing.Size(740, 40); $txtOptWarn.Location = New-Object System.Drawing.Point(10, 60)

# --- NÚT 1: TỐI ƯU WINDOWS FULL OPTION ---
$btnOptWinFull = New-Object System.Windows.Forms.Button
$btnOptWinFull.Text = "1. TOI UU HOA WINDOWS FULL OPTION (DON RAC, CPU, RAM, MANG)"
$btnOptWinFull.Location = New-Object System.Drawing.Point(10, 110)
$btnOptWinFull.Size = New-Object System.Drawing.Size(740, 80)
$btnOptWinFull.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
ThietKeNut $btnOptWinFull $MauNen.Do

# --- NÚT 2: TỐI ƯU OFFICE FULL OPTION ---
$btnOptOfficeFull = New-Object System.Windows.Forms.Button
$btnOptOfficeFull.Text = "2. TOI UU HOA OFFICE (WORD / EXCEL / POWERPOINT)"
$btnOptOfficeFull.Location = New-Object System.Drawing.Point(10, 210)
$btnOptOfficeFull.Size = New-Object System.Drawing.Size(740, 80)
$btnOptOfficeFull.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
ThietKeNut $btnOptOfficeFull $MauNen.XanhDuong

# --- NÚT 3: ĐỒNG BỘ GIỜ VIỆT NAM ---
$btnDongBoGio = New-Object System.Windows.Forms.Button
$btnDongBoGio.Text = "3. DONG BO GIO VIET NAM (CHONG LOI CHUNG CHI WEB)"
$btnDongBoGio.Location = New-Object System.Drawing.Point(10, 310)
$btnDongBoGio.Size = New-Object System.Drawing.Size(740, 80)
$btnDongBoGio.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnDongBoGio.BackColor = [System.Drawing.Color]::FromArgb(255, 39, 174, 96) # Màu xanh lá
$btnDongBoGio.ForeColor = [System.Drawing.Color]::White
$btnDongBoGio.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDongBoGio.FlatAppearance.BorderSize = 0
$btnDongBoGio.Cursor = [System.Windows.Forms.Cursors]::Hand

# Gắn TẤT CẢ các nút vào khung giao diện cùng lúc
$pnlOpt.Controls.AddRange(@($lblOpt, $txtOptWarn, $btnOptWinFull, $btnOptOfficeFull, $btnDongBoGio))
$khungChinh.Controls.Add($pnlOpt)

# ==============================================================================
# LOGIC 1: TOI UU WINDOWS FULL OPTION (TICH HOP FULL SCRIPT CHRIS TITUS)
# ==============================================================================
$btnOptWinFull.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Ban co muon chay Toi uu Windows FULL OPTION (Bao gom Go App Rac & Full Chris Titus Script) khong?", "Xac nhan", "YesNo", "Warning") -eq "Yes") {
        ChayTacVu "Dang toi uu Windows FULL..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BAT DAU TOI UU HOA WINDOWS FULL OPTION <<<"

            GhiLog "B1: Tao System Restore Point (Diem khoi phuc an toan)..."
            try { Checkpoint-Computer -Description "Truoc_Khi_Toi_Uu" -RestorePointType "MODIFY_SETTINGS" -ErrorAction SilentlyContinue } catch {}

            GhiLog "B2: Ep xung CPU (Kich hoat High Performance Plan)..."
            powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c | Out-Null

            GhiLog "B3: Toi uu Mang (Flush DNS giup luot Web muot hon)..."
            ipconfig /flushdns | Out-Null

            GhiLog "B4: Don dep rac he thong chuyen sau (Bypass Cleanmgr)..."
            
            $DanhSachRac = @(
                "$env:TEMP\*", "$env:windir\Temp\*", "$env:windir\Prefetch\*", 
                "$env:windir\SoftwareDistribution\Download\*", 
                "$env:ProgramData\Microsoft\Windows\WER\ReportArchive\*", 
                "$env:ProgramData\Microsoft\Windows\WER\ReportQueue\*", 
                "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\*"
            )
            foreach ($ThuMuc in $DanhSachRac) { Remove-Item -Path $ThuMuc -Recurse -Force -ErrorAction SilentlyContinue }
            Clear-RecycleBin -Force -Confirm:$false -ErrorAction SilentlyContinue
            Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Filter "thumbcache_*.db" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Filter "iconcache_*.db" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

            GhiLog "B5: Thuc thi FULL loi toi uu Chris Titus (Quyen rieng tu & Hieu suat)..."
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableActivityFeed" 0 "DWord"
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "PublishUserActivities" 0 "DWord"
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "UploadUserActivities" 0 "DWord"
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" 1 "DWord"
            Set-Reg "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" "DisableWpbtExecution" 1 "DWord"
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location" "Value" "Deny" "String"
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" "SensorPermissionState" 0 "DWord"
            Set-Reg "HKLM:\SYSTEM\Maps" "AutoUpdateEnabled" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" "Enabled" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" "TailoredExperiencesWithDiagnosticDataEnabled" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy" "HasAccepted" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Input\TIPC" "Enabled" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\InputPersonalization" "RestrictImplicitInkCollection" 1 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\InputPersonalization" "RestrictImplicitTextCollection" 1 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore" "HarvestContacts" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Personalization\Settings" "AcceptedPrivacyPolicy" 0 "DWord"
            Set-Reg "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" "AllowTelemetry" 0 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Siuf\Rules" "NumberOfSIUFInPeriod" 0 "DWord"

            Remove-Item "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags" -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU" -Recurse -Force -ErrorAction SilentlyContinue
            Set-Reg "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell" "FolderType" "NotSpecified" "String"

            GhiLog "B6: Toi uu DICH VU (Services) theo chuan Chris Titus..."
            $ServicesManual = @("ALG", "AppMgmt", "AppReadiness", "Appinfo", "AxInstSV", "BTAGService", "COMSysApp", "CertPropSvc", "CscService", "DevQueryBroker", "DeviceAssociationService", "DeviceInstall", "DisplayEnhancementService", "EFS", "EapHost", "FDResPub", "FrameServer", "FrameServerMonitor", "GraphicsPerfSvc", "HvHost", "InstallService", "IpxlatCfgSvc", "KtmRm", "LicenseManager", "LxpSvc", "MSDTC", "MSiSCSI", "McpManagementService", "MicrosoftEdgeElevationService", "NaturalAuthentication", "NcaSvc", "NcbService", "NcdAutoSetup", "NetSetupSvc", "Netman", "NlaSvc", "PeerDistSvc", "PerfHost", "PhoneSvc", "PlugPlay", "PolicyAgent", "PrintNotify", "PushToInstall", "QWAVE", "RasAuto", "RasMan", "RmSvc", "RpcLocator", "SCPolicySvc", "SCardSvr", "SDRSVC", "SEMgrSvc", "SNMPTRAP", "SSDPSRV", "ScDeviceEnum", "SensorDataService", "SensorService", "SensrSvc", "SessionEnv", "SharedAccess", "SmsRouter", "SstpSvc", "StiSvc", "TapiSrv", "TieringEngineService", "TokenBroker", "TroubleshootingSvc", "TrustedInstaller", "UmRdpService", "VSS", "W32Time", "WEPHOSTSVC", "WFDSConMgrSvc", "WMPNetworkSvc", "WManSvc", "WPDBusEnum", "WSAIFabricSvc", "WalletService", "WarpJITSvc", "WbioSrvc", "WdiServiceHost", "WdiSystemHost", "WebClient", "Wecsvc", "WerSvc", "WiaRpc", "WinRM", "WpcMonSvc", "XblAuthManager", "XblGameSave", "XboxGipSvc", "XboxNetApiSvc", "autotimesvc", "bthserv", "cloudidsvc", "dcsvc", "defragsvc", "diagsvc", "dot3svc", "edgeupdatem", "fdPHost", "fhsvc", "hidserv", "icssvc", "lltdsvc", "lmhosts", "netprofm", "perceptionsimulation", "pla", "seclogon", "smphost", "svsvc", "swprv", "upnphost", "vds", "vmicguestinterface", "vmicheartbeat", "vmickvpexchange", "vmicrdv", "vmicshutdown", "vmictimesync", "vmicvmsession", "vmicvss", "wbengine", "wcncsvc", "webthreatdefsvc", "wercplsupport", "wisvc", "wlidsvc", "wlpasvc", "wmiApSrv", "workfolderssvc", "wuauserv")
            $ServicesDisabled = @("AppVClient", "AssignedAccessManagerSvc", "DialogBlockingService", "NetTcpPortSharing", "RemoteAccess", "RemoteRegistry", "UevAgentService", "shpamsvc", "ssh-agent", "tzautoupdate")

            foreach ($svc in $ServicesManual) { Set-Service -Name $svc -StartupType Manual -ErrorAction SilentlyContinue }
            foreach ($svc in $ServicesDisabled) { 
                Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
                Set-Service -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue 
            }

            GhiLog "B7: Dang go bo cac Ung dung rac (Bloatware) cua Windows..."
            $Bloatwares = @("*BingWeather*", "*BingNews*", "*BingSports*", "*BingFinance*", "*ZuneVideo*", "*ZuneMusic*", "*SkypeApp*", "*YourPhone*", "*3DBuilder*", "*MicrosoftSolitaireCollection*", "*FeedbackHub*", "*MixedReality.Portal*", "*GetHelp*", "*Getstarted*", "*WindowsMaps*", "*OfficeHub*")
            foreach ($app in $Bloatwares) {
                Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.NonRemovable -eq $false } | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            }

            GhiLog "B8: Toi uu Trai nghiem UX (NumLock, Mo This PC, Hien duoi File)..."
            Set-Reg "Registry::HKEY_USERS\.DEFAULT\Control Panel\Keyboard" "InitialKeyboardIndicators" "2" "String"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 1 "DWord"
            Set-Reg "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0 "DWord"

            GhiLog "B9: Khoi phuc Menu chuot phai kieu cu (Classic Context Menu) cho Win 11..."
            $ContextMenuPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
            New-Item -Path $ContextMenuPath -Value "" -Force | Out-Null

            GhiLog "B10: Dang lam moi giao dien he thong (Restarting Explorer)..."
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            
            GhiLog "B11: Toi uu Gaming (Bat Game Mode, Tat Gia toc chuot de chuan Aim)..."
            Set-Reg "HKCU:\Software\Microsoft\GameBar" "AllowAutoGameMode" 1 "DWord"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseSpeed" "0" "String"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseThreshold1" "0" "String"
            Set-Reg "HKCU:\Control Panel\Mouse" "MouseThreshold2" "0" "String"
            Set-Reg "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" "AllowGameDVR" 0 "DWord"

            GhiLog ">>> TOI UU WINDOWS HOAN TAT! <<<"
            [System.Windows.Forms.MessageBox]::Show("Toi uu Windows FULL thanh cong! Menu chuot phai da tro ve kieu Classic.", "Thanh cong")
        }
    }
})

# ==============================================================================
# LOGIC 2: TOI UU OFFICE FULL OPTION
# ==============================================================================
$btnOptOfficeFull.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("He thong se dong tat ca file Word/Excel dang mo de toi uu. Ban da luu bai chua?", "Xac nhan", "YesNo", "Question") -eq "Yes") {
        ChayTacVu "Dang toi uu Office FULL..." {
            ChuyenTab $pnlLog $btnMenuLog
            GhiLog ">>> BAT DAU TOI UU HOA OFFICE FULL OPTION <<<"

            GhiLog "B1: Dong tat ca tien trinh Word, Excel, PowerPoint..."
            Stop-Process -Name "winword", "excel", "powerpnt", "officeclicktorun" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2

            GhiLog "B2: Toi uu Registry cho Word/Excel/PowerPoint..."
            $OfficeVers = @("14.0", "15.0", "16.0")
            foreach ($Ver in $OfficeVers) {
                $Path = "HKCU:\Software\Microsoft\Office\$Ver"
                if (Test-Path $Path) {
                    Set-Reg "$Path\Word\Options" "NoScalingPaperRes" 1
                    Set-Reg "$Path\Word\Options" "MeasurementUnit" 1
                    Set-Reg "$Path\Excel\Options" "AutomatedScaling" 0
                    Set-Reg "$Path\PowerPoint\Options" "DisableHardwareAcceleration" 1
                    Set-Reg "$Path\Common\Graphics" "DisableHardwareAcceleration" 1
                    Set-Reg "$Path\Common\Graphics" "DisableAnimations" 1
                    Set-Reg "$Path\Word\Options" "AutoSpell" 0
                    Set-Reg "$Path\Common\General" "DisableBootToOfficeStart" 1
                    
                    Set-Reg "$Path\Excel\Options" "UseSystemSeparators" 0 "DWord"
                    Set-Reg "$Path\Excel\Options" "DecimalSeparator" "," "String"
                    Set-Reg "$Path\Excel\Options" "ThousandsSeparator" "." "String"
                }
            }

            GhiLog "B3: Ep cau hinh vao Normal.dotm (Can le ND30, Font TNR 14, Ruler)..."
            try {
                $AppData = [Environment]::GetFolderPath("ApplicationData")
                $NormalPath = Join-Path $AppData "Microsoft\Templates\Normal.dotm"
                if (Test-Path $NormalPath) { Set-ItemProperty -Path $NormalPath -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue }

                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                try { $word.Options.MeasurementUnit = 1 } catch {}
                try { $word.Options.NoScalingPaperRes = $true } catch {}

                $doc = $word.NormalTemplate.OpenAsDocument()
                $word.ActiveWindow.DisplayRulers = $true
                $word.ActiveWindow.DisplayVerticalRuler = $true
                
                $doc.PageSetup.PaperSize = 7 
                $doc.PageSetup.TopMargin = 56.7    
                $doc.PageSetup.BottomMargin = 56.7 
                $doc.PageSetup.LeftMargin = 85.05  
                $doc.PageSetup.RightMargin = 42.55 
                
                $doc.Styles.Item("Normal").Font.Name = "Times New Roman"
                $doc.Styles.Item("Normal").Font.Size = 14

                $doc.Save()
                $doc.Close()
                GhiLog " -> Da cau hinh file mau Word (Normal.dotm) thanh cong."
            } catch {
                GhiLog " -> Bo qua buoc Normal.dotm vi chua mo Word lan nao hoac bi loi."
            } finally {
                if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
            }

            GhiLog "B4: Thiet lap Region Windows (Ngay thang dd/MM/yyyy, Don vi cm)..."
            $Intl = "HKCU:\Control Panel\International"
            Set-Reg $Intl "sShortDate" "dd/MM/yyyy" "String"
            Set-Reg $Intl "iMeasure" 0 "String" 
            Set-Reg $Intl "sDecimal" "," "String"
            Set-Reg $Intl "sThousand" "." "String"

            GhiLog "B5: Lam moi giao dien de ap dung ngay lap tuc..."
            try {
                $Sig = '[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)] public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);'
                $UpdateWin = Add-Type -MemberDefinition $Sig -Name "Win32Opt$([Guid]::NewGuid().ToString().Replace('-',''))" -Namespace Win32 -PassThru
                $res = [UIntPtr]::Zero
                $UpdateWin::SendMessageTimeout([IntPtr]0xffff, 0x001A, [UIntPtr]::Zero, "Environment", 0x02, 5000, [ref]$res) | Out-Null
            } catch {}
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue

            GhiLog ">>> TOI UU OFFICE HOAN TAT! <<<"
            [System.Windows.Forms.MessageBox]::Show("Toi uu Office FULL thanh cong!", "Thanh cong")
        }
    }
})

# ==============================================================================
# LOGIC 3: ĐỒNG BỘ GIỜ VIỆT NAM (UTC+7)
# ==============================================================================
$btnDongBoGio.Add_Click({
    ChayTacVu "Đang đồng bộ thời gian hệ thống..." {
        try {
            # Ép múi giờ về chuẩn Việt Nam
            Set-TimeZone -Id "SE Asia Standard Time" -ErrorAction Stop
            
            # Khởi động lại dịch vụ Thời gian
            Start-Service -Name w32time -ErrorAction SilentlyContinue
            Restart-Service -Name w32time -ErrorAction SilentlyContinue
            
            # Ép cập nhật giờ
            w32tm /resync /nowait | Out-Null
            
            [System.Windows.Forms.MessageBox]::Show("Đã chuyển múi giờ về Việt Nam (UTC+7) và cập nhật thời gian chuẩn xác!", "Thành công", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("LỖI: Bạn phải chạy phần mềm bằng quyền Quản trị viên (Run as Administrator) mới có thể đổi giờ hệ thống!", "Thiếu quyền", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})