$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$winmtrPath = Join-Path -Path $scriptPath -ChildPath "winmtr.exe"
$tcpingPath = Join-Path -Path $scriptPath -ChildPath "tcping.exe"
$speedtestPath = Join-Path -Path $scriptPath -ChildPath "speedtest.exe"

#Remove-Item -Path "$env:TEMP\speedtest" -Recurse -Force -ErrorAction SilentlyContinue
#Remove-Item -Path "$env:TEMP\speedtest.zip" -Force -ErrorAction SilentlyContinue


function Get-SystemInfo {
    param (
        [string]$OutputFile
    )
    
    $output = New-Object System.Text.StringBuilder
    
    [void]$output.AppendLine("*****************************************************************************")
    [void]$output.AppendLine("* SYSTEM INFORMATION *")
    [void]$output.AppendLine("*****************************************************************************")
    [void]$output.AppendLine("")
    
    $os = Get-WmiObject Win32_OperatingSystem
    [void]$output.AppendLine("OPERATING SYSTEM")
    [void]$output.AppendLine("----------------")
    [void]$output.AppendLine("OS Name: $($os.Caption)")
    [void]$output.AppendLine("Version: $($os.Version)")
    [void]$output.AppendLine("Build Number: $($os.BuildNumber)")
    [void]$output.AppendLine("Architecture: $($os.OSArchitecture)")
    [void]$output.AppendLine("")
    
    $uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
    [void]$output.AppendLine("COMPUTER UPTIME")
    [void]$output.AppendLine("---------------")
    [void]$output.AppendLine("Days: $($uptime.Days)")
    [void]$output.AppendLine("Hours: $($uptime.Hours)")
    [void]$output.AppendLine("Minutes: $($uptime.Minutes)")
    [void]$output.AppendLine("Seconds: $($uptime.Seconds)")
    [void]$output.AppendLine("")
    
    $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $usedRAM = [math]::Round($totalRAM - $freeRAM, 2)
    $ramPercentage = [math]::Round(($usedRAM / $totalRAM) * 100, 2)
    
    [void]$output.AppendLine("RAM UTILIZATION")
    [void]$output.AppendLine("--------------")
    [void]$output.AppendLine("Total RAM: $totalRAM GB")
    [void]$output.AppendLine("Used RAM: $usedRAM GB")
    [void]$output.AppendLine("Free RAM: $freeRAM GB")
    [void]$output.AppendLine("RAM Usage: $ramPercentage%")
    [void]$output.AppendLine("")
    
    $cpuLoad = (Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
    
    [void]$output.AppendLine("CPU UTILIZATION")
    [void]$output.AppendLine("--------------")
    [void]$output.AppendLine("CPU Usage: $cpuLoad%")
    [void]$output.AppendLine("")
    
    [void]$output.AppendLine("STORAGE UTILIZATION")
    [void]$output.AppendLine("------------------")
    
    Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
        $driveLetter = $_.DeviceID
        $totalSpace = [math]::Round($_.Size / 1GB, 2)
        $freeSpace = [math]::Round($_.FreeSpace / 1GB, 2)
        $usedSpace = [math]::Round($totalSpace - $freeSpace, 2)
        $usedPercentage = [math]::Round(($usedSpace / $totalSpace) * 100, 2)
        
        [void]$output.AppendLine("Drive $driveLetter")
        [void]$output.AppendLine("  Total Space: $totalSpace GB")
        [void]$output.AppendLine("  Used Space: $usedSpace GB")
        [void]$output.AppendLine("  Free Space: $freeSpace GB")
        [void]$output.AppendLine("  Usage: $usedPercentage%")
        [void]$output.AppendLine("")
    }
    
    $output.ToString() | Out-File -FilePath $OutputFile
    
    return @{
        RAMUsage = $ramPercentage
        CPUUsage = $cpuLoad
        Uptime   = $uptime
        DiskInfo = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
            @{
                Drive           = $_.DeviceID
                UsagePercentage = [math]::Round((($_.Size - $_.FreeSpace) / $_.Size) * 100, 2)
                FreeSpace       = [math]::Round($_.FreeSpace / 1GB, 2)
            }
        }
    }
}



function Take-Screenshot {
    param (
        [string]$FilePath
    )
    
    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
    $bmp.Save($FilePath)
    $graphics.Dispose()
    $bmp.Dispose()
}

function Start-SpeedTest {
    param (
        [string]$OutputFile
    )
    
    Write-Host "Starting Ookla Speedtest..." -ForegroundColor Blue
    
    # Check if speedtest CLI is installed
    $speedtestPath = "C:\Program Files\Speedtest\speedtest.exe"
    $altSpeedtestPath = "C:\Program Files (x86)\Speedtest\speedtest.exe"
    
    if (Test-Path -Path $speedtestPath) {
        $speedtestExe = $speedtestPath
    } 
    elseif (Test-Path -Path $altSpeedtestPath) {
        $speedtestExe = $altSpeedtestPath
    }
    else {
        # If not found, try to find it in PATH
        $speedtestExe = Get-Command speedtest.exe -ErrorAction SilentlyContinue
        
        if ($null -eq $speedtestExe) {
            Write-Host "Speedtest CLI not found. Attempting to download..." -ForegroundColor Red
            
            # URL for Speedtest CLI
            $speedtestUrl = "https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip"
            $tempZipFile = "$env:TEMP\speedtest.zip"
            $tempExtractPath = "$env:TEMP\speedtest"
            
            try {
                # Download the Speedtest CLI zip file
                Invoke-WebRequest -Uri $speedtestUrl -OutFile $tempZipFile
                
                # Create directory if it doesn't exist
                if (-not (Test-Path -Path $tempExtractPath)) {
                    New-Item -ItemType Directory -Path $tempExtractPath | Out-Null
                }
                
                # Extract the zip file
                Expand-Archive -Path $tempZipFile -DestinationPath $tempExtractPath -Force
                
                # Set the path to the extracted speedtest.exe
                $speedtestExe = "$tempExtractPath\speedtest.exe"
                
                Write-Host "Speedtest CLI downloaded successfully." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to download Speedtest CLI: $_" -ForegroundColor Red
                Write-Host "Skipping speedtest..." -ForegroundColor Yellow
                return $null
            }
        }
        else {
            $speedtestExe = $speedtestExe.Source
        }
    }
    
    # Run speedtest with progress command
    try {
        Write-Host "Running speedtest. This may take a minute..." -ForegroundColor Blue
        
        # Run the speedtest and capture output
        $speedtestOutput = & $speedtestExe --accept-license --accept-gdpr --format=json
        
        # Convert JSON output to PowerShell object
        $speedtestResult = $speedtestOutput | ConvertFrom-Json
        
        # Format the results
        $speedtestReport = New-Object System.Text.StringBuilder
        
        [void]$speedtestReport.AppendLine("*****************************************************************************")
        [void]$speedtestReport.AppendLine("* SPEEDTEST RESULTS *")
        [void]$speedtestReport.AppendLine("*****************************************************************************")
        [void]$speedtestReport.AppendLine("")
        [void]$speedtestReport.AppendLine("TEST PERFORMED ON: $(Get-Date)")
        [void]$speedtestReport.AppendLine("")
        [void]$speedtestReport.AppendLine("SERVER INFORMATION")
        [void]$speedtestReport.AppendLine("-----------------")
        [void]$speedtestReport.AppendLine("Host: $($speedtestResult.server.host)")
        [void]$speedtestReport.AppendLine("Name: $($speedtestResult.server.name)")
        [void]$speedtestReport.AppendLine("Location: $($speedtestResult.server.location), $($speedtestResult.server.country)")
        [void]$speedtestReport.AppendLine("")
        [void]$speedtestReport.AppendLine("CONNECTION INFORMATION")
        [void]$speedtestReport.AppendLine("---------------------")
        [void]$speedtestReport.AppendLine("ISP: $($speedtestResult.isp)")
        [void]$speedtestReport.AppendLine("Internal IP: $($speedtestResult.interface.internalIp)")
        [void]$speedtestReport.AppendLine("External IP: $($speedtestResult.interface.externalIp)")
        [void]$speedtestReport.AppendLine("")
        [void]$speedtestReport.AppendLine("SPEED RESULTS")
        [void]$speedtestReport.AppendLine("-------------")
        [void]$speedtestReport.AppendLine("Download: $([math]::Round($speedtestResult.download.bandwidth * 8 / 1000000, 2)) Mbps")
        [void]$speedtestReport.AppendLine("Upload: $([math]::Round($speedtestResult.upload.bandwidth * 8 / 1000000, 2)) Mbps")
        [void]$speedtestReport.AppendLine("Ping: $($speedtestResult.ping.latency) ms")
        [void]$speedtestReport.AppendLine("Jitter: $($speedtestResult.ping.jitter) ms")
        [void]$speedtestReport.AppendLine("")
        [void]$speedtestReport.AppendLine("PACKET LOSS")
        [void]$speedtestReport.AppendLine("-----------")
        [void]$speedtestReport.AppendLine("Packet Loss: $($speedtestResult.packetLoss)%")
        [void]$speedtestReport.AppendLine("")
        
        # Save the report to a file
        $speedtestReport.ToString() | Out-File -FilePath $OutputFile
        
        # Log results to console
        Write-Host "Speedtest completed successfully!" -ForegroundColor Green
        Write-Host "Download: $([math]::Round($speedtestResult.download.bandwidth * 8 / 1000000, 2)) Mbps" -ForegroundColor Green
        Write-Host "Upload: $([math]::Round($speedtestResult.upload.bandwidth * 8 / 1000000, 2)) Mbps" -ForegroundColor Green
        Write-Host "Ping: $($speedtestResult.ping.latency) ms" -ForegroundColor Green
        Write-Host "Results saved to: $OutputFile" -ForegroundColor Green
        
        return @{
            Download   = [math]::Round($speedtestResult.download.bandwidth * 8 / 1000000, 2)
            Upload     = [math]::Round($speedtestResult.upload.bandwidth * 8 / 1000000, 2)
            Ping       = $speedtestResult.ping.latency
            Jitter     = $speedtestResult.ping.jitter
            PacketLoss = $speedtestResult.packetLoss
        }
    }
    catch {
        Write-Host "Error running speedtest: $_" -ForegroundColor Red
        Write-Host "Saving error details to file..." -ForegroundColor Yellow
        "Error running speedtest: $_" | Out-File -FilePath $OutputFile
        return $null
    }
}

function Start-ParallelTests {
    param (
        [string]$Target,
        [string]$OutputFolder,
        [int]$TestDuration = 60
    )
    
    Write-Host "Starting tests for $Target..."
    
    $tcpingOutputFile = "$OutputFolder\${Target}_tcping.txt"
    
    if (-not (Test-Path -Path $tcpingPath)) {
        Write-Host "TCPing executable not found at: $tcpingPath. Checking in PATH..." -ForegroundColor Yellow
        
        $tcpingCommand = Get-Command tcping.exe -ErrorAction SilentlyContinue
        
        if ($null -eq $tcpingCommand) {
            Write-Host "TCPing not found in PATH. TCPing tests will be skipped." -ForegroundColor Red
            "TCPing executable not found. Test skipped." | Out-File -FilePath $tcpingOutputFile
            return $null
        }
        else {
            $tcpingPath = $tcpingCommand.Source
            Write-Host "TCPing found at: $tcpingPath" -ForegroundColor Green
        }
    }
    
    if (Test-Path -Path $tcpingPath) {
        $tcpingProcess = Start-Process -FilePath $tcpingPath -ArgumentList "-n 30 $Target 443" -RedirectStandardOutput $tcpingOutputFile -PassThru -WindowStyle Hidden
        Write-Host "TCPing started for $Target"
    }
    
    Start-Process -FilePath $winmtrPath -ArgumentList $Target -WindowStyle Maximized
    
    Write-Host "Running WinMTR for $Target for $TestDuration seconds..."
    Start-Sleep -Seconds $TestDuration
    
    Take-Screenshot -FilePath "$OutputFolder\${Target}_winmtr.png"
    
    Stop-Process -Name "winmtr" -Force -ErrorAction SilentlyContinue
    
    if ($tcpingProcess -and -not $tcpingProcess.HasExited) {
        Write-Host "Waiting for TCPing to complete for $Target..."
        $tcpingProcess.WaitForExit()
    }
    
    Write-Host "Tests completed for $Target."
    
    if (Test-Path -Path $tcpingOutputFile) {
        $tcpingContent = Get-Content -Path $tcpingOutputFile -Raw
        return @{
            Target     = $Target
            TCPingData = $tcpingContent
        }
    }
    
    return $null
}

function Parse-TCPingResults {
    param (
        [string]$TCPingData
    )
    
    $pattern = "time=(\d+\.\d+)ms"
    $matches = [regex]::Matches($TCPingData, $pattern)
    
    if ($matches.Count -eq 0) {
        return @{
            AverageTime = 0
            MinTime     = 0
            MaxTime     = 0
            TimedOut    = $true
        }
    }
    
    $times = $matches | ForEach-Object { [double]$_.Groups[1].Value }
    
    return @{
        AverageTime = ($times | Measure-Object -Average).Average
        MinTime     = ($times | Measure-Object -Minimum).Minimum
        MaxTime     = ($times | Measure-Object -Maximum).Maximum
        TimedOut    = $false
    }
}

function Create-SummaryReport {
    param (
        [hashtable]$SystemInfo,
        [hashtable]$SpeedTestResults,
        [array]$TCPingResults,
        [string]$OutputFile
    )
    
    $summary = New-Object System.Text.StringBuilder
    
    [void]$summary.AppendLine("*****************************************************************************")
    [void]$summary.AppendLine("* SUMMARY REPORT *")
    [void]$summary.AppendLine("*****************************************************************************")
    [void]$summary.AppendLine("")
    
    [void]$summary.AppendLine("SYSTEM HEALTH CHECK")
    [void]$summary.AppendLine("------------------")
    
    $systemIssues = @()
    
    if ($SystemInfo.RAMUsage -gt 80) {
        $systemIssues += "WARNING: High RAM usage ($($SystemInfo.RAMUsage)%) detected."
    }
    
    if ($SystemInfo.CPUUsage -gt 80) {
        $systemIssues += "WARNING: High CPU usage ($($SystemInfo.CPUUsage)%) detected."
    }
    
    foreach ($disk in $SystemInfo.DiskInfo) {
        if ($disk.UsagePercentage -gt 85) {
            $systemIssues += "WARNING: Low disk space on drive $($disk.Drive) (only $($disk.FreeSpace) GB free)."
        }
    }
    
    if ($SystemInfo.Uptime.Days -gt 7) {
        $systemIssues += "NOTE: System uptime is $($SystemInfo.Uptime.Days) days. Consider rebooting for optimal performance."
    }
    
    if ($systemIssues.Count -eq 0) {
        [void]$summary.AppendLine("✓ No system issues detected.")
    }
    else {
        foreach ($issue in $systemIssues) {
            [void]$summary.AppendLine("! $issue")
        }
    }
    
    [void]$summary.AppendLine("")
    [void]$summary.AppendLine("INTERNET CONNECTION CHECK")
    [void]$summary.AppendLine("------------------------")
    
    $connectionIssues = @()
    
    if ($null -eq $SpeedTestResults) {
        $connectionIssues += "ERROR: Could not perform speed test."
    }
    else {
        if ($SpeedTestResults.Download -lt 10) {
            $connectionIssues += "WARNING: Slow download speed ($($SpeedTestResults.Download) Mbps)."
        }
        
        if ($SpeedTestResults.Upload -lt 5) {
            $connectionIssues += "WARNING: Slow upload speed ($($SpeedTestResults.Upload) Mbps)."
        }
        
        if ($SpeedTestResults.Ping -gt 100) {
            $connectionIssues += "WARNING: High ping latency ($($SpeedTestResults.Ping) ms)."
        }
        
        if ($SpeedTestResults.Jitter -gt 30) {
            $connectionIssues += "WARNING: High jitter ($($SpeedTestResults.Jitter) ms)."
        }
        
        if ($SpeedTestResults.PacketLoss -gt 2) {
            $connectionIssues += "WARNING: Packet loss detected ($($SpeedTestResults.PacketLoss)%)."
        }
    }
    
    if ($connectionIssues.Count -eq 0) {
        [void]$summary.AppendLine("✓ Internet connection appears healthy.")
        [void]$summary.AppendLine("  Download: $($SpeedTestResults.Download) Mbps")
        [void]$summary.AppendLine("  Upload: $($SpeedTestResults.Upload) Mbps")
        [void]$summary.AppendLine("  Ping: $($SpeedTestResults.Ping) ms")
    }
    else {
        foreach ($issue in $connectionIssues) {
            [void]$summary.AppendLine("! $issue")
        }
    }
    
    [void]$summary.AppendLine("")
    [void]$summary.AppendLine("SERVER CONNECTIVITY CHECK")
    [void]$summary.AppendLine("-----------------------")
    
    if ($TCPingResults.Count -eq 0) {
        [void]$summary.AppendLine("! No TCPing results available.")
    }
    else {
        $parsedResults = @()
        
        foreach ($result in $TCPingResults) {
            if ($null -ne $result) {
                $parsedData = Parse-TCPingResults -TCPingData $result.TCPingData
                $parsedResults += [PSCustomObject]@{
                    Target      = $result.Target
                    AverageTime = $parsedData.AverageTime
                    MinTime     = $parsedData.MinTime
                    MaxTime     = $parsedData.MaxTime
                    TimedOut    = $parsedData.TimedOut
                }
            }
        }
        
        if ($parsedResults.Count -gt 0) {
            $bestServer = $parsedResults | Where-Object { -not $_.TimedOut } | Sort-Object AverageTime | Select-Object -First 1
            
            if ($null -ne $bestServer) {
                [void]$summary.AppendLine("BEST PERFORMING SERVER: $($bestServer.Target)")
                [void]$summary.AppendLine("  Average ping time: $([math]::Round($bestServer.AverageTime, 2)) ms")
                [void]$summary.AppendLine("  Min/Max ping time: $([math]::Round($bestServer.MinTime, 2)) / $([math]::Round($bestServer.MaxTime, 2)) ms")
                [void]$summary.AppendLine("")
            }
            
            [void]$summary.AppendLine("ALL SERVER RESULTS:")
            foreach ($server in $parsedResults) {
                if ($server.TimedOut) {
                    [void]$summary.AppendLine("! $($server.Target): Connection timed out or failed")
                }
                else {
                    [void]$summary.AppendLine("- $($server.Target): Avg=$([math]::Round($server.AverageTime, 2)) ms, Min=$([math]::Round($server.MinTime, 2)) ms, Max=$([math]::Round($server.MaxTime, 2)) ms")
                }
            }
        }
        else {
            [void]$summary.AppendLine("! No valid TCPing results found.")
        }
    }
    
    $summary.ToString() | Out-File -FilePath $OutputFile
}

function Start-Analysis {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $DATESTAMP = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

    $baseFolder = "$env:USERPROFILE\Downloads\Ace Network Result"
    $outputFolder = "$baseFolder\$DATESTAMP"
    
    if (-not (Test-Path -Path $baseFolder)) {
        New-Item -ItemType Directory -Path $baseFolder | Out-Null
    }
    
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
    
    if (-not (Test-Path -Path $winmtrPath)) {
        Write-Host "WinMTR executable not found at: $winmtrPath" -ForegroundColor Red
        $winmtrCommand = Get-Command winmtr.exe -ErrorAction SilentlyContinue
        if ($null -eq $winmtrCommand) {
            Write-Host "ERROR: WinMTR not found. Please make sure winmtr.exe is in the script folder or PATH." -ForegroundColor Red
            return
        }
        else {
            $winmtrPath = $winmtrCommand.Source
            Write-Host "WinMTR found at: $winmtrPath" -ForegroundColor Green
        }
    }
    
    Write-Host "Capturing initial system information..."
    $systemInfo = Get-SystemInfo -OutputFile "$outputFolder\SystemInfo.txt"
    
    Write-Host "Performing speed test..."
    $speedTestResults = Start-SpeedTest -OutputFile "$outputFolder\speedtest_results.txt"
    
    $targets = @(
        "RDGCHG.myrealdata.net",
        "RDGHTN.myrealdata.net",
        "RDGNV.myrealdata.net",
        "RDGATL.myrealdata.net"
    )
    
    $tcpingResults = @()
    
    foreach ($target in $targets) {
        $result = Start-ParallelTests -Target $target -OutputFolder $outputFolder -TestDuration 60
        if ($null -ne $result) {
            $tcpingResults += $result
        }
    }
    
    Write-Host "Creating summary report..."
    Create-SummaryReport -SystemInfo $systemInfo -SpeedTestResults $speedTestResults -TCPingResults $tcpingResults -OutputFile "$outputFolder\summary_report.txt"
    
    Remove-Item -Path "$env:TEMP\speedtest" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:TEMP\speedtest.zip" -Force -ErrorAction SilentlyContinue

    Write-Host "Analysis complete. Results saved to: $outputFolder"
    Write-Host "Summary report available at: $outputFolder\summary_report.txt"
}

Start-Analysis