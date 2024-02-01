# Created by Danny Ware
# 07/02/2023
# Update 02/01/2024
# Automated Reporting Script

# Configuration
$reportFolderPath = "C:\Reports"

# Create the report folder if it doesn't exist
if (-not (Test-Path $reportFolderPath)) {
    New-Item -ItemType Directory -Path $reportFolderPath | Out-Null
}

# Get system information
$computerName = $env:COMPUTERNAME
$operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem
$processor = Get-CimInstance -ClassName Win32_Processor
$memoryModules = Get-CimInstance -ClassName Win32_PhysicalMemory
$networkAdapter = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
$uptime = (Get-Date) - (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
$uptimeFormatted = "{0} Days, {1} Hours, {2} Minutes, {3} Seconds" -f $uptime.Days, $uptime.Hours, $uptime.Minutes, $uptime.Seconds

# Calculate the total RAM capacity
$totalRAMBytes = ($memoryModules | Measure-Object -Property Capacity -Sum).Sum
$totalRAMGB = "{0:N2}" -f ($totalRAMBytes / 1GB)

# Get RAM usage
$usedRAMBytes = (Get-CimInstance -ClassName Win32_OperatingSystem).TotalVisibleMemorySize
$usedRAMGB = "{0:N2}" -f ($usedRAMBytes / 1GB)

# Get hard drive information
$hardDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, VolumeName, Size, FreeSpace
$driveInfo = $hardDrive | ForEach-Object {
    $driveLetter = $_.DeviceID
    $volumeName = $_.VolumeName
    $driveSizeGB = "{0:N2}" -f ($_.Size / 1GB)
    $freeSpaceGB = "{0:N2}" -f ($_.FreeSpace / 1GB)
    $usedSpaceGB = "{0:N2}" -f (($_.Size - $_.FreeSpace) / 1GB)

    [PSCustomObject]@{
        DriveLetter = $driveLetter
        VolumeName = $volumeName
        DriveSizeGB = $driveSizeGB
        FreeSpaceGB = $freeSpaceGB
        UsedSpaceGB = $usedSpaceGB
    }
}

# Get pending Windows updates
$pendingUpdates = Get-WmiObject -Query "SELECT * FROM Win32_QuickFixEngineering WHERE HotFixID IS NOT NULL"

# Build the report content
$reportContent = @"
System Performance Report
------------------------

Computer Name        : $computerName
Operating System     : $($operatingSystem.Caption)
Processor            : $($processor.Name)
Total RAM (GB)       : $totalRAMGB
Used RAM (GB)        : $usedRAMGB
Network Adapter      : $($networkAdapter.Description)
IP Address           : $($networkAdapter.IPAddress[0])
Uptime               : $uptimeFormatted

Hard Drive Information:
-----------------------
$($driveInfo | Format-Table -AutoSize | Out-String)

Pending Updates:
----------------
$($pendingUpdates | Select-Object -ExpandProperty HotFixID -Unique | Out-String)
"@

# Generate a timestamp for the report file name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Generate the report file path
$reportFileName = "SystemPerformanceReport_$timestamp.txt"
$reportFilePath = Join-Path -Path $reportFolderPath -ChildPath $reportFileName

# Save the report to a text file
$reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8

Write-Host "System performance report generated and saved to: $reportFilePath"
