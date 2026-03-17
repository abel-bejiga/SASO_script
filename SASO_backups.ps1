# ================================
# Backup Script - saso_backup
# Abel B.
# ================================

# 🔒 Check for Administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'." -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    exit 1
}

$ErrorActionPreference = "Stop"

try {
    Write-Host "Starting backup process..." -ForegroundColor Cyan

    # Get computer name
    $computerName = $env:COMPUTERNAME

    # Define base backup path
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $backupRoot = Join-Path $desktopPath "backups"
    $backupPath = Join-Path $backupRoot $computerName

    # Create backup directories if not exist
    if (!(Test-Path $backupRoot)) {
        New-Item -ItemType Directory -Path $backupRoot | Out-Null
        Write-Host "Created folder: $backupRoot" -ForegroundColor Green
    }

    if (!(Test-Path $backupPath)) {
        New-Item -ItemType Directory -Path $backupPath | Out-Null
        Write-Host "Created folder: $backupPath" -ForegroundColor Green
    }

    # -------------------------------
    # A. Export Registry Key
    # -------------------------------
    try {
        $regPath = "HKLM\SOFTWARE\WOW6432Node\Printrak"
        $regExportFile = Join-Path $backupPath "Printrak.reg"

        reg export $regPath $regExportFile /y | Out-Null

        if (Test-Path $regExportFile) {
            Write-Host "Registry export successful: $regExportFile" -ForegroundColor Green
        } else {
            throw "Registry export failed."
        }
    } catch {
        Write-Host "ERROR exporting registry: $_" -ForegroundColor Red
    }

    # -------------------------------
    # B. Copy Bookings (Last 14 Days)
    # -------------------------------
    try {
        $sourceBookings = "D:\Printrak\Bookings"
        $destBookings = Join-Path $backupPath "Bookings_Last14Days"

        if (!(Test-Path $sourceBookings)) {
            throw "Source path not found: $sourceBookings"
        }

        New-Item -ItemType Directory -Path $destBookings -Force | Out-Null

        $cutoffDate = (Get-Date).AddDays(-14)

        $itemsToCopy = @(Get-ChildItem -Path $sourceBookings -Recurse | Where-Object { $_.LastWriteTime -ge $cutoffDate })
        $totalCopy = $itemsToCopy.Count
        $currentCopy = 0

        if ($totalCopy -gt 0) {
            foreach ($item in $itemsToCopy) {
                $currentCopy++
                $percent = [math]::Round(($currentCopy / $totalCopy) * 100)
                Write-Progress -Activity "Backing up Bookings (last 14 days)" -Status "Progress: $percent%" -PercentComplete $percent

                $destFile = $item.FullName.Replace($sourceBookings, $destBookings)
                $destDir = Split-Path $destFile

                if (!(Test-Path $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }

                Copy-Item $item.FullName -Destination $destFile -Force
            }
            Write-Progress -Activity "Backing up Bookings (last 14 days)" -Completed
        }

        Write-Host "Bookings (last 14 days) copied successfully." -ForegroundColor Green
    } catch {
        Write-Host "ERROR copying bookings: $_" -ForegroundColor Red
    }

    # -------------------------------
    # C. Copy Databases Folder
    # -------------------------------
    try {
        $sourceDB = "C:\ProgramData\Morphotrak\Printrak LiveScan\Databases"
        $destDB = Join-Path $backupPath "Databases"

        if (!(Test-Path $sourceDB)) {
            throw "Source path not found: $sourceDB"
        }

        Copy-Item -Path $sourceDB -Destination $destDB -Recurse -Force

        Write-Host "Databases folder copied successfully." -ForegroundColor Green
    } catch {
        Write-Host "ERROR copying databases: $_" -ForegroundColor Red
    }

    Write-Host "Backup process completed." -ForegroundColor Cyan
}
catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
}

# Pause at the end
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Yellow
[void][System.Console]::ReadKey($true)