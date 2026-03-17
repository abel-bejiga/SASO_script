# ==========================================
# RELOAD Script - SASO DB, Bookings, Registry RELOADER
# Abel B.
# ==========================================

# Check for Administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'." -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "Press any key to exit..."
    [void][System.Console]::ReadKey($true)
    exit 1
}

$ErrorActionPreference = "Stop"

try {
    Write-Host "Starting RELOAD process..." -ForegroundColor Cyan

    # -------------------------------
    # Define backup source
    # -------------------------------
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $backupSource = Join-Path $desktopPath "restore"

    if (!(Test-Path $backupSource)) {
        throw "Backup path does not exist: $backupSource"
    }

    # Paths
    $bookingsSource = Join-Path $backupSource "Bookings_Last14Days"
    $dbSource = Join-Path $backupSource "Databases"

    $bookingsDest = "D:\Printrak\Bookings"
    $dbDest = "C:\ProgramData\Morphotrak\Printrak LiveScan\Databases"

    # -------------------------------
    # B. Clean Bookings (except wip)
    # -------------------------------
    try {
        if (!(Test-Path $bookingsDest)) {
            throw "Bookings destination not found: $bookingsDest"
        }

        Write-Host "Cleaning Bookings folder (excluding 'wip')..." -ForegroundColor Yellow

        $itemsToRemove = @(Get-ChildItem $bookingsDest -Force | Where-Object { $_.Name -ne "wip" })
        $totalClean = $itemsToRemove.Count
        $currentClean = 0

        if ($totalClean -gt 0) {
            foreach ($item in $itemsToRemove) {
                $currentClean++
                $percent = [math]::Round(($currentClean / $totalClean) * 100)
                Write-Progress -Activity "Cleaning Bookings (except wip)" -Status "Progress: $percent%" -PercentComplete $percent
                Remove-Item $item.FullName -Recurse -Force
            }
            Write-Progress -Activity "Cleaning Bookings (except wip)" -Completed
        }

        Write-Host "Old bookings removed (except wip)." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR cleaning bookings: $_" -ForegroundColor Red
    }

    # -------------------------------
    # Restore Bookings (last 14 days)
    # -------------------------------
    try {
        if (!(Test-Path $bookingsSource)) {
            throw "Backup bookings not found: $bookingsSource"
        }

        Write-Host "Restoring Bookings..." -ForegroundColor Yellow

        $itemsToCopy = @(Get-ChildItem -Path $bookingsSource -Recurse)
        $totalCopy = $itemsToCopy.Count
        $currentCopy = 0

        if ($totalCopy -gt 0) {
            foreach ($item in $itemsToCopy) {
                $currentCopy++
                $percent = [math]::Round(($currentCopy / $totalCopy) * 100)
                Write-Progress -Activity "Restoring Bookings" -Status "Progress: $percent%" -PercentComplete $percent

                $destFile = $item.FullName.Replace($bookingsSource, $bookingsDest)
                if ($item.PSIsContainer) {
                    if (!(Test-Path $destFile)) {
                        New-Item -ItemType Directory -Path $destFile -Force | Out-Null
                    }
                }
                else {
                    $destDir = Split-Path $destFile
                    if (!(Test-Path $destDir)) {
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                    }
                    Copy-Item -Path $item.FullName -Destination $destFile -Force
                }
            }
            Write-Progress -Activity "Restoring Bookings" -Completed
        }

        Write-Host "Bookings restored successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR restoring bookings: $_" -ForegroundColor Red
    }

    # -------------------------------
    # C. Restore Databases
    # -------------------------------
    try {
        if (!(Test-Path $dbSource)) {
            throw "Backup databases not found: $dbSource"
        }

        Write-Host "Restoring Databases..." -ForegroundColor Yellow

        Copy-Item -Path "$dbSource\*" -Destination $dbDest -Recurse -Force

        Write-Host "Databases restored successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR restoring databases: $_" -ForegroundColor Red
    }

    # -------------------------------
    # D. Restore Printrak Registry
    # -------------------------------

    try {
        $regFile = Join-Path $restoreSource "Printrak.reg"

        if (!(Test-Path $regFile)) {
            throw "Registry backup file not found: $regFile"
        }

        Write-Host "Importing Printrak registry..." -ForegroundColor Yellow

        reg import $regFile

        if ($LASTEXITCODE -eq 0) {
            Write-Host "Registry imported successfully." -ForegroundColor Green
        }
        else {
            throw "Registry import failed with exit code $LASTEXITCODE"
        }
    }
    catch {
        Write-Host "ERROR importing registry: $_" -ForegroundColor Red
    }
    
    # -------------------------------
    # E. Rename Computer
    # -------------------------------
    try {
        $newName = Read-Host "Enter NEW computer name"

        if ([string]::IsNullOrWhiteSpace($newName)) {
            throw "Computer name cannot be empty."
        }

        Rename-Computer -NewName $newName -Force -ErrorAction Stop

        Write-Host "Computer renamed successfully to: $newName" -ForegroundColor Green
        Write-Host "⚠️ Restart required to apply name change." -ForegroundColor Yellow
    }
    catch {
        Write-Host "ERROR renaming computer: $_" -ForegroundColor Red
    }

    Write-Host "RELOAD process completed." -ForegroundColor Cyan
}
catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
}

# Pause at the end
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Yellow
[void][System.Console]::ReadKey($true)