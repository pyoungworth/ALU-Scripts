<#
.SYNOPSIS
    Sets up Awesome Sauce v2 auto-boot on a USB drive for the ALU.

.DESCRIPTION
    Copies the autostart scripts and files to the USB drive so the ALU
    will automatically boot into Awesome Sauce v2 when powered on:
      1. Copies 00_install_autostart.sh into <USB>\OneSauce\scripter\
      2. Extracts autostart.zip to the root of the USB drive
      3. Copies 00_init_menu.sh into the <USB>\autostart\ folder

    Can be run standalone (will prompt for the USB drive) or called from
    Extract-To-USB.ps1 with the -UsbPath parameter.

.PARAMETER UsbPath
    The root path of the USB drive where Awesome Sauce v2 has been extracted.
    If not provided, the script will guide you through selecting it.

.EXAMPLE
    .\Install-AutoBoot.ps1
    .\Install-AutoBoot.ps1 -UsbPath "E:\"

.NOTES
    Author: Philip Youngworth
    Project: ALU-Scripts (AwesomeSauce2 Utility Scripts)
    License: MIT
#>

param(
    [Parameter(Position = 0)]
    [string]$UsbPath
)

# -- Helpers ----------------------------------------------------------------------

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Write-Header {
    param([string]$Text)
    $line = "=" * 80
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Write-OK {
    param([string]$Text)
    Write-Host "  [OK] $Text" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Text)
    Write-Host "  [!!] $Text" -ForegroundColor Red
}

function Write-Info {
    param([string]$Text)
    Write-Host "  $Text" -ForegroundColor Gray
}

function Test-Quit {
    param([string]$Response)
    if ($Response -and ($Response.Trim().ToLower() -eq 'q' -or $Response.Trim().ToLower() -eq 'quit')) {
        Write-Host ""
        Write-Host "  Exiting. Run this script again whenever you're ready!" -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }
}

# Load zip support
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# =================================================================================
#  Resolve required files bundled with this script
# =================================================================================

$installAutoStartFile = Join-Path $scriptDir '00_install_autostart.sh'
$autoStartZipFile     = Join-Path $scriptDir 'autostart.zip'
$initMenuFile         = Join-Path $scriptDir '00_init_menu.sh'

$missingFiles = @()
if (-not (Test-Path $installAutoStartFile)) { $missingFiles += '00_install_autostart.sh' }
if (-not (Test-Path $autoStartZipFile))     { $missingFiles += 'autostart.zip' }
if (-not (Test-Path $initMenuFile))         { $missingFiles += '00_init_menu.sh' }

if ($missingFiles.Count -gt 0) {
    Write-Host ""
    Write-Fail "Missing required file(s) in the script directory ($scriptDir):"
    foreach ($f in $missingFiles) {
        Write-Host "    - $f" -ForegroundColor Red
    }
    Write-Host ""
    exit 1
}

# =================================================================================
#  STEP 1: WHERE IS THE USB DRIVE?
# =================================================================================

Write-Header "INSTALL AUTO-BOOT FOR AWESOME SAUCE v2"

if ($UsbPath) {
    if (-not (Test-Path $UsbPath)) {
        Write-Fail "Path not found: $UsbPath"
        exit 1
    }
    $UsbPath = (Resolve-Path $UsbPath).Path
}
else {
    Write-Host ""
    Write-Host "  Select the drive where Awesome Sauce v2 has been extracted:" -ForegroundColor White
    Write-Host ""

    # Build drive list
    $partitions = Get-Partition -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
    $disks = Get-Disk -ErrorAction SilentlyContinue

    $diskInfoByLetter = @{}
    foreach ($p in $partitions) {
        $dl = [string]$p.DriveLetter
        if (-not $dl -or $dl -eq "`0") { continue }
        $dk = $disks | Where-Object { $_.Number -eq $p.DiskNumber }
        if ($dk) {
            $diskInfoByLetter[$dl] = [PSCustomObject]@{
                BusType = $dk.BusType
                Model   = $dk.FriendlyName
            }
        }
    }

    $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
              Where-Object { $_.Free -gt 0 -and $_.Root } |
              Sort-Object Name

    $idx = 0
    $driveList = @()

    foreach ($drv in $drives) {
        $idx++
        $letter = $drv.Name
        $root   = $drv.Root

        $vol = Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue
        $drvLabel = if ($vol -and $vol.FileSystemLabel) { $vol.FileSystemLabel } else { "" }
        $freeSpace = Format-FileSize $drv.Free
        $totalSpace = Format-FileSize ($drv.Used + $drv.Free)

        $busTag = ""
        if ($diskInfoByLetter.ContainsKey($letter)) {
            $di = $diskInfoByLetter[$letter]
            switch ($di.BusType) {
                'USB'   { $busTag = "USB" }
                'NVMe'  { $busTag = "NVMe" }
                'SATA'  { $busTag = "SATA" }
                'SAS'   { $busTag = "SAS" }
                default { if ($di.BusType) { $busTag = $di.BusType } }
            }
        }

        $label = "${letter}:\"
        $detailStr = ""
        if ($drvLabel) { $detailStr += "  $drvLabel" }
        if ($busTag)   { $detailStr += "  [$busTag]" }

        # Check for OneSauce directory as a hint
        $hasOneSauce = Test-Path (Join-Path $root 'OneSauce')
        $hintStr = if ($hasOneSauce) { "  <-- OneSauce found" } else { "" }

        Write-Host "    [$idx]  $label$detailStr  ($freeSpace free of $totalSpace)" -NoNewline -ForegroundColor White
        if ($hasOneSauce) {
            Write-Host $hintStr -ForegroundColor Green
        }
        else {
            Write-Host ""
        }

        $driveList += [PSCustomObject]@{
            Index  = $idx
            Letter = $letter
            Root   = $root
        }
    }

    Write-Host ""
    $selectedDrive = $null
    while (-not $selectedDrive) {
        $driveChoice = Read-Host "  Pick a drive (enter number, q to quit)"
        Test-Quit $driveChoice

        if ($driveChoice -match '^\d+$') {
            $chosen = $driveList | Where-Object { $_.Index -eq [int]$driveChoice }
            if ($chosen) {
                $selectedDrive = $chosen
            }
            else {
                Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow
        }
    }

    $UsbPath = $selectedDrive.Root
}

# Normalize trailing backslash
$UsbPath = $UsbPath.TrimEnd('\') + '\'

# =================================================================================
#  VALIDATE: OneSauce\scripter exists on the target
# =================================================================================

$scripterDir = Join-Path $UsbPath 'OneSauce\scripter'

if (-not (Test-Path (Join-Path $UsbPath 'OneSauce'))) {
    Write-Host ""
    Write-Fail "Could not find 'OneSauce' folder on $UsbPath"
    Write-Host "  Make sure Awesome Sauce v2 has been extracted to this drive first." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

if (-not (Test-Path $scripterDir)) {
    Write-Host ""
    Write-Info "Creating scripter directory: $scripterDir"
    New-Item -Path $scripterDir -ItemType Directory -Force | Out-Null
}

# =================================================================================
#  STEP 1: Copy 00_install_autostart.sh -> OneSauce\scripter\
# =================================================================================

Write-Host ""
Write-Host "  Installing auto-boot files..." -ForegroundColor White
Write-Host ""

$dest1 = Join-Path $scripterDir '00_install_autostart.sh'
try {
    Copy-Item -Path $installAutoStartFile -Destination $dest1 -Force
    Write-OK "Copied 00_install_autostart.sh -> OneSauce\scripter\"
}
catch {
    Write-Fail "Failed to copy 00_install_autostart.sh: $($_.Exception.Message)"
    exit 1
}

# =================================================================================
#  STEP 2: Extract autostart.zip -> USB root
# =================================================================================

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($autoStartZipFile, $UsbPath)
    Write-OK "Extracted autostart.zip -> $UsbPath"
}
catch [System.IO.IOException] {
    # Files may already exist -- overwrite by extracting entry by entry
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($autoStartZipFile)
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }
            $destFile = Join-Path $UsbPath $entry.FullName
            $destDir = Split-Path $destFile -Parent
            if (-not (Test-Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }
            $entryStream = $entry.Open()
            $fileStream = [System.IO.File]::Create($destFile)
            try {
                $entryStream.CopyTo($fileStream)
            }
            finally {
                $fileStream.Close()
                $entryStream.Close()
            }
        }
        $zip.Dispose()
        Write-OK "Extracted autostart.zip -> $UsbPath (overwrote existing files)"
    }
    catch {
        Write-Fail "Failed to extract autostart.zip: $($_.Exception.Message)"
        if ($zip) { $zip.Dispose() }
        exit 1
    }
}
catch {
    Write-Fail "Failed to extract autostart.zip: $($_.Exception.Message)"
    exit 1
}

# =================================================================================
#  STEP 3: Copy 00_init_menu.sh -> autostart\
# =================================================================================

$autostartDir = Join-Path $UsbPath 'autostart'

if (-not (Test-Path $autostartDir)) {
    Write-Fail "autostart folder was not created by the zip extraction. Something went wrong."
    exit 1
}

$dest3 = Join-Path $autostartDir '00_init_menu.sh'
try {
    Copy-Item -Path $initMenuFile -Destination $dest3 -Force
    Write-OK "Copied 00_init_menu.sh -> autostart\"
}
catch {
    Write-Fail "Failed to copy 00_init_menu.sh: $($_.Exception.Message)"
    exit 1
}

# =================================================================================
#  DONE
# =================================================================================

Write-Header "AUTO-BOOT INSTALLED"
Write-Host ""
Write-Host "  Your USB drive is now set up for auto-boot!" -ForegroundColor Green
Write-Host ""
Write-Host "  What was installed:" -ForegroundColor White
Write-Host "    - OneSauce\scripter\00_install_autostart.sh" -ForegroundColor Gray
Write-Host "    - autostart\ (from autostart.zip)" -ForegroundColor Gray
Write-Host "    - autostart\00_init_menu.sh" -ForegroundColor Gray
Write-Host ""
Write-Host "  You may need to start OneSauce manually the first time." -ForegroundColor White
Write-Host "  After that, your ALU will automatically boot into" -ForegroundColor White
Write-Host "  Awesome Sauce v2 whenever the USB drive is inserted." -ForegroundColor White
Write-Host ""
