<#
.SYNOPSIS
    Downloads files from Google Drive using rclone -- fully guided for non-technical users.

.DESCRIPTION
    This script walks you through everything:
      1. Installs rclone if you don't have it
      2. Sets up the Google Drive connection (opens browser for sign-in)
      3. Lets you pick where to download to
      4. Optionally limits download speed so you can still use the internet
      5. Downloads everything (can be resumed if interrupted)

    Just run the script and follow the prompts!

.EXAMPLE
    .\Download-From-GDrive.ps1

.NOTES
    Author: Philip Youngworth
    Project: ALU-Scripts (AwesomeSauce2 Utility Scripts)
    License: MIT
#>

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

function Write-Step {
    param([string]$Text)
    Write-Host ""
    Write-Host "  >> $Text" -ForegroundColor White
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

function Refresh-PathInSession {
    # Reload PATH from the registry so we pick up newly installed programs
    # without needing to close and reopen the terminal
    $machinePath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine')
    $userPath = [System.Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machinePath;$userPath"
}

function Find-Rclone {
    # Try to find rclone in PATH
    $cmd = Get-Command rclone -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    # Check common install locations
    $commonPaths = @(
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links\rclone.exe",
        "$env:ProgramFiles\rclone\rclone.exe",
        "${env:ProgramFiles(x86)}\rclone\rclone.exe",
        "$env:USERPROFILE\scoop\shims\rclone.exe"
    )
    foreach ($p in $commonPaths) {
        if (Test-Path $p) { return $p }
    }

    # Search winget packages folder
    $wingetPkgs = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages"
    if (Test-Path $wingetPkgs) {
        $found = Get-ChildItem -Path $wingetPkgs -Filter 'rclone.exe' -Recurse -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($found) { return $found.FullName }
    }

    return $null
}

$remoteName = "RetroFE"

# =================================================================================
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  GOOGLE DRIVE DOWNLOAD TOOL" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  This script will help you download all the files from" -ForegroundColor White
Write-Host "  Google Drive. Just follow the prompts!" -ForegroundColor White
Write-Host ""
Write-Host "  What this will do:" -ForegroundColor Gray
Write-Host "    1. Install rclone (a download tool) if needed" -ForegroundColor Gray
Write-Host "    2. Connect to Google Drive (you'll sign in via browser)" -ForegroundColor Gray
Write-Host "    3. Pick where to save the files" -ForegroundColor Gray
Write-Host "    4. Download everything" -ForegroundColor Gray
Write-Host ""
Write-Host "  If the download gets interrupted, just run this script" -ForegroundColor Yellow
Write-Host "  again -- it will pick up where it left off!" -ForegroundColor Yellow
Write-Host ""

$ready = Read-Host "  Ready to start? (y/n, q to quit)"
Test-Quit $ready
if ($ready.ToLower() -ne 'y') {
    Write-Host "  No problem. Run this script whenever you're ready!" -ForegroundColor Gray
    exit 0
}

# =================================================================================
#  STEP 1: INSTALL RCLONE
# =================================================================================

Write-Header "STEP 1: CHECKING FOR RCLONE"

$rclonePath = Find-Rclone

if ($rclonePath) {
    Write-OK "rclone is already installed"
    Write-Info "Location: $rclonePath"
    try {
        $verOutput = & $rclonePath version 2>&1 | Select-Object -First 1
        Write-Info "Version:  $verOutput"
    } catch {}
}
else {
    Write-Step "rclone is not installed. Installing it now..."
    Write-Host ""

    # Check if winget is available
    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetCmd) {
        Write-Fail "winget is not available on this computer."
        Write-Host ""
        Write-Host "  winget comes with Windows 10 (version 1809+) and Windows 11." -ForegroundColor Yellow
        Write-Host "  You may need to update the 'App Installer' from the Microsoft Store." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Alternatively, you can install rclone manually from:" -ForegroundColor Yellow
        Write-Host "  https://rclone.org/downloads/" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  After installing rclone, run this script again." -ForegroundColor Yellow
        exit 1
    }

    Write-Info "Installing rclone via winget (this may take a minute)..."
    Write-Host ""

    try {
        $installResult = & winget install --id "Rclone.Rclone" --silent --accept-source-agreements --accept-package-agreements 2>&1
        $installResult | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    }
    catch {
        Write-Fail "Installation failed: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "  Try installing rclone manually from: https://rclone.org/downloads/" -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Step "Refreshing system PATH..."
    Refresh-PathInSession
    Start-Sleep -Seconds 2

    $rclonePath = Find-Rclone
    if (-not $rclonePath) {
        Write-Fail "Could not find rclone after installation."
        Write-Host ""
        Write-Host "  This sometimes happens. Try these steps:" -ForegroundColor Yellow
        Write-Host "    1. Close this window completely" -ForegroundColor Yellow
        Write-Host "    2. Open a NEW PowerShell window" -ForegroundColor Yellow
        Write-Host "    3. Run this script again" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  If it still doesn't work, install rclone manually from:" -ForegroundColor Yellow
        Write-Host "  https://rclone.org/downloads/" -ForegroundColor Cyan
        exit 1
    }

    Write-OK "rclone installed successfully!"
    Write-Info "Location: $rclonePath"
    try {
        $verOutput = & $rclonePath version 2>&1 | Select-Object -First 1
        Write-Info "Version:  $verOutput"
    } catch {}
}

# =================================================================================
#  STEP 2: CONFIGURE GOOGLE DRIVE CONNECTION
# =================================================================================

Write-Header "STEP 2: GOOGLE DRIVE CONNECTION"

# Check if the remote already exists
$remoteExists = $false
try {
    $remotes = & $rclonePath listremotes 2>&1
    if ($remotes -match "${remoteName}:") {
        $remoteExists = $true
    }
} catch {}

if ($remoteExists) {
    Write-OK "Google Drive connection '$remoteName' already exists"
    Write-Step "Testing the connection..."
    Write-Host ""

    $testOk = $false
    try {
        $testOutput = & $rclonePath lsd "${remoteName}:" 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0 -and $testOutput) {
            $testOk = $true
            Write-OK "Connection is working! Found these folders:"
            Write-Host ""
            foreach ($line in $testOutput) {
                Write-Host "    $line" -ForegroundColor DarkGray
            }
        }
    } catch {}

    if (-not $testOk) {
        Write-Fail "Connection test failed. The token may have expired."
        Write-Host ""
        $reconfig = Read-Host "  Would you like to set it up again? (y/n, q to quit)"
        Test-Quit $reconfig
        if ($reconfig.ToLower() -eq 'y') {
            # Delete old config and redo
            try { & $rclonePath config delete $remoteName 2>&1 | Out-Null } catch {}
            $remoteExists = $false
        }
        else {
            Write-Host "  Cannot continue without a working connection." -ForegroundColor Red
            exit 1
        }
    }
}

if (-not $remoteExists) {
    Write-Host ""
    Write-Host "  We need to connect to Google Drive. Here's what will happen:" -ForegroundColor White
    Write-Host ""
    Write-Host "    1. A browser window will open automatically" -ForegroundColor Gray
    Write-Host "    2. Sign in with the Google account that has access to the files" -ForegroundColor Gray
    Write-Host "    3. Click 'Allow' to grant read-only access" -ForegroundColor Gray
    Write-Host "    4. Close the browser tab when it says 'Success'" -ForegroundColor Gray
    Write-Host "    5. Come back to this window" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  " -NoNewline
    Write-Host "IMPORTANT:" -ForegroundColor Red -NoNewline
    Write-Host " Make sure you sign in with the correct Google account!" -ForegroundColor Yellow
    Write-Host ""

    $authReady = Read-Host "  Ready? Press Enter to open the browser (or 'q' to quit)"
    if ($authReady.ToLower() -eq 'q') {
        Write-Host "  Exiting." -ForegroundColor Yellow
        exit 0
    }

    Write-Step "Opening browser for Google sign-in..."
    Write-Host ""
    Write-Host "  Waiting for you to sign in..." -ForegroundColor Yellow
    Write-Host "  (this window will update automatically when done)" -ForegroundColor Gray
    Write-Host ""

    # Create the remote -- this triggers the OAuth browser flow
    try {
        & $rclonePath config create $remoteName drive scope=drive.readonly config_is_local=true 2>&1 |
            ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    }
    catch {
        Write-Fail "Failed to create Google Drive connection: $($_.Exception.Message)"
        exit 1
    }

    # Check if it worked
    $authCheck = $false
    try {
        $remotes = & $rclonePath listremotes 2>&1
        if ($remotes -match "${remoteName}:") { $authCheck = $true }
    } catch {}

    if (-not $authCheck) {
        Write-Fail "Something went wrong with the setup."
        Write-Host "  Try running this script again." -ForegroundColor Yellow
        exit 1
    }

    Write-OK "Google sign-in successful!"
    Write-Host ""

    # -- Shared Drive selection --
    Write-Step "Looking for Shared Drives on your Google account..."
    Write-Host ""

    $sharedDrivesRaw = $null
    try {
        $sharedDrivesRaw = & $rclonePath backend drives "${remoteName}:" 2>&1
    } catch {}

    # Parse the shared drives JSON output
    $sharedDrives = @()
    if ($sharedDrivesRaw) {
        try {
            $parsed = $sharedDrivesRaw | Out-String | ConvertFrom-Json
            if ($parsed -and $parsed.Count -gt 0) {
                foreach ($sd in $parsed) {
                    $sharedDrives += [PSCustomObject]@{
                        Name = $sd.name
                        Id   = $sd.id
                    }
                }
            }
        }
        catch {
            # If JSON parsing fails, try line-by-line parsing as fallback
            Write-Info "Could not auto-detect shared drives."
        }
    }

    if ($sharedDrives.Count -gt 0) {
        # Auto-detect the RetroFE ALU drive
        $retroFeIdx = $null
        $sdIdx = 0
        foreach ($sd in $sharedDrives) {
            $sdIdx++
            if ($sd.Name -match 'RetroFE') { $retroFeIdx = $sdIdx }
        }

        Write-Host "  Found $($sharedDrives.Count) Shared Drive(s):" -ForegroundColor White
        Write-Host ""
        $sdIdx = 0
        foreach ($sd in $sharedDrives) {
            $sdIdx++
            if ($sdIdx -eq $retroFeIdx) {
                Write-Host "    [$sdIdx]  $($sd.Name)  " -NoNewline -ForegroundColor White
                Write-Host "<-- This one!" -ForegroundColor Green
            }
            else {
                Write-Host "    [$sdIdx]  $($sd.Name)" -ForegroundColor DarkGray
            }
        }
        Write-Host "    [0]  Use My Drive (personal, not shared)" -ForegroundColor DarkGray
        Write-Host ""

        if ($retroFeIdx) {
            Write-Host "  You should select " -NoNewline -ForegroundColor Gray
            Write-Host "RetroFE ALU" -NoNewline -ForegroundColor Yellow
            Write-Host " -- that's where the files are." -ForegroundColor Gray
            Write-Host ""
            $sdChoice = Read-Host "  Select a Shared Drive (enter number, Enter for $retroFeIdx, q to quit)"
            Test-Quit $sdChoice
            if ($sdChoice.Trim() -eq '') { $sdChoice = "$retroFeIdx" }
        }
        else {
            Write-Host "  Look for a drive named " -NoNewline -ForegroundColor Gray
            Write-Host "RetroFE ALU" -NoNewline -ForegroundColor Yellow
            Write-Host " and select it." -ForegroundColor Gray
            Write-Host ""
            $sdChoice = Read-Host "  Select a Shared Drive (enter number, q to quit)"
            Test-Quit $sdChoice
        }

        if ($sdChoice -match '^\d+$' -and [int]$sdChoice -ge 1 -and [int]$sdChoice -le $sharedDrives.Count) {
            $selectedDrive = $sharedDrives[[int]$sdChoice - 1]
            Write-Host ""
            Write-Step "Configuring for Shared Drive: $($selectedDrive.Name)"
            Write-Host ""
            Write-Host "  Your browser may open again for Google sign-in." -ForegroundColor Yellow
            Write-Host "  If it does, sign in with the same account and click Allow." -ForegroundColor Yellow
            Write-Host ""

            try {
                & $rclonePath config update $remoteName team_drive=$($selectedDrive.Id) 2>&1 | Out-Null
                Write-OK "Shared Drive configured: $($selectedDrive.Name)"
            }
            catch {
                Write-Fail "Failed to set Shared Drive. You may need to configure manually."
            }
        }
        else {
            Write-Info "Using personal Google Drive (no Shared Drive selected)"
        }
    }
    else {
        Write-Info "No Shared Drives found. Using personal Google Drive."
        Write-Info "(This is normal if you haven't been invited to a Shared Drive yet)"
    }

    # Final verification
    Write-Host ""
    Write-Step "Testing the connection..."
    Write-Host ""

    $finalTest = $false
    try {
        $testOutput = & $rclonePath lsd "${remoteName}:" 2>&1
        if ($LASTEXITCODE -eq 0 -and $testOutput) {
            $finalTest = $true
            Write-OK "Everything is working! Found these folders:"
            Write-Host ""
            foreach ($line in $testOutput) {
                Write-Host "    $line" -ForegroundColor DarkGray
            }
        }
    } catch {}

    if (-not $finalTest) {
        Write-Fail "Connection test failed."
        Write-Host ""
        Write-Host "  This might mean:" -ForegroundColor Yellow
        Write-Host "    - You signed in with the wrong Google account" -ForegroundColor Yellow
        Write-Host "    - You don't have access to the Shared Drive yet" -ForegroundColor Yellow
        Write-Host "    - The Shared Drive selection was incorrect" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  To start over, run this script again." -ForegroundColor Yellow
        Write-Host "  The script will detect the broken config and offer to redo it." -ForegroundColor Gray
        exit 1
    }
}

# =================================================================================
#  STEP 3: WHERE TO DOWNLOAD
# =================================================================================

Write-Header "STEP 3: WHERE DO YOU WANT TO DOWNLOAD TO?"

Write-Host ""
Write-Host "  You'll need at least " -NoNewline -ForegroundColor Gray
Write-Host "1 TB" -NoNewline -ForegroundColor Yellow
Write-Host " of free space for the full download." -ForegroundColor Gray
Write-Host ""

# Build disk info lookup (bus type, model)
$diskInfoByLetter = @{}
try {
    $partitions = Get-Partition -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
    $disks = Get-Disk -ErrorAction SilentlyContinue
    $diskLookup = @{}
    foreach ($d in $disks) { $diskLookup[$d.Number] = $d }
    foreach ($p in $partitions) {
        $dl = [string]$p.DriveLetter
        $dk = $diskLookup[$p.DiskNumber]
        if ($dk) {
            $diskInfoByLetter[$dl] = [PSCustomObject]@{
                BusType = $dk.BusType
                Model   = $dk.FriendlyName
            }
        }
    }
} catch {}

# Show available drives
$drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
          Where-Object { $_.Free -gt 0 -and $_.Root } |
          Sort-Object Name

$driveList = @()
$idx = 0

foreach ($drv in $drives) {
    $idx++
    $freeSpace = Format-FileSize $drv.Free
    $totalSpace = Format-FileSize ($drv.Used + $drv.Free)

    $drvLabel = ""
    try {
        $vol = Get-Volume -DriveLetter $drv.Name -ErrorAction SilentlyContinue
        if ($vol -and $vol.FileSystemLabel) { $drvLabel = $vol.FileSystemLabel }
    } catch {}

    $busTag = ""
    $modelTag = ""
    $di = $diskInfoByLetter[[string]$drv.Name]
    if ($di) {
        switch ($di.BusType) {
            'USB'   { $busTag = "USB" }
            'NVMe'  { $busTag = "NVMe" }
            'SATA'  { $busTag = "SATA" }
            'SAS'   { $busTag = "SAS" }
            default { if ($di.BusType) { $busTag = "$($di.BusType)" } }
        }
        if ($di.Model) { $modelTag = $di.Model }
    }

    $label = "$($drv.Name):\ drive"
    $details = @()
    if ($busTag)   { $details += $busTag }
    if ($drvLabel) { $details += $drvLabel }
    if ($modelTag) { $details += $modelTag }
    $detailStr = ""
    if ($details.Count -gt 0) { $detailStr = "  [$($details -join ' - ')]" }

    $driveList += [PSCustomObject]@{
        Index = $idx
        Letter = $drv.Name
        Root  = $drv.Root
        Free  = $drv.Free
    }

    $freeColor = if ($drv.Free -ge 1TB) { "Green" } elseif ($drv.Free -ge 500GB) { "Yellow" } else { "Red" }
    Write-Host "    [$idx]  $label$detailStr  " -NoNewline -ForegroundColor White
    Write-Host "($freeSpace free of $totalSpace)" -ForegroundColor $freeColor
}

Write-Host ""
$driveChoice = Read-Host "  Pick a drive (enter number, q to quit)"
Test-Quit $driveChoice

$destRoot = $null
if ($driveChoice -match '^\d+$') {
    $chosen = $driveList | Where-Object { $_.Index -eq [int]$driveChoice }
    if ($chosen) {
        $destRoot = $chosen.Root
    }
}

if (-not $destRoot) {
    Write-Host ""
    $customPath = Read-Host "  Enter the full path (e.g. E:\Downloads, q to quit)"
    Test-Quit $customPath
    if ($customPath -and $customPath.Trim() -ne '') {
        $destRoot = $customPath.Trim()
    }
    else {
        Write-Fail "No destination selected. Exiting."
        exit 1
    }
}

Write-Host ""
Write-Host "  What do you want to name the download folder?" -ForegroundColor White
Write-Host ""
Write-Host "  Examples: " -NoNewline -ForegroundColor Gray
Write-Host "AwesomeSauce2" -NoNewline -ForegroundColor Yellow
Write-Host ", " -NoNewline -ForegroundColor Gray
Write-Host "RetroFE-Download" -NoNewline -ForegroundColor Yellow
Write-Host ", " -NoNewline -ForegroundColor Gray
Write-Host "AS2" -ForegroundColor Yellow
Write-Host ""

$folderName = ""
while (-not $folderName -or $folderName.Trim() -eq '') {
    $folderName = Read-Host "  Folder name (q to quit)"
    Test-Quit $folderName
    if (-not $folderName -or $folderName.Trim() -eq '') {
        Write-Host "  A folder name is required so the files stay organized." -ForegroundColor Red
    }
}

$downloadPath = Join-Path $destRoot $folderName.Trim()

if (-not (Test-Path $downloadPath)) {
    New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
    Write-OK "Created: $downloadPath"
}

Write-Host ""
Write-Host "  Download destination: " -NoNewline -ForegroundColor White
Write-Host "$downloadPath" -ForegroundColor Green

# Space warning
$destQualifier = Split-Path -Qualifier $downloadPath -ErrorAction SilentlyContinue
if ($destQualifier) {
    $dl = $destQualifier.TrimEnd(':')
    $psd = Get-PSDrive -Name $dl -ErrorAction SilentlyContinue
    if ($psd -and $psd.Free -lt 1TB) {
        Write-Host ""
        Write-Host "  WARNING: This drive has less than 1 TB free." -ForegroundColor Red
        Write-Host "  You may not have enough space for the full download." -ForegroundColor Red
        Write-Host "  Free space: $(Format-FileSize $psd.Free)" -ForegroundColor Yellow
        Write-Host ""
        $spaceOk = Read-Host "  Continue anyway? (y/n, q to quit)"
        Test-Quit $spaceOk
        if ($spaceOk.ToLower() -ne 'y') {
            Write-Host "  Exiting. Pick a drive with more space and try again." -ForegroundColor Yellow
            exit 0
        }
    }
}

# =================================================================================
#  STEP 4: BANDWIDTH SETTINGS
# =================================================================================

Write-Header "STEP 4: DOWNLOAD SPEED"

Write-Host ""
Write-Host "  The download will use your full internet speed by default." -ForegroundColor White
Write-Host "  If you want to browse the web or stream video while downloading," -ForegroundColor Gray
Write-Host "  you can limit the download speed." -ForegroundColor Gray
Write-Host ""
Write-Host "    [1]  Full speed (use all bandwidth)" -ForegroundColor White
Write-Host "    [2]  Limit speed (save some for other use)" -ForegroundColor White
Write-Host ""
$bwChoice = Read-Host "  Choice (1, 2, or q to quit)"
Test-Quit $bwChoice

$bwLimit = $null

if ($bwChoice -eq '2') {
    Write-Host ""
    Write-Host "  What is your internet speed? Pick the closest option:" -ForegroundColor White
    Write-Host ""
    Write-Host "    Internet Speed     Download Limit    " -ForegroundColor Gray
    Write-Host "    ----------------------------------------" -ForegroundColor Gray
    Write-Host "    [1]   100 Mbps       10 MB/s" -ForegroundColor White
    Write-Host "    [2]   300 Mbps       30 MB/s" -ForegroundColor White
    Write-Host "    [3]   500 Mbps       50 MB/s" -ForegroundColor White
    Write-Host "    [4]   750 Mbps       75 MB/s" -ForegroundColor White
    Write-Host "    [5]   1 Gbps        100 MB/s" -ForegroundColor White
    Write-Host "    [6]   2 Gbps        200 MB/s" -ForegroundColor White
    Write-Host ""
    Write-Host "  (These limits use ~80% of your bandwidth, leaving ~20% for other use)" -ForegroundColor Gray
    Write-Host ""

    $speedChoice = Read-Host "  Pick a speed (1-6, custom value like '50M', or q to quit)"
    Test-Quit $speedChoice

    switch ($speedChoice) {
        '1' { $bwLimit = '10M' }
        '2' { $bwLimit = '30M' }
        '3' { $bwLimit = '50M' }
        '4' { $bwLimit = '75M' }
        '5' { $bwLimit = '100M' }
        '6' { $bwLimit = '200M' }
        default {
            if ($speedChoice -match '^\d+M?$') {
                $bwLimit = $speedChoice
                if ($bwLimit -notmatch 'M$') { $bwLimit += 'M' }
            }
            else {
                Write-Info "Invalid choice. Using full speed."
                $bwLimit = $null
            }
        }
    }

    if ($bwLimit) {
        Write-OK "Download speed limited to $bwLimit/s"
    }
}
else {
    Write-OK "Using full download speed"
}

# =================================================================================
#  STEP 5: DOWNLOAD
# =================================================================================

Write-Header "STEP 5: DOWNLOADING FILES"

Write-Host ""
Write-Host "  Source:      Google Drive ($remoteName)" -ForegroundColor White
Write-Host "  Destination: $downloadPath" -ForegroundColor White
if ($bwLimit) {
    Write-Host "  Speed limit: $bwLimit/s" -ForegroundColor White
}
Write-Host ""
Write-Host "  The download will show progress as it runs." -ForegroundColor Gray
Write-Host "  This will take a while for large collections." -ForegroundColor Gray
Write-Host ""
Write-Host "  " -NoNewline
Write-Host "TIP:" -ForegroundColor Yellow -NoNewline
Write-Host " If you need to stop, press " -NoNewline -ForegroundColor Gray
Write-Host "Ctrl+C" -NoNewline -ForegroundColor Yellow
Write-Host ". Run this script again to resume!" -ForegroundColor Gray
Write-Host "  Files that were already downloaded will NOT be re-downloaded." -ForegroundColor Gray
Write-Host ""

$startNow = Read-Host "  Start downloading? (y/n, q to quit)"
Test-Quit $startNow
if ($startNow.ToLower() -ne 'y') {
    Write-Host "  No problem. Run this script whenever you're ready!" -ForegroundColor Gray
    Write-Host "  All your settings have been saved." -ForegroundColor Gray
    exit 0
}

Write-Host ""
Write-Host "  Starting download..." -ForegroundColor Green
Write-Host ""

# Build the rclone command arguments
$rcloneArgs = @(
    'copy',
    '--buffer-size', '500M',
    '--local-no-sparse',
    '-P',
    '-v',
    '--stats', '10s'
)

if ($bwLimit) {
    $rcloneArgs += '--bwlimit'
    $rcloneArgs += $bwLimit
}

$rcloneArgs += "${remoteName}:/"
$rcloneArgs += $downloadPath

# Run the download
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    & $rclonePath @rcloneArgs
    $dlExitCode = $LASTEXITCODE
}
catch {
    $dlExitCode = 1
    Write-Fail "An error occurred: $($_.Exception.Message)"
}

$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed
$mins = [math]::Floor($elapsed.TotalMinutes)
$secs = $elapsed.Seconds
$timeStr = if ($mins -gt 60) {
    "$([math]::Floor($mins / 60))h $($mins % 60)m"
} elseif ($mins -gt 0) {
    "${mins}m ${secs}s"
} else {
    "$($elapsed.TotalSeconds.ToString('N0'))s"
}

# =================================================================================
#  DONE
# =================================================================================

Write-Host ""

if ($dlExitCode -eq 0) {
    Write-Header "DOWNLOAD COMPLETE!"
    Write-Host ""
    Write-OK "All files have been downloaded successfully!"
    Write-Host ""
    Write-Host "  Destination: $downloadPath" -ForegroundColor Green
    Write-Host "  Time:        $timeStr" -ForegroundColor White
    Write-Host ""
    Write-Host "  You can now use these files or run the Extract-To-USB script" -ForegroundColor Gray
    Write-Host "  to build your USB drive." -ForegroundColor Gray
}
else {
    Write-Header "DOWNLOAD INTERRUPTED OR INCOMPLETE"
    Write-Host ""
    Write-Host "  The download was interrupted or encountered errors." -ForegroundColor Yellow
    Write-Host "  Time elapsed: $timeStr" -ForegroundColor White
    Write-Host ""
    Write-Host "  " -NoNewline
    Write-Host "Don't worry!" -ForegroundColor Green -NoNewline
    Write-Host " Just run this script again to resume." -ForegroundColor White
    Write-Host "  It will skip files that were already downloaded." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Destination: $downloadPath" -ForegroundColor White
}

Write-Host ""
