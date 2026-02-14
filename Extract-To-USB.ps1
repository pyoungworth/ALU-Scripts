<#
.SYNOPSIS
    Extracts zip files from selected folders into a single output folder for USB.

.DESCRIPTION
    Walks you through:
      1. Where are the downloaded files (pick a drive and folder)
      2. Where to extract to (pick a drive or use a folder here)
      3. Extract everything or customize (system packs, themes, marquees, screensaver, Daphne games)
      4. Reviews everything and checks free space before starting
      5. Extracts all zips, preserving internal folder structure
      6. Supports resume if interrupted -- just run again

.PARAMETER Path
    The root directory containing the downloaded files. If not provided, the script
    will guide you through selecting the drive and folder interactively.

.EXAMPLE
    .\Extract-To-USB.ps1
    .\Extract-To-USB.ps1 -Path "F:\AS2"

.NOTES
    Author: Philip Youngworth
    Project: ALU-Scripts (AwesomeSauce2 Utility Scripts)
    License: MIT
#>

param(
    [Parameter(Position = 0)]
    [string]$Path
)

# Load zip support for reading real uncompressed sizes from headers
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

# =================================================================================
#  STEP 1: WHERE ARE THE DOWNLOADED FILES?
# =================================================================================

Write-Header "STEP 1: WHERE ARE THE DOWNLOADED FILES?"

if ($Path) {
    # User passed -Path on the command line -- use it directly
    if (-not (Test-Path $Path)) {
        Write-Fail "Path not found: $Path"
        exit 1
    }
    $Path = (Resolve-Path $Path).Path
}
else {
    # Check if the download script saved a path from a previous run
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $savedPathFile = Join-Path $scriptDir '.last-download-path'
    $savedPath = $null

    if (Test-Path $savedPathFile) {
        $savedPath = (Get-Content $savedPathFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
        if ($savedPath -and (Test-Path $savedPath)) {
            Write-Host ""
            Write-Host "  Last download location: " -NoNewline -ForegroundColor White
            Write-Host "$savedPath" -ForegroundColor Green
            Write-Host ""
            $useSaved = Read-Host "  Use this folder? (y/n, q to quit)"
            Test-Quit $useSaved
            if ($useSaved.ToLower() -ne 'n') {
                $Path = $savedPath
            }
        }
    }

    if (-not $Path) {
    Write-Host ""
    Write-Host "  Where did you download the files to?" -ForegroundColor White
    Write-Host "  (This is the folder you chose in the download script)" -ForegroundColor Gray
    Write-Host ""

    # Build disk info lookup for drive display
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
            Index  = $idx
            Letter = $drv.Name
            Root   = $drv.Root
        }

        Write-Host "    [$idx]  $label$detailStr  ($freeSpace free of $totalSpace)" -ForegroundColor White
    }

    Write-Host ""
    $srcDriveChoice = Read-Host "  Pick the drive where you downloaded to (enter number, q to quit)"
    Test-Quit $srcDriveChoice

    $srcRoot = $null
    if ($srcDriveChoice -match '^\d+$') {
        $chosen = $driveList | Where-Object { $_.Index -eq [int]$srcDriveChoice }
        if ($chosen) { $srcRoot = $chosen.Root }
    }

    if (-not $srcRoot) {
        Write-Fail "Invalid selection. Exiting."
        exit 1
    }

    # Ask for the folder name
    Write-Host ""
    Write-Host "  What was the folder name you used? (e.g. AS2, AwesomeSauce2)" -ForegroundColor White
    Write-Host ""
    $srcFolder = Read-Host "  Folder name (q to quit)"
    Test-Quit $srcFolder

    if (-not $srcFolder -or $srcFolder.Trim() -eq '') {
        Write-Fail "No folder name entered. Exiting."
        exit 1
    }

    $Path = Join-Path $srcRoot $srcFolder.Trim()

    if (-not (Test-Path $Path)) {
        Write-Fail "Folder not found: $Path"
        Write-Host ""
        Write-Host "  Make sure you've run the download script first and the folder exists." -ForegroundColor Yellow
        exit 1
    }
    } # end of manual path selection
}

# Verify the folder has expected content (subfolders with zips)
$subDirCheck = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue
$zipCheck = Get-ChildItem -Path $Path -Filter '*.zip' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $subDirCheck -or -not $zipCheck) {
    Write-Fail "No build files found in: $Path"
    Write-Host ""
    Write-Host "  This folder should contain subfolders with .zip files from the download." -ForegroundColor Yellow
    Write-Host "  Make sure you selected the right folder." -ForegroundColor Yellow
    exit 1
}

Write-OK "Source folder: $Path"

$subCount = ($subDirCheck | Measure-Object).Count
$zipCount = (Get-ChildItem -Path $Path -Filter '*.zip' -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
Write-Info "  Found $subCount folder(s) with $zipCount zip file(s)"

# =================================================================================
#  STEP 2: WHERE TO EXTRACT
# =================================================================================

Write-Header "STEP 2: WHERE DO YOU WANT TO EXTRACT TO?"
Write-Host ""

# Build a lookup of disk info (bus type, model, filesystem) keyed by drive letter
$diskInfoByLetter = @{}
try {
    $partitions = Get-Partition -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
    $disks = Get-Disk -ErrorAction SilentlyContinue
    $diskLookup = @{}
    foreach ($d in $disks) { $diskLookup[$d.Number] = $d }
    foreach ($p in $partitions) {
        $dl = [string]$p.DriveLetter
        $dk = $diskLookup[$p.DiskNumber]
        # Get the filesystem type from the volume
        $fsType = ""
        try {
            $vol = Get-Volume -DriveLetter $dl -ErrorAction SilentlyContinue
            if ($vol -and $vol.FileSystemType) { $fsType = $vol.FileSystemType }
        } catch {}
        if ($dk) {
            $diskInfoByLetter[$dl] = [PSCustomObject]@{
                BusType    = $dk.BusType
                Model      = $dk.FriendlyName
                FileSystem = $fsType
            }
        }
    }
} catch {}

# Show available drives with free space
$drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
          Where-Object { $_.Free -gt 0 -and $_.Root } |
          Sort-Object Name

$driveList = @()
$idx = 0

foreach ($drv in $drives) {
    $idx++
    $freeSpace = Format-FileSize $drv.Free
    $totalSpace = Format-FileSize ($drv.Used + $drv.Free)

    # Volume label
    $drvLabel = ""
    try {
        $vol = Get-Volume -DriveLetter $drv.Name -ErrorAction SilentlyContinue
        if ($vol -and $vol.FileSystemLabel) { $drvLabel = $vol.FileSystemLabel }
    } catch {}

    # Bus type & disk model from lookup
    $busTag = ""
    $modelTag = ""
    $di = $diskInfoByLetter[[string]$drv.Name]
    if ($di) {
        switch ($di.BusType) {
            'USB'   { $busTag = "USB" }
            'NVMe'  { $busTag = "NVMe" }
            'SATA'  { $busTag = "SATA" }
            'SAS'   { $busTag = "SAS" }
            'RAID'  { $busTag = "RAID" }
            default { if ($di.BusType) { $busTag = "$($di.BusType)" } }
        }
        if ($di.Model) { $modelTag = $di.Model }
    }

    # Get filesystem type
    $fsType = ""
    if ($di -and $di.FileSystem) { $fsType = $di.FileSystem }

    # Build the display line:  [1]  H:\ drive  [NTFS]  USB - Realtek RTL9210B-CG  (925 GB free of 925 GB)
    $label = "$($drv.Name):\ drive"
    $details = @()
    if ($busTag)   { $details += $busTag }
    if ($drvLabel) { $details += $drvLabel }
    if ($modelTag) { $details += $modelTag }
    $detailStr = ""
    if ($details.Count -gt 0) { $detailStr = "  [$($details -join ' - ')]" }

    $fsTag = ""
    $isNTFS = ($fsType -eq 'NTFS')
    if ($fsType) { $fsTag = "  ($fsType)" }

    $driveList += [PSCustomObject]@{
        Index      = $idx
        Letter     = $drv.Name
        Root       = $drv.Root
        Free       = $drv.Free
        FreeText   = $freeSpace
        FileSystem = $fsType
    }

    if ($isNTFS) {
        Write-Host "    [$idx]  $label$detailStr$fsTag  ($freeSpace free of $totalSpace)" -ForegroundColor White
    }
    else {
        Write-Host "    [$idx]  $label$detailStr$fsTag  ($freeSpace free of $totalSpace)" -ForegroundColor DarkGray -NoNewline
        Write-Host "  ** Not NTFS **" -ForegroundColor Red
    }
}

$idx++
$driveList += [PSCustomObject]@{
    Index    = $idx
    Letter   = $null
    Root     = $null
    Free     = 0
    FreeText = ""
}
Write-Host "    [$idx]  Enter a custom path" -ForegroundColor White

Write-Host ""
Write-Host "  NOTE: " -ForegroundColor Yellow -NoNewline
Write-Host "The destination drive must be formatted as NTFS." -ForegroundColor White
Write-Host "        Drives marked " -ForegroundColor Gray -NoNewline
Write-Host "** Not NTFS **" -ForegroundColor Red -NoNewline
Write-Host " will need to be reformatted before use." -ForegroundColor Gray
Write-Host ""

$destDrive = $null
while (-not $destDrive) {
    $driveChoice = Read-Host "  Pick a destination (enter number, q to quit)"
    Test-Quit $driveChoice

    if ($driveChoice -match '^\d+$') {
        $chosen = $driveList | Where-Object { $_.Index -eq [int]$driveChoice }
        if ($chosen -and $chosen.Root) {
            # Check NTFS requirement
            if ($chosen.FileSystem -and $chosen.FileSystem -ne 'NTFS') {
                Write-Host ""
                Write-Host "  *** $($chosen.Letter):\ is formatted as $($chosen.FileSystem) -- this will NOT work ***" -ForegroundColor Red
                Write-Host ""
                Write-Host "  The ALU build requires NTFS because it uses files larger than 4 GB" -ForegroundColor Yellow
                Write-Host "  and relies on NTFS features like long file paths." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  To use this drive, you'll need to reformat it as NTFS:" -ForegroundColor White
                Write-Host "    1. Open File Explorer, right-click on $($chosen.Letter):\, select 'Format...'" -ForegroundColor Gray
                Write-Host "    2. Change 'File system' to NTFS" -ForegroundColor Gray
                Write-Host "    3. Click Start (this will erase everything on the drive)" -ForegroundColor Gray
                Write-Host ""
                Write-Host "  Please pick a different drive or reformat first." -ForegroundColor Yellow
                Write-Host ""
                continue
            }
            $destDrive = $chosen
        }
        elseif ($chosen -and -not $chosen.Root) {
            # Custom path -- check NTFS on the underlying drive
            Write-Host ""
            $customPath = Read-Host "  Enter the full path (e.g. D:\MyFolder, \\server\share)"
            if (-not $customPath -or $customPath.Trim() -eq '') {
                Write-Host "  No path entered. Exiting." -ForegroundColor Yellow
                exit 0
            }
            $customPath = $customPath.Trim()

            # Check NTFS on the custom path's drive
            $customQualifier = Split-Path -Qualifier $customPath -ErrorAction SilentlyContinue
            if ($customQualifier) {
                $customDL = $customQualifier.TrimEnd(':')
                $customFS = ""
                try {
                    $customVol = Get-Volume -DriveLetter $customDL -ErrorAction SilentlyContinue
                    if ($customVol -and $customVol.FileSystemType) { $customFS = $customVol.FileSystemType }
                } catch {}
                if ($customFS -and $customFS -ne 'NTFS') {
                    Write-Host ""
                    Write-Host "  *** $customQualifier\ is formatted as $customFS -- this will NOT work ***" -ForegroundColor Red
                    Write-Host ""
                    Write-Host "  The ALU build requires NTFS because it uses files larger than 4 GB" -ForegroundColor Yellow
                    Write-Host "  and relies on NTFS features like long file paths." -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "  Please pick a different drive or reformat first." -ForegroundColor Yellow
                    Write-Host ""
                    continue
                }
            }

            if (-not (Test-Path $customPath)) {
                Write-Host "  Path does not exist. Create it? (y/n)" -ForegroundColor Yellow
                $mk = Read-Host "  "
                if ($mk.ToLower() -eq 'y') {
                    New-Item -Path $customPath -ItemType Directory -Force | Out-Null
                }
                else {
                    Write-Host "  Exiting." -ForegroundColor Yellow
                    exit 0
                }
            }
            $driveLetter = $customQualifier
            $freeBytes = 0
            if ($driveLetter) {
                $dl = $driveLetter.TrimEnd(':')
                $psd = Get-PSDrive -Name $dl -ErrorAction SilentlyContinue
                if ($psd) { $freeBytes = $psd.Free }
            }
            $destDrive = [PSCustomObject]@{
                Index      = 0
                Letter     = $null
                Root       = $customPath
                Free       = $freeBytes
                FreeText   = Format-FileSize $freeBytes
                FileSystem = $customFS
            }
        }
        else {
            Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow
    }
}

# Ask for optional folder name
Write-Host ""
Write-Host "  Enter a folder name, or press Enter to extract directly to the drive root." -ForegroundColor White
Write-Host ""
Write-Host "  IMPORTANT: " -ForegroundColor Red -NoNewline
Write-Host "If you are extracting to an external USB drive, just press Enter." -ForegroundColor Yellow
Write-Host "             Do NOT enter a folder name -- files need to go to the root of the drive." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Only enter a name if you're extracting to a local drive and want a subfolder." -ForegroundColor Gray
Write-Host "  Example: AwesomeSauce2-Extracted, AS2-Build, etc." -ForegroundColor Gray
Write-Host ""

$folderName = Read-Host "  Folder name (optional)"

if ($destDrive.Letter) {
    $basePath = "$($destDrive.Letter):\"
}
else {
    $basePath = $destDrive.Root
}

if ($folderName -and $folderName.Trim() -ne '') {
    $outputPath = Join-Path $basePath $folderName.Trim()
}
else {
    $outputPath = $basePath
}

if (Test-Path $outputPath) {
    # Only warn if it's a subfolder (not a bare drive root)
    if ($folderName -and $folderName.Trim() -ne '') {
        Write-Host ""
        Write-Host "  '$outputPath' already exists." -ForegroundColor Yellow
        $cont = Read-Host "  Continue extracting into it? (y/n)"
        if ($cont.ToLower() -ne 'y') {
            Write-Host "  Exiting." -ForegroundColor Yellow
            exit 0
        }
    }
}
else {
    New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "  Destination: $outputPath" -ForegroundColor Green

# =================================================================================
#  STEP 3: CONTENT SELECTION
# =================================================================================

Write-Header "STEP 3: CONTENT SELECTION"

# Gather all subfolders
$subDirs = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
           Where-Object { $_.FullName -ne $outputPath }

if ($subDirs.Count -eq 0) {
    Write-Host "  No source folders found in $Path" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "    [A]  Extract EVERYTHING (default)" -ForegroundColor Green
Write-Host "    [C]  Customize what to extract" -ForegroundColor Yellow
Write-Host ""

$extractChoice = Read-Host "  Choice (A/C, q to quit)"
Test-Quit $extractChoice

$customizeExtraction = ($extractChoice.ToUpper() -eq 'C')

if ($customizeExtraction) {
    Write-Host ""
    Write-Host "  For each optional item below, choose whether to include it." -ForegroundColor Gray
    Write-Host "  Required folders will be included automatically." -ForegroundColor Gray
}
else {
    Write-OK "Extracting everything"
}

# Helper: get actual uncompressed size of a single zip from its headers
function Get-ZipExtractedSize {
    param([string]$ZipPath)
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $total = [long]($archive.Entries | Measure-Object -Property Length -Sum).Sum
        $archive.Dispose()
        return $total
    }
    catch {
        # Fallback: estimate at 1.05x if we can't read the zip
        return [long]((Get-Item $ZipPath).Length * 1.05)
    }
}

# Helper: get all zips in a folder (including immediate subfolders) and their total sizes
function Get-ZipInfo {
    param([string]$FolderPath)
    $zips = @(Get-ChildItem -Path $FolderPath -Filter '*.zip' -File -ErrorAction SilentlyContinue)
    $subFolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    foreach ($sf in $subFolders) {
        $zips += @(Get-ChildItem -Path $sf.FullName -Filter '*.zip' -File -ErrorAction SilentlyContinue)
    }
    $compressedSize = ($zips | Measure-Object -Property Length -Sum).Sum
    if (-not $compressedSize) { $compressedSize = 0 }

    # Read actual uncompressed sizes from zip headers
    $extractedSize = [long]0
    foreach ($z in $zips) {
        $extractedSize += Get-ZipExtractedSize -ZipPath $z.FullName
    }

    return [PSCustomObject]@{
        Zips          = $zips
        Count         = $zips.Count
        Size          = $compressedSize
        ExtractedSize = $extractedSize
    }
}

# -- Daphne game selection helpers -------------------------------------------------

# Friendly display names for Daphne games
$script:DaphneGameNames = @{
    'ace'            = "Space Ace"
    'actionmax'      = "Action Max"
    'asterix'        = "Asterix"
    'astron'         = "Astron Belt"
    'badlands'       = "Badlands"
    'bega'           = "Bega's Battle"
    'carbon'         = "Carbon"
    'chantze-hd'     = "Chantze HD"
    'cliff'          = "Cliff Hanger"
    'cobra'          = "Cobra Command"
    'conan'          = "Conan"
    'crimepatrol'    = "Crime Patrol"
    'daitarn'        = "Daitarn 3"
    'dl2e'           = "Dragon's Lair II Enhanced"
    'dle21'          = "Dragon's Lair Enhanced"
    'dltv'           = "Dragon's Lair TV"
    'dragon'         = "Dragon"
    'drugwars'       = "Drug Wars"
    'esh'            = "Esh's Aurunmilla"
    'fireandice'     = "Fire and Ice"
    'freedomfighter' = "Freedom Fighter"
    'friday13'       = "Friday the 13th"
    'galaxy'         = "Galaxy Ranger"
    'gpworld'        = "GP World"
    'hayate'         = "Hayate"
    'interstellar'   = "Interstellar"
    'jack'           = "Mad Dog McCree: Jack"
    'johnnyrock'     = "Johnny Rock"
    'lair'           = "Dragon's Lair"
    'lair2'          = "Dragon's Lair II: Time Warp"
    'lbh'            = "Last Bounty Hunter"
    'mach3'          = "M.A.C.H. 3"
    'maddog'         = "Mad Dog McCree"
    'maddog2'        = "Mad Dog McCree II"
    'mononoke'       = "Princess Mononoke"
    'platoon'        = "Platoon"
    'pussinboots'    = "Puss in Boots"
    'roadblaster'    = "Road Blaster"
    'sae'            = "Space Ace Enhanced"
    'sdq'            = "Super Don Quix-ote"
    'spacepirates'   = "Space Pirates"
    'starblazers'    = "Star Blazers"
    'suckerpunch'    = "Sucker Punch"
    'timegal'        = "Time Gal"
    'timegalv2'      = "Time Gal v2"
    'timetraveler'   = "Time Traveler"
    'titanae'        = "Titan A.E."
    'triad'          = "Triad Stone"
    'tron'           = "Tron"
    'uvt'            = "Us vs Them"
}

# Dragon's Lair + Space Ace only (the iconic titles)
$script:DaphneDLSpaceAce = @(
    'lair', 'lair2', 'dle21', 'dl2e',           # Dragon's Lair series
    'ace', 'sae'                                  # Space Ace series
)

# Popular/recommended Daphne laserdisc games
$script:DaphnePopularGames = @(
    'lair', 'lair2', 'dle21', 'dl2e',           # Dragon's Lair series
    'ace', 'sae',                                 # Space Ace
    'cliff',                                      # Cliff Hanger
    'cobra',                                      # Cobra Command
    'roadblaster',                                # Road Blaster
    'badlands',                                   # Badlands
    'bega',                                       # Bega's Battle
    'astron',                                     # Astron Belt
    'galaxy',                                     # Galaxy Ranger
    'sdq',                                        # Super Don Quix-ote
    'esh',                                        # Esh's Aurunmilla
    'mach3',                                      # Mach 3
    'uvt',                                        # Us vs Them
    'timetraveler',                               # Time Traveler
    'interstellar'                                # Interstellar
)

# Scan Daphne zip and build a per-game size map
function Get-DaphneGameMap {
    param([string]$ZipPath)
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    $games = @{}
    $sharedSize = [long]0

    foreach ($entry in $archive.Entries) {
        if ($entry.Length -eq 0) { continue }
        $p = $entry.FullName
        $gameName = $null

        # Match: assets/singe/{game}/ or assets/mpegs/{game}/
        if ($p -match 'assets/(singe|mpegs)/([^/]+)/') {
            $gameName = $Matches[2]
        }
        # Match: roms/{game}.zip or roms/{game}.singe
        elseif ($p -match 'roms/([^/]+)\.(zip|singe)$') {
            $gameName = $Matches[1]
        }
        # Match: medium_artwork/{subfolder}/{game}.ext
        elseif ($p -match 'medium_artwork/[^/]+/([^/.]+)') {
            $gameName = $Matches[1]
        }

        if ($gameName) {
            if (-not $games.ContainsKey($gameName)) {
                $games[$gameName] = [long]0
            }
            $games[$gameName] += $entry.Length
        }
        else {
            # Shared/infrastructure files (system_artwork, config, etc.)
            $sharedSize += $entry.Length
        }
    }

    $archive.Dispose()
    return [PSCustomObject]@{
        Games      = $games
        SharedSize = $sharedSize
    }
}

# Top-level folder keywords that are optional (only prompted when customizing)
$optionalFolderKeywords = @('Optional', 'BitLCD', 'Marquee')

# Subfolder keywords that are optional (like screensavers inside Key Folders)
$optionalSubfolderKeywords = @('screensaver')

# System Packs folder keyword -- gets special handling with per-system selection
$systemPacksKeyword = 'System Packs'

# Collect what to extract: list of PSCustomObjects with Name, FullPath, Zips
$selectedSources = @()
$skippedExtracted = [long]0

function Ask-Optional {
    param([string]$Label, [int]$ZipCount, [long]$Size, [long]$ExtractedSize)

    $extFormatted = Format-FileSize $ExtractedSize

    Write-Host ""
    Write-Host "  Include " -NoNewline
    Write-Host "$Label" -ForegroundColor Yellow -NoNewline
    Write-Host "?" -ForegroundColor White
    Write-Host "    $ZipCount zips, ~$(Format-FileSize $Size) compressed, ~$extFormatted extracted" -ForegroundColor Gray
    Write-Host "    Selecting " -NoNewline -ForegroundColor Gray
    Write-Host "N" -ForegroundColor Red -NoNewline
    Write-Host " saves ~$extFormatted of space" -ForegroundColor Gray
    $answer = Read-Host "  (y/n)"
    return ($answer.ToLower() -eq 'y')
}

foreach ($dir in $subDirs) {
    $info = Get-ZipInfo -FolderPath $dir.FullName
    if ($info.Count -eq 0) { continue }

    # Check if this is the System Packs folder
    $isSystemPacks = ($dir.Name -match $systemPacksKeyword)

    # Check if this top-level folder is optional
    $isOptional = $false
    if (-not $isSystemPacks) {
        foreach ($kw in $optionalFolderKeywords) {
            if ($dir.Name -match $kw) { $isOptional = $true; break }
        }
    }

    if ($isSystemPacks -and $customizeExtraction) {
        # System Packs: let user pick which systems to extract
        $spZips = @(Get-ChildItem -Path $dir.FullName -Filter '*.zip' -File -ErrorAction SilentlyContinue |
                    Sort-Object Name)

        # The final list of system pack zips to include (built by selection below)
        $finalSpZips = @()

        if ($spZips.Count -gt 0) {
            Write-Host ""
            Write-Host "  Which " -NoNewline
            Write-Host "System Packs" -ForegroundColor Yellow -NoNewline
            Write-Host " do you want to extract?" -ForegroundColor White
            Write-Host ""
            Write-Host "    [A]  Extract ALL systems (default)" -ForegroundColor Green
            Write-Host "    [S]  Let me pick which systems I want" -ForegroundColor Yellow
            Write-Host ""

            $spChoice = Read-Host "  Choice (A/S)"

            if ($spChoice.ToUpper() -eq 'S') {
                # Show system list with sizes
                Write-Host ""
                $spList = @()
                $spIdx = 0
                foreach ($spz in $spZips) {
                    $spIdx++
                    $extractedSz = Get-ZipExtractedSize -ZipPath $spz.FullName
                    $displayName = [System.IO.Path]::GetFileNameWithoutExtension($spz.Name)
                    $spList += [PSCustomObject]@{
                        Index         = $spIdx
                        FileName      = $spz.Name
                        DisplayName   = $displayName
                        File          = $spz
                        ExtractedSize = $extractedSz
                    }
                }

                # Display in two columns
                $nameWidth = 40
                $colCount = 2
                $rows = [math]::Ceiling($spList.Count / $colCount)

                Write-Host "  $($spList.Count) system packs available:" -ForegroundColor White
                Write-Host ""

                for ($row = 0; $row -lt $rows; $row++) {
                    $line = ""
                    for ($col = 0; $col -lt $colCount; $col++) {
                        $i = $col * $rows + $row
                        if ($i -lt $spList.Count) {
                            $sp = $spList[$i]
                            $num = ($sp.Index).ToString().PadLeft(2)
                            $name = $sp.DisplayName
                            if ($name.Length -gt $nameWidth) { $name = $name.Substring(0, $nameWidth - 3) + "..." }
                            $szText = Format-FileSize $sp.ExtractedSize
                            $entry = "  [$num] $($name.PadRight($nameWidth)) $($szText.PadLeft(9))"
                            $line += $entry
                        }
                    }
                    Write-Host "  $line" -ForegroundColor White
                }

                Write-Host ""
                Write-Host "  HOW TO SELECT:" -ForegroundColor Yellow
                Write-Host "    all             Extract everything (default)" -ForegroundColor Gray
                Write-Host "    1,3,5           Pick specific systems" -ForegroundColor Gray
                Write-Host "    1-20            Pick a range" -ForegroundColor Gray
                Write-Host "    all,!5,!12      Extract all EXCEPT #5 and #12" -ForegroundColor Gray
                Write-Host ""

                $spResponse = Read-Host "  Selection (Enter for all)"

                if (-not $spResponse -or $spResponse.Trim() -eq '' -or $spResponse.Trim().ToLower() -eq 'all') {
                    if (-not ($spResponse -and $spResponse -match '!')) {
                        # All systems
                        $finalSpZips = $spZips
                        Write-OK "Extracting ALL system packs"
                        $spResponse = $null
                    }
                }

                if ($spResponse) {
                    # Parse selection with support for ranges and exclusions
                    $selectedNums = [System.Collections.Generic.HashSet[int]]::new()
                    $excludedNums = [System.Collections.Generic.HashSet[int]]::new()
                    $tokens = $spResponse -split ',' | ForEach-Object { $_.Trim() }

                    foreach ($token in $tokens) {
                        if ($token.ToLower() -eq 'all') {
                            for ($n = 1; $n -le $spList.Count; $n++) { [void]$selectedNums.Add($n) }
                        }
                        elseif ($token -match '^!(\d+)-(\d+)$') {
                            $s = [int]$Matches[1]; $e = [int]$Matches[2]
                            for ($n = $s; $n -le $e; $n++) { [void]$excludedNums.Add($n) }
                        }
                        elseif ($token -match '^!(\d+)$') {
                            [void]$excludedNums.Add([int]$Matches[1])
                        }
                        elseif ($token -match '^(\d+)-(\d+)$') {
                            $s = [int]$Matches[1]; $e = [int]$Matches[2]
                            for ($n = $s; $n -le $e; $n++) { [void]$selectedNums.Add($n) }
                        }
                        elseif ($token -match '^\d+$') {
                            [void]$selectedNums.Add([int]$token)
                        }
                    }

                    foreach ($ex in $excludedNums) { [void]$selectedNums.Remove($ex) }

                    $includedZips = @()
                    $includedSize = [long]0
                    $excludedSize = [long]0

                    foreach ($sp in $spList) {
                        if ($selectedNums.Contains($sp.Index)) {
                            $includedZips += $sp.File
                            $includedSize += $sp.ExtractedSize
                        }
                        else {
                            $excludedSize += $sp.ExtractedSize
                            $skippedExtracted += $sp.ExtractedSize
                        }
                    }

                    $finalSpZips = $includedZips

                    Write-Host ""
                    Write-Host "  Including $($includedZips.Count) system(s) (~$(Format-FileSize $includedSize))" -ForegroundColor Green
                    if ($excludedSize -gt 0) {
                        Write-Host "  Skipping $($spList.Count - $includedZips.Count) system(s) -- saving ~$(Format-FileSize $excludedSize)" -ForegroundColor DarkGray
                    }
                }
            }
            else {
                # All systems
                $finalSpZips = $spZips
                Write-OK "Extracting ALL system packs"
            }

            # Now separate Daphne from the rest and handle it with game selection
            $daphneSpZips  = @($finalSpZips | Where-Object { $_.Name -match '^Daphne\b' })
            $regularSpZips = @($finalSpZips | Where-Object { $_.Name -notmatch '^Daphne\b' })

            # Add the regular (non-Daphne) system packs
            if ($regularSpZips.Count -gt 0) {
                $selectedSources += [PSCustomObject]@{
                    Name = $dir.Name
                    Zips = $regularSpZips
                }
            }

            # Handle Daphne with game selection
            foreach ($dz in $daphneSpZips) {
                Write-Host ""
                Write-Host "  Scanning " -NoNewline
                Write-Host "Daphne Laserdisc Games" -ForegroundColor Yellow -NoNewline
                Write-Host " zip..." -ForegroundColor Gray

                $gameMap = Get-DaphneGameMap -ZipPath $dz.FullName
                $allGameNames = @($gameMap.Games.Keys | Sort-Object)
                $totalGameSize = [long]0
                foreach ($gn in $allGameNames) { $totalGameSize += $gameMap.Games[$gn] }
                $totalWithShared = $totalGameSize + $gameMap.SharedSize

                # Calculate DL + Space Ace size
                $dlsaSize = [long]0
                $dlsaCount = 0
                foreach ($dg in $script:DaphneDLSpaceAce) {
                    if ($gameMap.Games.ContainsKey($dg)) {
                        $dlsaSize += $gameMap.Games[$dg]
                        $dlsaCount++
                    }
                }
                $dlsaWithShared = $dlsaSize + $gameMap.SharedSize

                # Calculate popular size
                $popularSize = [long]0
                $popularCount = 0
                foreach ($pg in $script:DaphnePopularGames) {
                    if ($gameMap.Games.ContainsKey($pg)) {
                        $popularSize += $gameMap.Games[$pg]
                        $popularCount++
                    }
                }
                $popularWithShared = $popularSize + $gameMap.SharedSize

                Write-Host ""
                Write-Host "  Daphne Laserdisc Games" -ForegroundColor Yellow
                Write-Host "    $($allGameNames.Count) games, ~$(Format-FileSize $dz.Length) compressed, ~$(Format-FileSize $totalWithShared) extracted" -ForegroundColor Gray
                Write-Host ""
                Write-Host "    How would you like to handle it?" -ForegroundColor White
                Write-Host ""
                Write-Host "    [1]  Include ALL $($allGameNames.Count) games           (~$(Format-FileSize $totalWithShared))" -ForegroundColor White
                Write-Host "    [2]  Dragon's Lair + Space Ace only  (~$(Format-FileSize $dlsaWithShared)) -- saves ~$(Format-FileSize ($totalWithShared - $dlsaWithShared))" -ForegroundColor White
                Write-Host "    [3]  Popular games ($popularCount games)       (~$(Format-FileSize $popularWithShared)) -- saves ~$(Format-FileSize ($totalWithShared - $popularWithShared))" -ForegroundColor White
                Write-Host "    [4]  Let me pick which games to include" -ForegroundColor White
                Write-Host "    [N]  Skip Daphne entirely            -- saves ~$(Format-FileSize $totalWithShared)" -ForegroundColor White
                Write-Host ""
                $daphneChoice = Read-Host "  Choice"

                $selectedGames = $null  # null = all games (no filter)
                $daphneExtracted = $totalWithShared

                switch ($daphneChoice.ToLower()) {
                    '1' {
                        $selectedGames = $null
                        $daphneExtracted = $totalWithShared
                        Write-Host "    + Including ALL Daphne games" -ForegroundColor Green
                    }
                    '2' {
                        $selectedGames = @($script:DaphneDLSpaceAce | Where-Object { $gameMap.Games.ContainsKey($_) })
                        $daphneExtracted = $dlsaWithShared
                        $saved = $totalWithShared - $dlsaWithShared
                        $gameNames = ($selectedGames | ForEach-Object {
                            if ($script:DaphneGameNames.ContainsKey($_)) { $script:DaphneGameNames[$_] } else { $_ }
                        }) -join ', '
                        Write-Host "    + Including: $gameNames" -ForegroundColor Green
                        Write-Host "    + Saving ~$(Format-FileSize $saved)" -ForegroundColor Green
                    }
                    '3' {
                        $selectedGames = @($script:DaphnePopularGames | Where-Object { $gameMap.Games.ContainsKey($_) })
                        $daphneExtracted = $popularWithShared
                        $saved = $totalWithShared - $popularWithShared
                        Write-Host "    + Including $($selectedGames.Count) popular games (~$(Format-FileSize $daphneExtracted)), saving ~$(Format-FileSize $saved)" -ForegroundColor Green
                    }
                    '4' {
                        Write-Host ""
                        Write-Host "    Games list (popular titles marked with *):" -ForegroundColor Gray
                        Write-Host ""

                        $gameList = @()
                        $gi = 0
                        foreach ($gn in $allGameNames) {
                            if ($gameMap.Games[$gn] -lt 1MB) { continue }
                            $gi++
                            $isPop = $script:DaphnePopularGames -contains $gn
                            $marker = if ($isPop) { "*" } else { " " }
                            $sz = Format-FileSize $gameMap.Games[$gn]
                            $displayName = if ($script:DaphneGameNames.ContainsKey($gn)) { $script:DaphneGameNames[$gn] } else { $gn }
                            $gameList += [PSCustomObject]@{ Index = $gi; Name = $gn; DisplayName = $displayName; Size = $gameMap.Games[$gn]; Popular = $isPop }
                            Write-Host ("    [{0,2}]{1} {2,-30} ~{3}" -f $gi, $marker, $displayName, $sz) -ForegroundColor White
                        }

                        Write-Host ""
                        Write-Host "    Enter game numbers to INCLUDE (e.g. 1,2,3 or 1-5 or 'popular')" -ForegroundColor Yellow
                        Write-Host "    Type 'all' for everything, 'popular' for the starred ones" -ForegroundColor Gray
                        Write-Host ""
                        $gameResponse = Read-Host "  Include"

                        if ($gameResponse.ToLower() -eq 'all') {
                            $selectedGames = $null
                            $daphneExtracted = $totalWithShared
                            Write-Host "    + Including ALL games" -ForegroundColor Green
                        }
                        elseif ($gameResponse.ToLower() -eq 'popular') {
                            $selectedGames = @($script:DaphnePopularGames | Where-Object { $gameMap.Games.ContainsKey($_) })
                            $daphneExtracted = $popularWithShared
                            Write-Host "    + Including $($selectedGames.Count) popular games" -ForegroundColor Green
                        }
                        else {
                            $toInclude = [System.Collections.Generic.HashSet[int]]::new()
                            $tokens = $gameResponse -split ',' | ForEach-Object { $_.Trim() }
                            foreach ($token in $tokens) {
                                if ($token -match '^(\d+)-(\d+)$') {
                                    $s = [int]$Matches[1]; $e = [int]$Matches[2]
                                    for ($n = $s; $n -le $e; $n++) { [void]$toInclude.Add($n) }
                                }
                                elseif ($token -match '^\d+$') {
                                    [void]$toInclude.Add([int]$token)
                                }
                            }

                            $selectedGames = @()
                            $customSize = [long]0
                            foreach ($gl in $gameList) {
                                if ($toInclude.Contains($gl.Index)) {
                                    $selectedGames += $gl.Name
                                    $customSize += $gl.Size
                                }
                            }
                            $daphneExtracted = $customSize + $gameMap.SharedSize
                            Write-Host "    + Including $($selectedGames.Count) games (~$(Format-FileSize $daphneExtracted))" -ForegroundColor Green
                        }
                    }
                    'n' {
                        $skippedExtracted += $totalWithShared
                        Write-Host "    - Skipped Daphne entirely" -ForegroundColor DarkGray
                        continue
                    }
                    default {
                        $selectedGames = $null
                        $daphneExtracted = $totalWithShared
                        Write-Host "    + Including ALL Daphne games (default)" -ForegroundColor Green
                    }
                }

                # Add to selected sources with Daphne metadata
                $selectedSources += [PSCustomObject]@{
                    Name            = $dir.Name
                    Zips            = @($dz)
                    DaphneFilter    = $selectedGames
                    DaphneGameMap   = $gameMap
                    DaphneExtracted = $daphneExtracted
                }
            }
        }
    }
    elseif ($isSystemPacks) {
        # Not customizing -- include all system packs
        $selectedSources += [PSCustomObject]@{
            Name = $dir.Name
            Zips = $info.Zips
        }
    }
    elseif ($isOptional -and $customizeExtraction) {
        $include = Ask-Optional -Label $dir.Name -ZipCount $info.Count -Size $info.Size -ExtractedSize $info.ExtractedSize
        if ($include) {
            $selectedSources += [PSCustomObject]@{
                Name = $dir.Name
                Zips = $info.Zips
            }
            Write-Host "    + Included" -ForegroundColor Green
        }
        else {
            $skippedExtracted += $info.ExtractedSize
            Write-Host "    - Skipped" -ForegroundColor DarkGray
        }
    }
    elseif ($isOptional) {
        # Not customizing -- include all optional content
        $selectedSources += [PSCustomObject]@{
            Name = $dir.Name
            Zips = $info.Zips
        }
    }
    else {
        # Required folder -- but check for optional subfolders (like screensavers)
        # First, gather zips directly in this folder (always required)
        $directZips = @(Get-ChildItem -Path $dir.FullName -Filter '*.zip' -File -ErrorAction SilentlyContinue)

        # Separate Daphne zip(s) from regular zips for special handling
        $daphneZips = @($directZips | Where-Object { $_.Name -match '^Daphne\b' })
        $regularZips = @($directZips | Where-Object { $_.Name -notmatch '^Daphne\b' })

        if ($regularZips.Count -gt 0) {
            $selectedSources += [PSCustomObject]@{
                Name = $dir.Name
                Zips = $regularZips
            }
        }

        if (-not $customizeExtraction) {
            # Not customizing -- include ALL Daphne zips and ALL subfolders
            if ($daphneZips.Count -gt 0) {
                foreach ($dz in $daphneZips) {
                    $selectedSources += [PSCustomObject]@{
                        Name  = $dir.Name
                        Zips  = @($dz)
                    }
                }
            }

            $subFolders = Get-ChildItem -Path $dir.FullName -Directory -ErrorAction SilentlyContinue
            foreach ($sf in $subFolders) {
                $sfZips = @(Get-ChildItem -Path $sf.FullName -Filter '*.zip' -File -ErrorAction SilentlyContinue)
                if ($sfZips.Count -gt 0) {
                    $selectedSources += [PSCustomObject]@{
                        Name = "$($dir.Name)\$($sf.Name)"
                        Zips = $sfZips
                    }
                }
            }
        }
        else {
        # Customizing -- prompt for Daphne and optional subfolders

        # Handle Daphne zip(s) with game selection
        foreach ($dz in $daphneZips) {
            Write-Host ""
            Write-Host "  Scanning " -NoNewline
            Write-Host "Daphne Laserdisc Games" -ForegroundColor Yellow -NoNewline
            Write-Host " zip..." -ForegroundColor Gray

            $gameMap = Get-DaphneGameMap -ZipPath $dz.FullName
            $allGameNames = @($gameMap.Games.Keys | Sort-Object)
            $totalGameSize = [long]0
            foreach ($gn in $allGameNames) { $totalGameSize += $gameMap.Games[$gn] }
            $totalWithShared = $totalGameSize + $gameMap.SharedSize

            # Calculate DL + Space Ace size
            $dlsaSize = [long]0
            $dlsaCount = 0
            foreach ($dg in $script:DaphneDLSpaceAce) {
                if ($gameMap.Games.ContainsKey($dg)) {
                    $dlsaSize += $gameMap.Games[$dg]
                    $dlsaCount++
                }
            }
            $dlsaWithShared = $dlsaSize + $gameMap.SharedSize

            # Calculate popular size
            $popularSize = [long]0
            $popularCount = 0
            foreach ($pg in $script:DaphnePopularGames) {
                if ($gameMap.Games.ContainsKey($pg)) {
                    $popularSize += $gameMap.Games[$pg]
                    $popularCount++
                }
            }
            $popularWithShared = $popularSize + $gameMap.SharedSize

            Write-Host ""
            Write-Host "  Daphne Laserdisc Games" -ForegroundColor Yellow
            Write-Host "    $($allGameNames.Count) games, ~$(Format-FileSize $dz.Length) compressed, ~$(Format-FileSize $totalWithShared) extracted" -ForegroundColor Gray
            Write-Host ""
            Write-Host "    How would you like to handle it?" -ForegroundColor White
            Write-Host ""
            Write-Host "    [1]  Include ALL $($allGameNames.Count) games           (~$(Format-FileSize $totalWithShared))" -ForegroundColor White
            Write-Host "    [2]  Dragon's Lair + Space Ace only  (~$(Format-FileSize $dlsaWithShared)) -- saves ~$(Format-FileSize ($totalWithShared - $dlsaWithShared))" -ForegroundColor White
            Write-Host "    [3]  Popular games ($popularCount games)       (~$(Format-FileSize $popularWithShared)) -- saves ~$(Format-FileSize ($totalWithShared - $popularWithShared))" -ForegroundColor White
            Write-Host "    [4]  Let me pick which games to include" -ForegroundColor White
            Write-Host "    [N]  Skip Daphne entirely            -- saves ~$(Format-FileSize $totalWithShared)" -ForegroundColor White
            Write-Host ""
            $daphneChoice = Read-Host "  Choice"

            $selectedGames = $null  # null = all games (no filter)
            $daphneExtracted = $totalWithShared

            switch ($daphneChoice.ToLower()) {
                '1' {
                    # All games
                    $selectedGames = $null
                    $daphneExtracted = $totalWithShared
                    Write-Host "    + Including ALL Daphne games" -ForegroundColor Green
                }
                '2' {
                    # Dragon's Lair + Space Ace only
                    $selectedGames = @($script:DaphneDLSpaceAce | Where-Object { $gameMap.Games.ContainsKey($_) })
                    $daphneExtracted = $dlsaWithShared
                    $saved = $totalWithShared - $dlsaWithShared
                    $gameNames = ($selectedGames | ForEach-Object {
                        if ($script:DaphneGameNames.ContainsKey($_)) { $script:DaphneGameNames[$_] } else { $_ }
                    }) -join ', '
                    Write-Host "    + Including: $gameNames" -ForegroundColor Green
                    Write-Host "    + Saving ~$(Format-FileSize $saved)" -ForegroundColor Green
                }
                '3' {
                    # Popular only
                    $selectedGames = @($script:DaphnePopularGames | Where-Object { $gameMap.Games.ContainsKey($_) })
                    $daphneExtracted = $popularWithShared
                    $saved = $totalWithShared - $popularWithShared
                    Write-Host "    + Including $($selectedGames.Count) popular games (~$(Format-FileSize $daphneExtracted)), saving ~$(Format-FileSize $saved)" -ForegroundColor Green
                }
                '4' {
                    # Custom selection
                    Write-Host ""
                    Write-Host "    Games list (popular titles marked with *):" -ForegroundColor Gray
                    Write-Host ""

                    $gameList = @()
                    $gi = 0
                    foreach ($gn in $allGameNames) {
                        # Skip tiny entries (artwork-only aliases with 0 actual game data)
                        if ($gameMap.Games[$gn] -lt 1MB) { continue }
                        $gi++
                        $isPop = $script:DaphnePopularGames -contains $gn
                        $marker = if ($isPop) { "*" } else { " " }
                        $sz = Format-FileSize $gameMap.Games[$gn]
                        $displayName = if ($script:DaphneGameNames.ContainsKey($gn)) { $script:DaphneGameNames[$gn] } else { $gn }
                        $gameList += [PSCustomObject]@{ Index = $gi; Name = $gn; DisplayName = $displayName; Size = $gameMap.Games[$gn]; Popular = $isPop }
                        Write-Host ("    [{0,2}]{1} {2,-30} ~{3}" -f $gi, $marker, $displayName, $sz) -ForegroundColor White
                    }

                    Write-Host ""
                    Write-Host "    Enter game numbers to INCLUDE (e.g. 1,2,3 or 1-5 or 'popular')" -ForegroundColor Yellow
                    Write-Host "    Type 'all' for everything, 'popular' for the starred ones" -ForegroundColor Gray
                    Write-Host ""
                    $gameResponse = Read-Host "  Include"

                    if ($gameResponse.ToLower() -eq 'all') {
                        $selectedGames = $null
                        $daphneExtracted = $totalWithShared
                        Write-Host "    + Including ALL games" -ForegroundColor Green
                    }
                    elseif ($gameResponse.ToLower() -eq 'popular') {
                        $selectedGames = @($script:DaphnePopularGames | Where-Object { $gameMap.Games.ContainsKey($_) })
                        $daphneExtracted = $popularWithShared
                        Write-Host "    + Including $($selectedGames.Count) popular games" -ForegroundColor Green
                    }
                    else {
                        # Parse number selection
                        $toInclude = [System.Collections.Generic.HashSet[int]]::new()
                        $tokens = $gameResponse -split ',' | ForEach-Object { $_.Trim() }
                        foreach ($token in $tokens) {
                            if ($token -match '^(\d+)-(\d+)$') {
                                $s = [int]$Matches[1]; $e = [int]$Matches[2]
                                for ($n = $s; $n -le $e; $n++) { [void]$toInclude.Add($n) }
                            }
                            elseif ($token -match '^\d+$') {
                                [void]$toInclude.Add([int]$token)
                            }
                        }

                        $selectedGames = @()
                        $customSize = [long]0
                        foreach ($gl in $gameList) {
                            if ($toInclude.Contains($gl.Index)) {
                                $selectedGames += $gl.Name
                                $customSize += $gl.Size
                            }
                        }
                        $daphneExtracted = $customSize + $gameMap.SharedSize
                        Write-Host "    + Including $($selectedGames.Count) games (~$(Format-FileSize $daphneExtracted))" -ForegroundColor Green
                    }
                }
                'n' {
                    # Skip entirely
                    $skippedExtracted += $totalWithShared
                    Write-Host "    - Skipped Daphne entirely" -ForegroundColor DarkGray
                    continue
                }
                default {
                    # Treat unknown input as ALL
                    $selectedGames = $null
                    $daphneExtracted = $totalWithShared
                    Write-Host "    + Including ALL Daphne games (default)" -ForegroundColor Green
                }
            }

            # Add to selected sources with Daphne metadata
            $selectedSources += [PSCustomObject]@{
                Name           = $dir.Name
                Zips           = @($dz)
                DaphneFilter   = $selectedGames
                DaphneGameMap  = $gameMap
                DaphneExtracted = $daphneExtracted
            }
        }

        # Then check each subfolder
        $subFolders = Get-ChildItem -Path $dir.FullName -Directory -ErrorAction SilentlyContinue
        foreach ($sf in $subFolders) {
            $sfZips = @(Get-ChildItem -Path $sf.FullName -Filter '*.zip' -File -ErrorAction SilentlyContinue)
            if ($sfZips.Count -eq 0) { continue }

            $sfIsOptional = $false
            foreach ($kw in $optionalSubfolderKeywords) {
                if ($sf.Name -match $kw) { $sfIsOptional = $true; break }
            }

            if ($sfIsOptional) {
                $sfSize = ($sfZips | Measure-Object -Property Length -Sum).Sum
                if (-not $sfSize) { $sfSize = 0 }
                $sfExtracted = [long]0
                foreach ($sfz in $sfZips) {
                    $sfExtracted += Get-ZipExtractedSize -ZipPath $sfz.FullName
                }
                $include = Ask-Optional -Label $sf.Name -ZipCount $sfZips.Count -Size $sfSize -ExtractedSize $sfExtracted
                if ($include) {
                    $selectedSources += [PSCustomObject]@{
                        Name = "$($dir.Name)\$($sf.Name)"
                        Zips = $sfZips
                    }
                    Write-Host "    + Included" -ForegroundColor Green
                }
                else {
                    $skippedExtracted += $sfExtracted
                    Write-Host "    - Skipped" -ForegroundColor DarkGray
                }
            }
            else {
                $selectedSources += [PSCustomObject]@{
                    Name = "$($dir.Name)\$($sf.Name)"
                    Zips = $sfZips
                }
            }
        }
        } # end customizeExtraction else block for required folders
    }
}

if ($selectedSources.Count -eq 0) {
    Write-Host ""
    Write-Host "  No zip files to extract. Exiting." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "  Folders to extract:" -ForegroundColor Green
foreach ($ss in $selectedSources) {
    $cnt = $ss.Zips.Count
    $sz = ($ss.Zips | Measure-Object -Property Length -Sum).Sum
    $extra = ""
    if ($ss.PSObject.Properties['DaphneFilter'] -and $ss.DaphneFilter -ne $null) {
        $extra = " [$($ss.DaphneFilter.Count) games selected]"
    }
    elseif ($ss.PSObject.Properties['DaphneExtracted'] -and $ss.DaphneFilter -eq $null) {
        $extra = " [all games]"
    }
    Write-Host "    + $($ss.Name)  ($cnt zips, ~$(Format-FileSize $sz))$extra" -ForegroundColor White
}
if ($skippedExtracted -gt 0) {
    Write-Host ""
    Write-Host "  Skipped content: ~$(Format-FileSize $skippedExtracted) saved" -ForegroundColor DarkGray
}

# =================================================================================
#  STEP 4: GATHER ZIPS AND CHECK SPACE
# =================================================================================

Write-Header "STEP 4: REVIEW AND SPACE CHECK"

Write-Host ""
Write-Host "  Reading actual extracted sizes from zip headers..." -ForegroundColor Gray

# Build flat list of all zips from selected sources (with real extracted sizes)
$allZips = [System.Collections.ArrayList]::new()
foreach ($ss in $selectedSources) {
    foreach ($z in $ss.Zips) {
        # For Daphne zips, use the pre-calculated filtered size
        $isDaphne = ($z.Name -match '^Daphne\b')
        if ($isDaphne -and $ss.PSObject.Properties['DaphneExtracted']) {
            $extractedSz = $ss.DaphneExtracted
        }
        else {
            $extractedSz = Get-ZipExtractedSize -ZipPath $z.FullName
        }

        $zipObj = [PSCustomObject]@{
            Name          = $z.Name
            FullPath      = $z.FullName
            Size          = $z.Length
            ExtractedSize = $extractedSz
            Source        = $ss.Name
            Excluded      = $false
            DaphneFilter  = $null
        }

        # Attach Daphne filter if applicable
        if ($isDaphne -and $ss.PSObject.Properties['DaphneFilter']) {
            $zipObj.DaphneFilter = $ss.DaphneFilter
        }

        [void]$allZips.Add($zipObj)
    }
}

# Get current free space on the destination drive
$destFree = 0
$destQualifier = Split-Path -Qualifier $outputPath -ErrorAction SilentlyContinue
if ($destQualifier) {
    $dl = $destQualifier.TrimEnd(':')
    $psd = Get-PSDrive -Name $dl -ErrorAction SilentlyContinue
    if ($psd) { $destFree = $psd.Free }
}

# Check if destination already has a previous build (overwritten files = reclaimed space)
$existingSize = 0
$knownBuildFolders = @('content', 'appdata', 'base_assets', 'docs')
$hasPreviousBuild = $false
foreach ($kbf in $knownBuildFolders) {
    if (Test-Path (Join-Path $outputPath $kbf)) { $hasPreviousBuild = $true; break }
}
if ($hasPreviousBuild) {
    Write-Host ""
    Write-Host "  Scanning destination for existing build..." -ForegroundColor Gray
    foreach ($kbf in $knownBuildFolders) {
        $kbfPath = Join-Path $outputPath $kbf
        if (Test-Path $kbfPath) {
            $kbfSize = (Get-ChildItem -Path $kbfPath -Recurse -File -ErrorAction SilentlyContinue |
                        Measure-Object -Property Length -Sum).Sum
            if ($kbfSize) { $existingSize += $kbfSize }
        }
    }
}

$effectiveFree = $destFree + $existingSize

# --- Space check loop: if not enough space, let user exclude large zips ---
$spaceOk = $false

while (-not $spaceOk) {
    $includedZips = $allZips | Where-Object { -not $_.Excluded }
    $totalZipSize = ($includedZips | Measure-Object -Property Size -Sum).Sum
    $totalExtracted = [long]($includedZips | Measure-Object -Property ExtractedSize -Sum).Sum

    Write-Host ""
    Write-Host "  Zip files:       $($includedZips.Count) files" -ForegroundColor White
    Write-Host "  Compressed size: ~$(Format-FileSize $totalZipSize)" -ForegroundColor White
    Write-Host "  Extracted size:  ~$(Format-FileSize $totalExtracted)  (actual from zip headers)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Destination:     $outputPath" -ForegroundColor White
    Write-Host "  Free space:      $(Format-FileSize $destFree)" -ForegroundColor White

    if ($existingSize -gt 0) {
        Write-Host "  Existing build:  ~$(Format-FileSize $existingSize)  (will be overwritten)" -ForegroundColor Gray
        Write-Host "  Effective free:  ~$(Format-FileSize $effectiveFree)  (free + overwritten)" -ForegroundColor White
    }

    if ($effectiveFree -gt 0 -and $totalExtracted -gt $effectiveFree) {
        $shortfall = $totalExtracted - $effectiveFree
        Write-Host ""
        Write-Host "  *** NOT ENOUGH SPACE ***" -ForegroundColor Red
        Write-Host "  You need to free up about ~$(Format-FileSize $shortfall)" -ForegroundColor Red
        Write-Host ""
        Write-Host "  You can exclude some of the largest system packs to make it fit." -ForegroundColor Yellow
        Write-Host "  Here are the biggest ones you can remove:" -ForegroundColor Yellow
        Write-Host ""

        # Show the top 15 largest included zips (skip base/key folders -- those are required)
        $removable = $includedZips |
            Where-Object { $_.Source -notmatch 'Base Build|Key Folders' } |
            Sort-Object Size -Descending |
            Select-Object -First 15

        $idx = 0
        $removeList = @()
        foreach ($r in $removable) {
            $idx++
            $extSz = Format-FileSize $r.ExtractedSize
            $removeList += [PSCustomObject]@{ Index = $idx; Zip = $r }
            Write-Host "    [$idx]  $($r.Name)  (~$extSz extracted)" -ForegroundColor White
        }

        Write-Host ""
        Write-Host "  Enter numbers to exclude (e.g. 1,2,3 or 1-5)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Or type " -NoNewline -ForegroundColor Gray
        Write-Host "go" -NoNewline -ForegroundColor Green
        Write-Host " to extract anyway (ignore space warning)" -ForegroundColor Gray
        Write-Host "  Or type " -NoNewline -ForegroundColor Gray
        Write-Host "quit" -NoNewline -ForegroundColor Red
        Write-Host " to exit" -ForegroundColor Gray
        Write-Host ""
        $trimResponse = Read-Host "  Choice"

        if ($trimResponse.ToLower() -eq 'quit') {
            Write-Host "  Exiting." -ForegroundColor Yellow
            exit 0
        }
        elseif ($trimResponse.ToLower() -eq 'go' -or $trimResponse.ToLower() -eq 'done') {
            Write-Host ""
            Write-Host "  Proceeding without enough space -- good luck!" -ForegroundColor Yellow
            $spaceOk = $true
        }
        else {
            # Parse selection
            $toExclude = [System.Collections.Generic.HashSet[int]]::new()
            $tokens = $trimResponse -split ',' | ForEach-Object { $_.Trim() }
            foreach ($token in $tokens) {
                if ($token -match '^(\d+)-(\d+)$') {
                    $s = [int]$Matches[1]; $e = [int]$Matches[2]
                    for ($n = $s; $n -le $e; $n++) { [void]$toExclude.Add($n) }
                }
                elseif ($token -match '^\d+$') {
                    [void]$toExclude.Add([int]$token)
                }
            }

            $removedExtracted = [long]0
            foreach ($rl in $removeList) {
                if ($toExclude.Contains($rl.Index)) {
                    $rl.Zip.Excluded = $true
                    $removedExtracted += $rl.Zip.ExtractedSize
                    Write-Host "    - Excluded: $($rl.Zip.Name)" -ForegroundColor DarkGray
                }
            }
            Write-Host ""
            Write-Host "  Saved ~$(Format-FileSize $removedExtracted). Recalculating..." -ForegroundColor Green
        }
    }
    else {
        # We have enough space
        if ($effectiveFree -gt 0) {
            $headroom = $effectiveFree - $totalExtracted
            Write-Host "  Headroom:        ~$(Format-FileSize $headroom) to spare" -ForegroundColor Green
        }
        $spaceOk = $true
    }
}

# Final list of zips to extract
$allZips = [System.Collections.ArrayList]@($allZips | Where-Object { -not $_.Excluded })

if ($allZips.Count -eq 0) {
    Write-Host ""
    Write-Host "  Nothing left to extract. Exiting." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "  Files to extract ($($allZips.Count)):" -ForegroundColor Gray
foreach ($z in $allZips) {
    $note = ""
    if ($z.DaphneFilter -ne $null -and $z.DaphneFilter.Count -gt 0) {
        $note = "  [$($z.DaphneFilter.Count) games, ~$(Format-FileSize $z.ExtractedSize) extracted]"
    }
    Write-Host "    $($z.Source)\$($z.Name)  ($(Format-FileSize $z.Size))$note" -ForegroundColor DarkGray
}

Write-Host ""
$confirm = Read-Host "  Ready to extract? (y/n)"

if ($confirm.ToLower() -ne 'y') {
    Write-Host ""
    Write-Host "  Cancelled." -ForegroundColor Yellow
    exit 0
}

# =================================================================================
#  STEP 5: EXTRACT
# =================================================================================

Write-Header "STEP 5: EXTRACTING"

# -- Resume support ---------------------------------------------------------------
$progressFile = Join-Path $outputPath '.extraction-progress'
$completedZips = @{}
$resuming = $false

if (Test-Path $progressFile) {
    $prevCompleted = @(Get-Content $progressFile -ErrorAction SilentlyContinue | Where-Object { $_.Trim() -ne '' })
    if ($prevCompleted.Count -gt 0) {
        Write-Host ""
        Write-Host "  Previous extraction detected!" -ForegroundColor Yellow
        Write-Host "  $($prevCompleted.Count) of $($allZips.Count) zips were completed last time." -ForegroundColor Gray
        Write-Host ""
        $resumeAnswer = Read-Host "  Resume where you left off? (y/n)"
        if ($resumeAnswer.ToLower() -eq 'y') {
            $resuming = $true
            foreach ($cz in $prevCompleted) { $completedZips[$cz] = $true }
            Write-Host "  Resuming -- skipping $($completedZips.Count) completed zips" -ForegroundColor Green
        }
        else {
            Remove-Item $progressFile -Force -ErrorAction SilentlyContinue
            Write-Host "  Starting fresh" -ForegroundColor Gray
        }
    }
}

$extracted = 0
$skipped = 0
$failed = 0
$totalCount = $allZips.Count
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Wrap extraction in try/finally so Ctrl+C still leaves a valid progress file
try {
    foreach ($z in $allZips) {
        $extracted++
        $pct = [math]::Round(($extracted / $totalCount) * 100)

        # Skip if already completed in a previous run
        if ($completedZips.ContainsKey($z.Name)) {
            $skipped++
            Write-Host ""
            Write-Host "  [$extracted/$totalCount] ($pct%) " -ForegroundColor DarkGray -NoNewline
            Write-Host "$($z.Name)" -ForegroundColor DarkGray -NoNewline
            Write-Host "  -- already completed, skipping" -ForegroundColor DarkGray
            continue
        }

        Write-Host ""
        Write-Host "  [$extracted/$totalCount] ($pct%) " -ForegroundColor Cyan -NoNewline
        Write-Host "$($z.Name)" -NoNewline
        Write-Host "  ($(Format-FileSize $z.Size))" -ForegroundColor Gray

        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($z.FullPath)
            $entryCount = $zip.Entries.Count
            $processedEntries = 0
            $extractedEntries = 0
            $skippedEntries = 0

            # Build Daphne game filter set if applicable
            $daphneFilterSet = $null
            if ($z.DaphneFilter -ne $null -and $z.DaphneFilter.Count -gt 0) {
                $daphneFilterSet = [System.Collections.Generic.HashSet[string]]::new(
                    [string[]]$z.DaphneFilter,
                    [System.StringComparer]::OrdinalIgnoreCase
                )
                Write-Host "    (filtering to $($daphneFilterSet.Count) selected games)" -ForegroundColor Gray
            }

            foreach ($entry in $zip.Entries) {
                $processedEntries++

                if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }

                # Apply Daphne filter: skip entries for non-selected games
                if ($daphneFilterSet) {
                    $entryGame = $null
                    $ep = $entry.FullName

                    if ($ep -match 'assets/(singe|mpegs)/([^/]+)/') { $entryGame = $Matches[2] }
                    elseif ($ep -match 'roms/([^/]+)\.(zip|singe)$') { $entryGame = $Matches[1] }
                    elseif ($ep -match 'medium_artwork/[^/]+/([^/.]+)') { $entryGame = $Matches[1] }

                    # If we identified a game name and it's NOT in our filter, skip it
                    if ($entryGame -and -not $daphneFilterSet.Contains($entryGame)) {
                        $skippedEntries++
                        continue
                    }
                }

                $destFile = Join-Path $outputPath $entry.FullName

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

                $extractedEntries++

                if ($entryCount -gt 100 -and $processedEntries % 500 -eq 0) {
                    $entryPct = [math]::Round(($processedEntries / $entryCount) * 100)
                    Write-Host "    ... $processedEntries / $entryCount entries ($entryPct%)" -ForegroundColor DarkGray
                }
            }

            $zip.Dispose()
            if ($skippedEntries -gt 0) {
                Write-Host "    Done ($extractedEntries files extracted, $skippedEntries skipped)" -ForegroundColor Green
            }
            else {
                Write-Host "    Done ($extractedEntries files)" -ForegroundColor Green
            }

            # Record this zip as completed for resume support
            $z.Name | Out-File -FilePath $progressFile -Append -Encoding utf8
            $completedZips[$z.Name] = $true
        }
        catch {
            $failed++
            Write-Host "    FAILED: $($_.Exception.Message)" -ForegroundColor Red
            if ($zip) { $zip.Dispose() }
        }
    }
}
finally {
    # If interrupted (Ctrl+C), the progress file already has all completed zips
    # The zip that was mid-extraction is NOT in the file, so it'll be re-extracted on resume
    if ($failed -gt 0 -or $extracted -lt $totalCount) {
        Write-Host ""
        Write-Host "  Progress saved to $progressFile" -ForegroundColor Yellow
        Write-Host "  Run the script again to resume from where you left off." -ForegroundColor Yellow
    }
}

$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed

# =================================================================================
#  DONE
# =================================================================================

Write-Header "COMPLETE"
Write-Host ""

# Clean up progress file on full success
$actualExtracted = $extracted - $skipped - $failed
if ($failed -eq 0) {
    Remove-Item $progressFile -Force -ErrorAction SilentlyContinue
    if ($skipped -gt 0) {
        Write-Host "  Extracted: $actualExtracted / $totalCount  ($skipped resumed from earlier)" -ForegroundColor Green
    }
    else {
        Write-Host "  Extracted: $actualExtracted / $totalCount  (all successful)" -ForegroundColor Green
    }
}
else {
    Write-Host "  Extracted: $actualExtracted / $totalCount  ($failed failed)" -ForegroundColor Yellow
    Write-Host "  Run the script again to retry failed items." -ForegroundColor Yellow
}

$mins = [math]::Floor($elapsed.TotalMinutes)
$secs = $elapsed.Seconds
$timeStr = if ($mins -gt 0) { "$mins min $secs sec" } else { "$($elapsed.TotalSeconds.ToString('N0')) sec" }
Write-Host "  Time:      $timeStr" -ForegroundColor White

Write-Host ""
Write-Host "  Output:    $outputPath" -ForegroundColor Green

# Final size check
$outputSize = (Get-ChildItem -Path $outputPath -Recurse -File -ErrorAction SilentlyContinue |
               Measure-Object -Property Length -Sum).Sum
$outputFileCount = (Get-ChildItem -Path $outputPath -Recurse -File -ErrorAction SilentlyContinue).Count

Write-Host "  Size:      ~$(Format-FileSize $outputSize)" -ForegroundColor White
Write-Host "  Files:     $outputFileCount" -ForegroundColor White
Write-Host ""
Write-Host "  Ready to go!" -ForegroundColor Green
Write-Host ""
