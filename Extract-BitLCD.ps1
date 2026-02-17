<#
.SYNOPSIS
    Extracts BitLCD marquee packs to a USB drive for the ALU BitLCD display.

.DESCRIPTION
    Walks you through:
      1. Where are the downloaded files (finds the __BitLCD Marquees folder)
      2. Where to extract to (pick a USB drive -- FAT32 or exFAT recommended)
      3. Extract all marquee packs or pick specific ones
      4. Reviews everything and checks free space before starting
      5. Extracts into \bitlcd\thirdparty\OneSauce\ on the USB drive

    The BitLCD USB is a SEPARATE drive from the main OneSauce drive.

.PARAMETER Path
    The path to the __BitLCD Marquees folder. If not provided, the script
    will try to find it from the last download location or guide you.

.EXAMPLE
    .\Extract-BitLCD.ps1
    .\Extract-BitLCD.ps1 -Path "F:\AS2\__BitLCD Marquees"

.NOTES
    Author: Philip Youngworth
    Project: ALU-Scripts (AwesomeSauce2 Utility Scripts)
    License: MIT
#>

param(
    [Parameter(Position = 0)]
    [string]$Path
)

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

function Get-ZipExtractedSize {
    param([string]$ZipPath)
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $total = [long]($archive.Entries | Measure-Object -Property Length -Sum).Sum
        $archive.Dispose()
        return $total
    }
    catch {
        return [long]((Get-Item $ZipPath).Length * 1.05)
    }
}

function Get-ZipTopFolder {
    param([string]$ZipPath)
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $firstEntry = $archive.Entries | Select-Object -First 1
        $folder = $null
        if ($firstEntry) {
            $parts = $firstEntry.FullName -split '[/\\]'
            if ($parts.Count -gt 1) { $folder = $parts[0] }
        }
        $archive.Dispose()
        return $folder
    }
    catch { return $null }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# =================================================================================
#  STEP 1: FIND THE BITLCD MARQUEES FOLDER
# =================================================================================

Write-Header "STEP 1: WHERE ARE THE BITLCD MARQUEE FILES?"

if ($Path) {
    if (-not (Test-Path $Path)) {
        Write-Fail "Path not found: $Path"
        exit 1
    }
    $Path = (Resolve-Path $Path).Path
}
else {
    # Try to find __BitLCD Marquees from the last download location
    $savedPathFile = Join-Path $scriptDir '.last-download-path'
    $foundPath = $null

    if (Test-Path $savedPathFile) {
        $savedPath = (Get-Content $savedPathFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
        if ($savedPath -and (Test-Path $savedPath)) {
            # Look for __BitLCD Marquees in the saved download path
            $marqueeDir = Get-ChildItem -Path $savedPath -Directory -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match 'BitLCD|Marquee' } |
                          Select-Object -First 1
            if ($marqueeDir) {
                Write-Host ""
                Write-Host "  Found marquee folder: " -NoNewline -ForegroundColor White
                Write-Host "$($marqueeDir.FullName)" -ForegroundColor Green
                Write-Host ""
                $useSaved = Read-Host "  Use this folder? (y/n, q to quit)"
                Test-Quit $useSaved
                if ($useSaved.ToLower() -ne 'n') {
                    $foundPath = $marqueeDir.FullName
                }
            }
        }
    }

    if (-not $foundPath) {
        Write-Host ""
        Write-Host "  Select the drive where you downloaded the files:" -ForegroundColor White
        Write-Host ""

        # Build drive list
        $partitions = Get-Partition -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
        $disks = Get-Disk -ErrorAction SilentlyContinue
        $diskLookup = @{}
        foreach ($d in $disks) { $diskLookup[$d.Number] = $d }

        $diskInfoByLetter = @{}
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

        $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
                  Where-Object { $_.Free -gt 0 -and $_.Root } |
                  Sort-Object Name

        $idx = 0
        $driveList = @()

        foreach ($drv in $drives) {
            $idx++
            $letter = $drv.Name
            $root   = $drv.Root
            $freeSpace = Format-FileSize $drv.Free
            $totalSpace = Format-FileSize ($drv.Used + $drv.Free)

            $vol = Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue
            $drvLabel = if ($vol -and $vol.FileSystemLabel) { $vol.FileSystemLabel } else { "" }

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
            $details = @()
            if ($busTag)   { $details += $busTag }
            if ($drvLabel) { $details += $drvLabel }
            if ($details.Count -gt 0) { $detailStr = "  [$($details -join ' - ')]" }

            Write-Host "    [$idx]  $label$detailStr  ($freeSpace free of $totalSpace)" -ForegroundColor White

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
                if ($chosen) { $selectedDrive = $chosen }
                else { Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow }
            }
            else { Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow }
        }

        # Look for marquee folder on the selected drive
        $searchPaths = @($selectedDrive.Root)
        # Also check common subfolder patterns
        $subDirs = Get-ChildItem -Path $selectedDrive.Root -Directory -ErrorAction SilentlyContinue
        foreach ($sd in $subDirs) {
            $marqueeDir = Get-ChildItem -Path $sd.FullName -Directory -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match 'BitLCD|Marquee' } |
                          Select-Object -First 1
            if ($marqueeDir) {
                $foundPath = $marqueeDir.FullName
                break
            }
        }

        # Also check the root itself
        if (-not $foundPath) {
            $marqueeDir = Get-ChildItem -Path $selectedDrive.Root -Directory -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match 'BitLCD|Marquee' } |
                          Select-Object -First 1
            if ($marqueeDir) {
                $foundPath = $marqueeDir.FullName
            }
        }

        if (-not $foundPath) {
            Write-Fail "Could not find a BitLCD Marquees folder on $($selectedDrive.Root)"
            Write-Host "  Look for a folder named '__BitLCD Marquees' in your download location." -ForegroundColor Yellow
            exit 1
        }

        Write-Host ""
        Write-Host "  Found: " -NoNewline -ForegroundColor White
        Write-Host "$foundPath" -ForegroundColor Green
    }

    $Path = $foundPath
}

# Validate the folder has zip files
$allZipFiles = @(Get-ChildItem -Path $Path -Filter '*.zip' -File -ErrorAction SilentlyContinue | Sort-Object Name)
if ($allZipFiles.Count -eq 0) {
    Write-Fail "No zip files found in $Path"
    exit 1
}

Write-Host ""
Write-OK "Source: $Path"
Write-Info "  Found $($allZipFiles.Count) zip file(s)"

# =================================================================================
#  DEDUPLICATE: Group by internal folder, pick one per group
# =================================================================================

Write-Host ""
Write-Info "Scanning zip contents for duplicates..."

$folderGroups = @{}
foreach ($z in $allZipFiles) {
    $topFolder = Get-ZipTopFolder -ZipPath $z.FullName
    $key = if ($topFolder) { $topFolder } else { $z.Name }

    if (-not $folderGroups.ContainsKey($key)) {
        $folderGroups[$key] = @()
    }
    $folderGroups[$key] += $z
}

# Pick the best zip from each group (prefer newer version, then larger file)
$uniquePacks = @()
foreach ($key in ($folderGroups.Keys | Sort-Object)) {
    $group = $folderGroups[$key]
    if ($group.Count -eq 1) {
        $uniquePacks += $group[0]
    }
    else {
        # Pick the largest file (usually the same, but covers edge cases)
        $best = $group | Sort-Object Length -Descending | Select-Object -First 1
        $uniquePacks += $best
    }
}

$dupeCount = $allZipFiles.Count - $uniquePacks.Count
if ($dupeCount -gt 0) {
    Write-Info "  $dupeCount duplicate(s) removed, $($uniquePacks.Count) unique marquee pack(s)"
}
else {
    Write-Info "  $($uniquePacks.Count) marquee pack(s) found (no duplicates)"
}

# =================================================================================
#  STEP 2: WHERE IS THE BITLCD USB DRIVE?
# =================================================================================

Write-Header "STEP 2: SELECT THE BITLCD USB DRIVE"
Write-Host ""
Write-Host "  This should be a SEPARATE USB drive from your OneSauce drive." -ForegroundColor Yellow
Write-Host "  FAT32 or exFAT is recommended for the BitLCD USB." -ForegroundColor Gray
Write-Host ""

# Build drive list
$partitions = Get-Partition -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
$disks = Get-Disk -ErrorAction SilentlyContinue
$diskLookup = @{}
foreach ($d in $disks) { $diskLookup[$d.Number] = $d }

$diskInfoByLetter = @{}
foreach ($p in $partitions) {
    $dl = [string]$p.DriveLetter
    $dk = $diskLookup[$p.DiskNumber]
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

$drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
          Where-Object { $_.Free -gt 0 -and $_.Root } |
          Sort-Object Name

$idx = 0
$driveList = @()

foreach ($drv in $drives) {
    $idx++
    $letter = $drv.Name
    $root   = $drv.Root
    $freeSpace = Format-FileSize $drv.Free
    $totalSpace = Format-FileSize ($drv.Used + $drv.Free)

    $vol = Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue
    $drvLabel = if ($vol -and $vol.FileSystemLabel) { $vol.FileSystemLabel } else { "" }

    $busTag = ""
    $modelTag = ""
    $fsType = ""
    $di = $diskInfoByLetter[[string]$letter]
    if ($di) {
        switch ($di.BusType) {
            'USB'   { $busTag = "USB" }
            'NVMe'  { $busTag = "NVMe" }
            'SATA'  { $busTag = "SATA" }
            'SAS'   { $busTag = "SAS" }
            default { if ($di.BusType) { $busTag = $di.BusType } }
        }
        if ($di.Model) { $modelTag = $di.Model }
        if ($di.FileSystem) { $fsType = $di.FileSystem }
    }

    $details = @()
    if ($busTag)   { $details += $busTag }
    if ($drvLabel) { $details += $drvLabel }
    if ($modelTag) { $details += $modelTag }
    $detailStr = ""
    if ($details.Count -gt 0) { $detailStr = "  [$($details -join ' - ')]" }

    $fsTag = ""
    if ($fsType) { $fsTag = "  ($fsType)" }

    # Check for OneSauce to warn against using the same drive
    $hasOneSauce = Test-Path (Join-Path $root 'OneSauce')
    $warnStr = if ($hasOneSauce) { "  <-- OneSauce drive (don't use this one)" } else { "" }

    Write-Host "    [$idx]  ${letter}:\$detailStr$fsTag  ($freeSpace free of $totalSpace)" -NoNewline -ForegroundColor White
    if ($hasOneSauce) {
        Write-Host $warnStr -ForegroundColor Red
    }
    else {
        Write-Host ""
    }

    $driveList += [PSCustomObject]@{
        Index      = $idx
        Letter     = $letter
        Root       = $root
        Free       = $drv.Free
        FileSystem = $fsType
        HasOneSauce = $hasOneSauce
    }
}

Write-Host ""
$destDrive = $null
while (-not $destDrive) {
    $driveChoice = Read-Host "  Pick a drive (enter number, q to quit)"
    Test-Quit $driveChoice

    if ($driveChoice -match '^\d+$') {
        $chosen = $driveList | Where-Object { $_.Index -eq [int]$driveChoice }
        if ($chosen) {
            # Warn if this is the OneSauce drive
            if ($chosen.HasOneSauce) {
                Write-Host ""
                Write-Host "  Warning: This drive has OneSauce on it." -ForegroundColor Red
                Write-Host "  The BitLCD USB should be a SEPARATE drive from your OneSauce drive." -ForegroundColor Yellow
                Write-Host ""
                $confirm = Read-Host "  Are you sure you want to use this drive? (y/n)"
                if (-not $confirm -or $confirm.Trim().ToLower() -ne 'y') {
                    Write-Host ""
                    continue
                }
            }
            $destDrive = $chosen
        }
        else {
            Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  Invalid selection. Try again." -ForegroundColor Yellow
    }
}

$basePath = Join-Path $destDrive.Root 'bitlcd\thirdparty'

Write-Host ""
Write-Host "  Files will be extracted to: " -NoNewline -ForegroundColor White
Write-Host "$basePath\OneSauce" -ForegroundColor Green
Write-Host ""
Write-Host "  The default folder name is " -NoNewline -ForegroundColor Gray
Write-Host "OneSauce" -ForegroundColor White -NoNewline
Write-Host ". You can change it if you'd like," -ForegroundColor Gray
Write-Host "  but it will still go under bitlcd\thirdparty\ on the drive." -ForegroundColor Gray
Write-Host ""
$folderName = Read-Host "  Folder name (press Enter for 'OneSauce')"

if (-not $folderName -or $folderName.Trim() -eq '') {
    $folderName = 'OneSauce'
}
else {
    $folderName = $folderName.Trim()
}

$outputPath = Join-Path $basePath $folderName

Write-Host ""
Write-OK "Extracting to: $outputPath"

# =================================================================================
#  STEP 3: CONTENT SELECTION
# =================================================================================

Write-Header "STEP 3: CONTENT SELECTION"
Write-Host ""
Write-Host "    [A]  Extract ALL marquee packs (default)" -ForegroundColor Green
Write-Host "    [S]  Select specific packs" -ForegroundColor Yellow
Write-Host ""

$selChoice = Read-Host "  Choice (A/S, q to quit)"
Test-Quit $selChoice

$selectedPacks = @()

if ($selChoice.ToUpper() -eq 'S') {
    # Calculate sizes and show the list
    Write-Host ""
    Write-Host "  Calculating sizes..." -ForegroundColor Gray

    $packList = @()
    $pi = 0
    foreach ($z in $uniquePacks) {
        $pi++
        $extractedSize = Get-ZipExtractedSize -ZipPath $z.FullName
        $displayName = [System.IO.Path]::GetFileNameWithoutExtension($z.Name)
        $packList += [PSCustomObject]@{
            Index         = $pi
            File          = $z
            DisplayName   = $displayName
            ExtractedSize = $extractedSize
        }
    }

    Write-Host ""
    Write-Host "  $($packList.Count) marquee packs available:" -ForegroundColor White
    Write-Host ""

    # Display in two columns
    $nameWidth = 42
    $colCount = 2
    $rows = [math]::Ceiling($packList.Count / $colCount)

    for ($row = 0; $row -lt $rows; $row++) {
        $line = ""
        for ($col = 0; $col -lt $colCount; $col++) {
            $i = $col * $rows + $row
            if ($i -lt $packList.Count) {
                $pk = $packList[$i]
                $num = ($pk.Index).ToString().PadLeft(2)
                $name = $pk.DisplayName
                if ($name.Length -gt $nameWidth) { $name = $name.Substring(0, $nameWidth - 3) + "..." }
                $szText = Format-FileSize $pk.ExtractedSize
                $entry = "  [$num] $($name.PadRight($nameWidth)) $($szText.PadLeft(9))"
                $line += $entry
            }
        }
        Write-Host "  $line" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "  HOW TO SELECT:" -ForegroundColor Yellow
    Write-Host "    all             Extract everything (default)" -ForegroundColor Gray
    Write-Host "    1,3,5           Pick specific packs" -ForegroundColor Gray
    Write-Host "    1-20            Pick a range" -ForegroundColor Gray
    Write-Host "    all,!5,!12      Extract all EXCEPT #5 and #12" -ForegroundColor Gray
    Write-Host ""

    $selResponse = Read-Host "  Selection (Enter for all)"

    if (-not $selResponse -or $selResponse.Trim() -eq '' -or $selResponse.Trim().ToLower() -eq 'all') {
        if (-not ($selResponse -and $selResponse -match '!')) {
            $selectedPacks = $uniquePacks
            Write-OK "Extracting ALL marquee packs"
        }
    }

    if ($selResponse -and $selectedPacks.Count -eq 0) {
        $selectedNums = [System.Collections.Generic.HashSet[int]]::new()
        $excludedNums = [System.Collections.Generic.HashSet[int]]::new()
        $tokens = $selResponse -split ',' | ForEach-Object { $_.Trim() }

        foreach ($token in $tokens) {
            if ($token.ToLower() -eq 'all') {
                for ($n = 1; $n -le $packList.Count; $n++) { [void]$selectedNums.Add($n) }
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

        $includedSize = [long]0
        $skippedSize = [long]0
        foreach ($pk in $packList) {
            if ($selectedNums.Contains($pk.Index)) {
                $selectedPacks += $pk.File
                $includedSize += $pk.ExtractedSize
            }
            else {
                $skippedSize += $pk.ExtractedSize
            }
        }

        Write-Host ""
        Write-Host "  Including $($selectedPacks.Count) pack(s) (~$(Format-FileSize $includedSize))" -ForegroundColor Green
        if ($skippedSize -gt 0) {
            Write-Host "  Skipping $($packList.Count - $selectedPacks.Count) pack(s) -- saving ~$(Format-FileSize $skippedSize)" -ForegroundColor DarkGray
        }
    }

    if ($selectedPacks.Count -eq 0) {
        $selectedPacks = $uniquePacks
        Write-OK "Extracting ALL marquee packs (default)"
    }
}
else {
    $selectedPacks = $uniquePacks
    Write-OK "Extracting ALL marquee packs"
}

# =================================================================================
#  STEP 4: REVIEW AND SPACE CHECK
# =================================================================================

Write-Header "STEP 4: REVIEW AND SPACE CHECK"

Write-Host ""
Write-Host "  Reading extracted sizes from zip headers..." -ForegroundColor Gray

$totalCompressed = [long]0
$totalExtracted = [long]0
foreach ($z in $selectedPacks) {
    $totalCompressed += $z.Length
    $totalExtracted += (Get-ZipExtractedSize -ZipPath $z.FullName)
}

Write-Host ""
Write-Host "  Marquee packs: $($selectedPacks.Count) file(s)" -ForegroundColor White
Write-Host "  Compressed:    ~$(Format-FileSize $totalCompressed)" -ForegroundColor White
Write-Host "  Extracted:     ~$(Format-FileSize $totalExtracted)" -ForegroundColor White
Write-Host ""
Write-Host "  Destination:   $outputPath" -ForegroundColor Green
Write-Host "  Free space:    $(Format-FileSize $destDrive.Free)" -ForegroundColor White

$headroom = $destDrive.Free - $totalExtracted
if ($headroom -lt 0) {
    Write-Host ""
    Write-Fail "Not enough space! You need ~$(Format-FileSize ([math]::Abs($headroom))) more."
    Write-Host "  Try selecting fewer packs, or use a larger drive." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
else {
    Write-Host "  Headroom:      ~$(Format-FileSize $headroom) to spare" -ForegroundColor White
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

# Create output directory
if (-not (Test-Path $outputPath)) {
    New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
}

# Resume support
$progressFile = Join-Path $destDrive.Root '.bitlcd-extraction-progress'
$completedZips = @{}

if (Test-Path $progressFile) {
    $lines = Get-Content $progressFile -ErrorAction SilentlyContinue
    foreach ($line in $lines) {
        if ($line.Trim()) { $completedZips[$line.Trim()] = $true }
    }
    if ($completedZips.Count -gt 0) {
        Write-Host ""
        Write-Host "  Resuming from previous run ($($completedZips.Count) already completed)" -ForegroundColor Yellow
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$extracted = 0
$skipped = 0
$failed = 0

Write-Host ""
Write-Host "  Extracting to: $outputPath" -ForegroundColor Green
Write-Host ""

try {
    foreach ($z in $selectedPacks) {
        $extracted++

        if ($completedZips.ContainsKey($z.Name)) {
            $skipped++
            Write-Host "  [$extracted/$($selectedPacks.Count)]  $($z.Name)  -- already done, skipping" -ForegroundColor DarkGray
            continue
        }

        $szText = Format-FileSize $z.Length
        Write-Host "  [$extracted/$($selectedPacks.Count)]  $($z.Name)  ($szText)" -ForegroundColor White

        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($z.FullName)
            $entryCount = $zip.Entries.Count
            $processedEntries = 0
            $extractedEntries = 0

            foreach ($entry in $zip.Entries) {
                $processedEntries++
                if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }

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
            Write-Host "    Done ($extractedEntries files)" -ForegroundColor Green

            # Record as completed for resume support
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
    if ($failed -gt 0 -or $extracted -lt $selectedPacks.Count) {
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

$actualExtracted = $extracted - $skipped - $failed
if ($failed -eq 0) {
    Remove-Item $progressFile -Force -ErrorAction SilentlyContinue
    if ($skipped -gt 0) {
        Write-Host "  Extracted: $actualExtracted / $($selectedPacks.Count)  ($skipped resumed from earlier)" -ForegroundColor Green
    }
    else {
        Write-Host "  Extracted: $actualExtracted / $($selectedPacks.Count)  (all successful)" -ForegroundColor Green
    }
}
else {
    Write-Host "  Extracted: $actualExtracted / $($selectedPacks.Count)  ($failed failed)" -ForegroundColor Yellow
    Write-Host "  Run the script again to retry failed items." -ForegroundColor Yellow
}

$mins = [math]::Floor($elapsed.TotalMinutes)
$secs = $elapsed.Seconds
$timeStr = if ($mins -gt 0) { "$mins min $secs sec" } else { "$($elapsed.TotalSeconds.ToString('N0')) sec" }
Write-Host "  Time:      $timeStr" -ForegroundColor White

Write-Host ""
Write-Host "  Output:    $outputPath" -ForegroundColor Green

$outputSize = (Get-ChildItem -Path $outputPath -Recurse -File -ErrorAction SilentlyContinue |
               Measure-Object -Property Length -Sum).Sum
$outputFileCount = (Get-ChildItem -Path $outputPath -Recurse -File -ErrorAction SilentlyContinue).Count

Write-Host "  Size:      ~$(Format-FileSize $outputSize)" -ForegroundColor White
Write-Host "  Files:     $outputFileCount" -ForegroundColor White
Write-Host ""
Write-Host "  Plug this USB into the BitLCD port on your ALU and enjoy!" -ForegroundColor Green
Write-Host ""
