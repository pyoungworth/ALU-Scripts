<#
.SYNOPSIS
    Analyzes directories for duplicate/versioned files, with interactive cleanup.

.DESCRIPTION
    Scans each subdirectory for files with multiple versions
    (e.g., "App v2.0b2.zip" and "App v2.0b3.zip").
    Shows old vs new, and lets you pick what to delete.

.PARAMETER Path
    The root directory to scan. Defaults to the script's own directory.

.EXAMPLE
    .\Analyze-Duplicates.ps1 -Path "F:\MyFiles"
    .\Analyze-Duplicates.ps1

.NOTES
    Author: Philip Youngworth
    Project: ALU-Scripts (AwesomeSauce2 Utility Scripts)
    License: MIT
#>

param(
    [Parameter(Position = 0)]
    [string]$Path
)

if (-not $Path) {
    $Path = Split-Path -Parent $MyInvocation.MyCommand.Definition
}
$Path = (Resolve-Path $Path).Path

# -- Helpers ----------------------------------------------------------------------

function Parse-VersionedName {
    param([string]$Name)

    if ($Name -match '^(.+?)\s+v(\d+)\.(\d+)b(\d+)(.*)$') {
        return [PSCustomObject]@{
            BaseName      = $Matches[1].Trim()
            Extension     = $Matches[5].Trim()
            VersionString = "v$($Matches[2]).$($Matches[3])b$($Matches[4])"
            SortKey       = ([int]$Matches[2]) * 100000 + ([int]$Matches[3]) * 10000 + ([int]$Matches[4])
            HasVersion    = $true
        }
    }

    if ($Name -match '^(.+?)\s+v(\d+(?:\.\d+)+)(.*)$') {
        $parts = $Matches[2] -split '\.'
        $sortKey = 0
        for ($i = 0; $i -lt $parts.Count; $i++) {
            $sortKey += [int]$parts[$i] * [math]::Pow(10000, ($parts.Count - 1 - $i))
        }
        return [PSCustomObject]@{
            BaseName      = $Matches[1].Trim()
            Extension     = $Matches[3].Trim()
            VersionString = "v$($Matches[2])"
            SortKey       = $sortKey
            HasVersion    = $true
        }
    }

    if ($Name -match '^(.+?)\s+(\d{4}-\d{2}-\d{2})(.*)$') {
        $dateVal = [datetime]::ParseExact($Matches[2], 'yyyy-MM-dd', $null)
        return [PSCustomObject]@{
            BaseName      = $Matches[1].Trim()
            Extension     = $Matches[3].Trim()
            VersionString = $Matches[2]
            SortKey       = $dateVal.Year * 10000 + $dateVal.Month * 100 + $dateVal.Day
            HasVersion    = $true
        }
    }

    $ext = [System.IO.Path]::GetExtension($Name)
    $base = if ($ext) { $Name.Substring(0, $Name.Length - $ext.Length) } else { $Name }
    return [PSCustomObject]@{
        BaseName      = $base.Trim()
        Extension     = $ext
        VersionString = "(no version)"
        SortKey       = -1
        HasVersion    = $false
    }
}

function Normalize-BaseName {
    param([string]$Name)
    return ($Name -replace '[\s\-_]', '').ToLower()
}

function Format-FileSize {
    param([long]$Bytes)
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

# -- Scan -------------------------------------------------------------------------

Write-Header "DUPLICATE & VERSION ANALYZER"
Write-Host "  Scanning: $Path" -ForegroundColor Gray
Write-Host ""

# Gather all items (files and folders) from each subdirectory
$allItems = @()
$subDirs = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

foreach ($dir in $subDirs) {
    foreach ($child in (Get-ChildItem -Path $dir.FullName -ErrorAction SilentlyContinue)) {
        $allItems += [PSCustomObject]@{
            Name         = $child.Name
            FullPath     = $child.FullName
            ParentDir    = $dir.Name
            IsDirectory  = $child.PSIsContainer
            Size         = if (-not $child.PSIsContainer) { $child.Length } else { 0 }
            LastModified = $child.LastWriteTime
        }
    }
}

Write-Host "  Found $($allItems.Count) items across $($subDirs.Count) directories." -ForegroundColor Gray

# -- Analysis ---------------------------------------------------------------------
# For each directory, group items by normalized base name.
# Within each group, show old vs new and flag old items (including any
# extracted folders whose version is no longer the newest).

$findings = [System.Collections.ArrayList]::new()
$findingId = 0

$groupedByDir = $allItems | Group-Object -Property ParentDir
$anyFound = $false

foreach ($dirGroup in $groupedByDir) {
    # Parse every item
    $parsed = foreach ($item in $dirGroup.Group) {
        $vInfo = Parse-VersionedName $item.Name
        [PSCustomObject]@{
            Item          = $item
            ParsedInfo    = $vInfo
            NormalizedKey = (Normalize-BaseName $vInfo.BaseName) + "|" + (Normalize-BaseName $vInfo.Extension)
        }
    }

    # Also create a key variant without extension so folders and zips group together
    # e.g. "SNES v2.0b2" (folder) and "SNES v2.0b2.zip" share base "SNES"
    $parsedWithBaseKey = foreach ($p in $parsed) {
        $baseOnly = Normalize-BaseName $p.ParsedInfo.BaseName
        [PSCustomObject]@{
            Item          = $p.Item
            ParsedInfo    = $p.ParsedInfo
            NormalizedKey = $p.NormalizedKey
            BaseKey       = $baseOnly
        }
    }

    # Group by BaseKey to catch folders + zips of the same thing
    $baseGroups = $parsedWithBaseKey | Group-Object -Property BaseKey | Where-Object { $_.Count -gt 1 }

    foreach ($bg in $baseGroups) {
        $members = $bg.Group | Sort-Object { $_.ParsedInfo.SortKey }

        # Need at least one versioned item
        $anyVersioned = $members | Where-Object { $_.ParsedInfo.HasVersion }
        if (-not $anyVersioned) { continue }

        # Find the highest version in this group
        $highestVersion = ($members | Where-Object { $_.ParsedInfo.HasVersion } |
                           Sort-Object { $_.ParsedInfo.SortKey } | Select-Object -Last 1).ParsedInfo.SortKey

        # Separate into: newest-version items vs older items
        $newestItems = @($members | Where-Object { $_.ParsedInfo.SortKey -eq $highestVersion })
        $olderItems  = @($members | Where-Object { $_.ParsedInfo.SortKey -ne $highestVersion })

        # Skip if nothing is older (no duplicates)
        if ($olderItems.Count -eq 0) { continue }

        $anyFound = $true
        $newestVer = $newestItems[0].ParsedInfo.VersionString

        Write-Host ""
        Write-Host "  --- $($dirGroup.Name) | $($newestItems[0].ParsedInfo.BaseName) ---" -ForegroundColor Yellow

        # Show what's OLD
        foreach ($entry in $olderItems) {
            $item = $entry.Item
            $ver = $entry.ParsedInfo.VersionString

            # If it's a folder, calculate its size
            $size = $item.Size
            if ($item.IsDirectory) {
                $size = (Get-ChildItem -Path $item.FullPath -Recurse -File -ErrorAction SilentlyContinue |
                         Measure-Object -Property Length -Sum).Sum
                if (-not $size) { $size = 0 }
            }

            $findingId++
            [void]$findings.Add([PSCustomObject]@{
                Id          = $findingId
                FilePath    = $item.FullPath
                FileName    = $item.Name
                Size        = $size
                IsDirectory = $item.IsDirectory
            })

            $typeTag = if ($item.IsDirectory) { "folder" } else { "file" }
            Write-Host "    [#$($findingId.ToString().PadLeft(3))] " -ForegroundColor Red -NoNewline
            Write-Host "OLD  " -ForegroundColor Red -NoNewline
            Write-Host "$($item.Name)  ($ver, $(Format-FileSize $size), $typeTag)" -ForegroundColor Gray
        }

        # Show what's NEW (keep)
        foreach ($entry in $newestItems) {
            $item = $entry.Item
            $size = $item.Size
            if ($item.IsDirectory) {
                $size = (Get-ChildItem -Path $item.FullPath -Recurse -File -ErrorAction SilentlyContinue |
                         Measure-Object -Property Length -Sum).Sum
                if (-not $size) { $size = 0 }
            }
            $typeTag = if ($item.IsDirectory) { "folder" } else { "file" }
            Write-Host "    [KEEP]   " -ForegroundColor Green -NoNewline
            Write-Host "NEW  " -ForegroundColor Green -NoNewline
            Write-Host "$($item.Name)  ($newestVer, $(Format-FileSize $size), $typeTag)" -ForegroundColor Gray
        }
    }
}

if (-not $anyFound) {
    Write-Host ""
    Write-Host "  No duplicates or version conflicts found. Your files are clean!" -ForegroundColor Green
    exit 0
}

# -- Summary & Deletion -----------------------------------------------------------

Write-Header "SUMMARY"

$totalReclaimable = ($findings | Measure-Object -Property Size -Sum).Sum

Write-Host ""
Write-Host "  $($findings.Count) old items found.  Potential savings: ~$(Format-FileSize $totalReclaimable)" -ForegroundColor White
Write-Host ""
Write-Host ("  " + ("-" * 61)) -ForegroundColor DarkGray

foreach ($f in $findings) {
    $sizeStr = if ($f.Size -gt 0) { "  ($(Format-FileSize $f.Size))" } else { "" }
    Write-Host "  #$($f.Id.ToString().PadLeft(3))  $($f.FileName)$sizeStr" -ForegroundColor Gray
}

Write-Host ("  " + ("-" * 61)) -ForegroundColor DarkGray
Write-Host ""
Write-Host "  WHAT TO DELETE:" -ForegroundColor Yellow
Write-Host "    Numbers:   1,3,5       Range:  1-5" -ForegroundColor Gray
Write-Host "    All:       all         Exclude: all,!3,!7" -ForegroundColor Gray
Write-Host "    Skip:      none  (or just press Enter)" -ForegroundColor Gray
Write-Host ""

$response = Read-Host "  Select items to DELETE"

if (-not $response -or $response.Trim().ToLower() -eq 'none') {
    Write-Host ""
    Write-Host "  No files deleted. Exiting." -ForegroundColor Green
    exit 0
}

# Parse selection
$selectedIds = [System.Collections.Generic.HashSet[int]]::new()
$excludedIds = [System.Collections.Generic.HashSet[int]]::new()
$tokens = $response -split ',' | ForEach-Object { $_.Trim() }

foreach ($token in $tokens) {
    if ($token.ToLower() -eq 'all') {
        foreach ($f in $findings) { [void]$selectedIds.Add($f.Id) }
    }
    elseif ($token -match '^!(\d+)-(\d+)$') {
        $s = [int]$Matches[1]; $e = [int]$Matches[2]
        for ($n = $s; $n -le $e; $n++) { [void]$excludedIds.Add($n) }
    }
    elseif ($token -match '^!(\d+)$') {
        [void]$excludedIds.Add([int]$Matches[1])
    }
    elseif ($token -match '^(\d+)-(\d+)$') {
        $s = [int]$Matches[1]; $e = [int]$Matches[2]
        for ($n = $s; $n -le $e; $n++) { [void]$selectedIds.Add($n) }
    }
    elseif ($token -match '^\d+$') {
        [void]$selectedIds.Add([int]$token)
    }
}

foreach ($ex in $excludedIds) { [void]$selectedIds.Remove($ex) }
$toDelete = $findings | Where-Object { $selectedIds.Contains($_.Id) }

if ($toDelete.Count -eq 0) {
    Write-Host ""
    Write-Host "  No valid items selected. Exiting." -ForegroundColor Yellow
    exit 0
}

# Confirm
Write-Host ""
Write-Host "  About to DELETE $($toDelete.Count) items:" -ForegroundColor Red
Write-Host ""

$deletionSize = 0
foreach ($item in $toDelete) {
    Write-Host "    #$($item.Id.ToString().PadLeft(3))  $($item.FileName)" -ForegroundColor Red
    $deletionSize += $item.Size
}

Write-Host ""
Write-Host "  Space to reclaim: ~$(Format-FileSize $deletionSize)" -ForegroundColor White
Write-Host ""
$confirm = Read-Host "  Type 'YES' to confirm (this cannot be undone)"

if ($confirm -ne 'YES') {
    Write-Host ""
    Write-Host "  Cancelled. Nothing was deleted." -ForegroundColor Yellow
    exit 0
}

# Delete
Write-Host ""
$deleted = 0; $failed = 0

foreach ($item in $toDelete) {
    try {
        if ($item.IsDirectory) {
            Remove-Item -Path $item.FilePath -Recurse -Force -ErrorAction Stop
        }
        else {
            Remove-Item -Path $item.FilePath -Force -ErrorAction Stop
        }
        Write-Host "  [DELETED] $($item.FileName)" -ForegroundColor Green
        $deleted++
    }
    catch {
        Write-Host "  [FAILED]  $($item.FileName) - $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
$skipped = $findings.Count - $toDelete.Count
Write-Host "  Done! Deleted: $deleted  |  Failed: $failed  |  Skipped: $skipped" -ForegroundColor Cyan
Write-Host ""
