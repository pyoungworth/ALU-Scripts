# ALU-Scripts

PowerShell utility scripts for managing AwesomeSauce2 (AS2) RetroFE builds for the Legends Ultimate (ALU).

These scripts are designed to be as user-friendly as possible -- just run them and follow the prompts. No technical knowledge required.

## Scripts

### Download-From-GDrive.ps1

Downloads build files from Google Drive. Handles everything automatically:

- Installs [rclone](https://rclone.org/) if not already installed
- Walks you through Google account sign-in (opens browser)
- Auto-detects the RetroFE ALU shared drive and build folder
- Lets you pick where to download and optionally limit bandwidth
- **Choose what to download** -- download everything, or customize:
  - BitLCD Marquees (with size shown)
  - Optional Themes (with size shown)
  - Screensaver (with size shown)
  - System Packs -- download all 76 systems, or pick specific ones (sizes shown for each)
- **Resumable** -- if interrupted, just run again and it picks up where it left off

```powershell
.\Download-From-GDrive.ps1
```

<details>
<summary>Example output (click to expand)</summary>

```
  ============================================================
  GOOGLE DRIVE DOWNLOAD TOOL
  ============================================================

  This script will help you download all the files from
  Google Drive. Just follow the prompts!

  What this will do:
    1. Install rclone (a download tool) if needed
    2. Connect to Google Drive (you'll sign in via browser)
    3. Pick where to save the files
    4. Optionally limit download speed
    5. Choose optional content (marquees, themes, system packs)
    6. Download everything

  If the download gets interrupted, just run this script
  again -- it will pick up where it left off!

  Ready to start? (y/n, q to quit): y

================================================================================
  STEP 1: CHECKING FOR RCLONE
================================================================================
  [OK] rclone is already installed

================================================================================
  STEP 2: GOOGLE DRIVE CONNECTION
================================================================================
  [OK] Google Drive connection 'RetroFE' already exists
  [OK] Connection is working!
  [OK] Build folder: ___One saUCE 2.0

================================================================================
  STEP 3: WHERE DO YOU WANT TO DOWNLOAD TO?
================================================================================

  You'll need at least 1 TB of free space for the full download.

    [1]  C:\ drive  [NVMe - Windows]  (1.40 TB free of 1.86 TB)
    [2]  D:\ drive  [USB - External Drive]  (5.08 TB free of 9.10 TB)

  Pick a drive (enter number, q to quit): 2

  What do you want to name the download folder?

  Examples: AwesomeSauce2, RetroFE-Download, AS2

  Folder name (q to quit): AS2

  Download destination: D:\AS2

================================================================================
  STEP 4: DOWNLOAD SPEED
================================================================================

  The download will use your full internet speed by default.
  If you want to browse the web or stream video while downloading,
  you can limit the download speed.

    [1]  Full speed (use all bandwidth)
    [2]  Limit speed (save some for other use)

  Choice (1, 2, or q to quit): 1
  [OK] Using full download speed

================================================================================
  STEP 5: CONTENT SELECTION
================================================================================

    [A]  Download EVERYTHING (default)
    [C]  Customize what to download

  Choice (A/C, q to quit): C

  Checking sizes of optional content...

  Include BitLCD Marquees?  (~35.87 GB)
  (LCD marquee images for supported games)

  (y/n, default: y, q to quit): y
    + Including BitLCD Marquees

  Include Optional Themes?  (~13.84 GB)
  (Additional visual themes for RetroFE)

  (y/n, default: y, q to quit): n
    - Skipping Optional Themes

  Include Screensaver?  (~247.51 GB)
  (Attract-mode screensaver videos)

  (y/n, default: y, q to quit): n
    - Skipping screensaver

  Which System Packs do you want?
  (e.g. Arcade, NES, SNES, Daphne, etc)

    [A]  Download ALL systems (default)
    [S]  Let me pick which systems I want

  Choice (A/S, q to quit): S

  76 system packs available (925.11 GB total):

    [ 1] AmstradCPC v2.0b2                3.26 GB  [39] NintendoFamicon v2.0b2          1.38 GB
    [ 2] AmstradGX4000 v2.0b2            86.56 MB  [40] NintendoFamiconDiskSystem ...  827.13 MB
    [ 3] Arcade v2.0b3                   15.09 GB  [41] NintendoGame&Watch v2.0b2      166.81 MB
    ...

  HOW TO SELECT:
    all             Download everything (default)
    1,3,5           Pick specific systems
    1-20            Pick a range
    all,!5,!12      Download all EXCEPT #5 and #12
    all,!30-40      Download all EXCEPT #30 through #40

  Selection (Enter for all): 3,19,57,64

  Including 4 system(s):
    + Arcade v2.0b3  (15.09 GB)
    + Daphne v2.0b3  (85.23 GB)
    + SegaGenesis v2.0b2  (3.80 GB)
    + SNES v2.0b2  (70.75 GB)

  Skipping 72 system(s) -- saving 750.24 GB

================================================================================
  STEP 6: DOWNLOADING FILES
================================================================================

  Source:      Google Drive (RetroFE) / ___One saUCE 2.0
  Destination: D:\AS2
  Excluding:   Optional Themes, Screensaver, 72 system pack(s)

  Start downloading? (y/n, q to quit): y

  Starting download...
```

</details>

### Extract-To-USB.ps1

Extracts downloaded zip files into a build folder ready to copy to USB. Features:

- Guided drive/folder selection with free space info
- Optional content prompts (BitLCD Marquees, screensaver, themes)
- **Daphne laserdisc game selection** -- choose all games, popular titles only, Dragon's Lair + Space Ace only, or hand-pick individual games
- Smart space checking with trim-to-fit (exclude large system packs if short on space)
- Accurate extracted sizes from zip headers (no guessing)
- **Resumable** -- tracks progress so interrupted extractions can continue

```powershell
.\Extract-To-USB.ps1
.\Extract-To-USB.ps1 -Path "D:\AS2"
```

### Analyze-Duplicates.ps1

Scans directories for duplicate/versioned files (e.g., `App v2.0b2.zip` and `App v2.0b3.zip`). Identifies old versions and offers interactive cleanup.

```powershell
.\Analyze-Duplicates.ps1
.\Analyze-Duplicates.ps1 -Path "D:\MyFiles"
```

## Requirements

- **Windows 10** (version 1809+) or **Windows 11**
- **PowerShell 5.1+** (comes with Windows)
- **winget** (comes with Windows, used to install rclone if needed)
- Access to the RetroFE ALU Google Shared Drive (for the download script)

## Getting Started (Step by Step)

If you've never used PowerShell before, don't worry -- just follow these steps.

### 1. Download These Scripts

- Click the green **Code** button at the top of this page, then click **Download ZIP**
- Extract the zip to a folder you'll remember (e.g., `C:\ALU-Scripts`)

### 2. Open PowerShell

There are two easy ways:

**Option A -- Open PowerShell directly in the folder (recommended):**

1. Open the folder where you extracted the scripts in File Explorer
2. Click in the **address bar** at the top (where it shows the folder path)
3. Type `powershell` and press **Enter**
4. A blue PowerShell window will open, already in the right folder

**Option B -- Open PowerShell from the Start Menu:**

1. Click the **Start** button and type `PowerShell`
2. Click **Windows PowerShell** (the blue icon -- not "PowerShell ISE")
3. Navigate to your scripts folder by typing:

```powershell
cd C:\ALU-Scripts
```

(Replace `C:\ALU-Scripts` with wherever you extracted the files)

### 3. Allow Scripts to Run (One-Time Setup)

Windows blocks downloaded scripts in two ways, and you need to handle **both**:

**Step A -- Set the execution policy:**

Copy and paste this into PowerShell and press **Enter**:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Type **Y** and press **Enter** when it asks to confirm. You only need to do this once.

**Step B -- Unblock the downloaded files:**

Since these scripts were downloaded from the internet, Windows marks them as blocked. Run this in the same PowerShell window:

```powershell
Get-ChildItem -Path . -Filter *.ps1 | Unblock-File
```

Or you can do it manually: right-click each `.ps1` file, select **Properties**, check **Unblock** at the bottom, and click **OK**.

### 4. Run the Scripts

Now you're ready! Just type the script name and press **Enter**:

```powershell
.\Download-From-GDrive.ps1
```

The script will walk you through everything from there -- just follow the prompts.

Once the download finishes, extract to USB:

```powershell
.\Extract-To-USB.ps1
```

### 5. Plug the USB into your Legends Ultimate and enjoy!

## Community

Join the [AwesomeSauce Facebook group](https://www.facebook.com/groups/2118196125010827) for discussion, help, and updates related to the AS2 RetroFE build and ALU modding.

## Author

**Philip Youngworth**

## License

MIT License -- see [LICENSE](LICENSE) for details.
