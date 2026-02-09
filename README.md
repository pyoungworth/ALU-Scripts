# ALU-Scripts

PowerShell utility scripts for managing AwesomeSauce2 (AS2) RetroFE builds for the Legends Ultimate (ALU).

These scripts are designed to be as user-friendly as possible -- just run them and follow the prompts. No technical knowledge required.

## Scripts

### Download-From-GDrive.ps1

Downloads all build files from Google Drive. Handles everything automatically:

- Installs [rclone](https://rclone.org/) if not already installed
- Walks you through Google account sign-in (opens browser)
- Auto-detects the RetroFE ALU shared drive
- Lets you pick where to download and optionally limit bandwidth
- **Resumable** -- if interrupted, just run again and it picks up where it left off

```powershell
.\Download-From-GDrive.ps1
```

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

Windows blocks scripts by default for security. You need to allow them once. Copy and paste this into PowerShell and press **Enter**:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Type **Y** and press **Enter** when it asks to confirm. You only need to do this once.

> **Alternative:** Right-click each `.ps1` script file, select **Properties**, and check **Unblock** at the bottom.

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
