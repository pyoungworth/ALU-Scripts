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
.\Extract-To-USB.ps1 -Path "F:\AS2"
```

### Analyze-Duplicates.ps1

Scans directories for duplicate/versioned files (e.g., `App v2.0b2.zip` and `App v2.0b3.zip`). Identifies old versions and offers interactive cleanup.

```powershell
.\Analyze-Duplicates.ps1
.\Analyze-Duplicates.ps1 -Path "F:\MyFiles"
```

## Requirements

- **Windows 10** (version 1809+) or **Windows 11**
- **PowerShell 5.1+** (comes with Windows)
- **winget** (comes with Windows, used to install rclone if needed)
- Access to the RetroFE ALU Google Shared Drive (for the download script)

## Quick Start

1. **Download the files from Google Drive:**

   ```powershell
   .\Download-From-GDrive.ps1
   ```

2. **Extract to USB:**

   ```powershell
   .\Extract-To-USB.ps1
   ```

3. **Plug the USB into your Legends Ultimate and enjoy!**

## Running Scripts for the First Time

If PowerShell blocks the script with a security warning, run this once first:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Or right-click the script, select Properties, and check "Unblock" at the bottom.

## Community

Join the [AwesomeSauce Facebook group](https://www.facebook.com/groups/2118196125010827) for discussion, help, and updates related to the AS2 RetroFE build and ALU modding.

## Author

**Philip Youngworth**

## License

MIT License -- see [LICENSE](LICENSE) for details.
