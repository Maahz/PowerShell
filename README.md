# Shortcut Hub (shortcutter.ps1)

Shortcut Hub is a small PowerShell GUI that helps you create Desktop shortcuts to commonly used applications and built-in Windows tools. It's designed so the script can later be hosted and executed via a one-liner (for example: `irm "https://example.com/shortcutter.ps1" | iex`).

> Security note: executing remote scripts via `iex` is risky. Prefer downloading and inspecting the script, verifying a hash/signature, or hosting the script in a place you control.

## Quick start

- Run locally:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Markku\source\PowerShell\shortcutter.ps1"
```

- Or (when you host the raw script) stream-and-run (not recommended without verifying):

```powershell
irm "https://yourhost.example/shortcutter.ps1" | iex
```

## What it does

- Shows a GUI grouped by categories (Browsers, Security, Gaming, Communication, Tools, Windows & System).
- Only shows buttons for items that appear to exist on the machine.
- Clicking a button creates a shortcut on the current user's Desktop. The shortcut supports executable files, MMC snap-ins (.msc), Control Panel applets (.cpl) and ms-settings: URIs.

## Supported programs and tools

The following programs and Windows built-ins are included in the current `shortcutter.ps1` catalog. The script checks for installed/existing paths and only shows buttons for items detected on the current machine.
### Full catalog (alphabetical)

The catalog in `shortcutter.ps1` has been expanded. Below is the full, alphabetized list of supported items (the script only shows those detected on the current machine):

- 1Password
## Supported programs and tools

The following programs and Windows built-ins are included in the current `shortcutter.ps1` catalog. The script checks for installed/existing paths and only shows buttons for items detected on the current machine.

### Catalog (grouped by category)

To make the catalog easier to read the supported items are grouped by category below. The script still only shows buttons for items that are actually detected on the current machine.
- Tor Browser
- Chromium

#### Security & Passwords

- Malwarebytes
- Avast Antivirus
- AVG Antivirus
- Bitdefender
- Bitdefender Security Center
- Bitdefender Wallet
- Kaspersky
- ESET Security
- Norton Security
- McAfee
- KeePass
- KeePassXC
- Bitwarden
- 1Password
- NordVPN
- OpenVPN GUI
- WireGuard

#### Gaming & Launchers

- Steam
- Epic Games Launcher
- Battle.net
- GOG Galaxy
- Riot Client
- EA App
- Ubisoft Connect
- Rockstar Games Launcher
- Xbox App
- Steam Deck Remote Play
- SteamVR
- Steam (Big Picture)
- Steam (webhelper)
- Epic Games (launcher alt)
- Epic Games (web)
- GOG Galaxy (helper)
- Blitz (game client)

#### Communication

- Discord
- Skype
- Skype for Business
- Telegram Desktop
- WhatsApp Desktop
- Microsoft Teams
- Slack
- Signal Desktop
- Element
- Thunderbird
- Mail (Windows Mail)
- Zoom
- Zoom Rooms

#### Tools & Utilities

- Visual Studio Code
- Notepad++
- Notepad (classic)
- 7-Zip
- WinRAR
- VLC Media Player
- VLC (alternate)
- VLC (portable)
- Spotify
- Spotify (WindowsApps)
- OBS Studio
- OBS Studio (alternate)
- OBS Studio (portable)
- OBS Studio (Streamlabs)
- GitHub Desktop
- Docker Desktop
- FileZilla
- WinSCP
- ShareX
- ShareX (alternate)
- Paint.NET
- GIMP
- GIMP (alternate)
- Blender
- Blender (CLI)
- Adobe Acrobat Reader
- Adobe Creative Cloud
- Adobe Photoshop
- Adobe Illustrator
- Lightroom
- HandBrake
- Plex Media Player
- iTunes
- Audacity
- foobar2000
- Rufus
- Balena Etcher
- CCleaner
- Snagit
- Greenshot
- Process Monitor
- Autoruns
- Sysinternals - Process Explorer
- Everything Search
- OneDrive
- Dropbox
- Google Drive (for desktop)

#### Development & IDEs

- Visual Studio 2022
- PyCharm
- IntelliJ IDEA
- WebStorm
- Android Studio
- Sublime Text
- Atom
- Brackets
- Postman
- Insomnia
- Node.js
- Python
- Anaconda Navigator
- Git (bash)
- Git GUI
- Sourcetree
- Docker Desktop

#### Databases & Admin

- SQL Server Management Studio
- Microsoft SQL Server Configuration Manager
- DBeaver
- HeidiSQL
- Tableau

#### Virtualization & DevOps

- VMware Workstation
- VirtualBox
- Vagrant
- Terraform
- Chocolatey
- Scoop

#### Multimedia & Creativity

- Figma
- SketchUp
- Premiere/Creative apps (via Creative Cloud)
- Plex Media Player

#### Windows & System (built-ins)

- Calculator
- Windows Terminal
- Task Manager
- Registry Editor
- Device Manager (via mmc)
- Disk Management (via mmc)
- Services (via mmc)
- Event Viewer (via mmc)
- Control Panel
- Sound (classic) — control.exe mmsys.cpl
- Sound (settings) — ms-settings:sound
- Network Connections — control.exe ncpa.cpl
- Programs & Features — control.exe appwiz.cpl
- SQL Server Configuration Manager (msc)

#### Misc / Other

- PuTTY
- MobaXterm
- Cmder
- ConEmu
- PuTTY
- RStudio
- MATLAB
- SAS
- qBittorrent
- BitTorrent
- Postman
- Plex Media Player
- RAID tools (if present)
