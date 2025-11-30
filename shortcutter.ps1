# MyShortcutHub.ps1
# GUI app: buttons grouped by category; click to create desktop shortcut.
# Only shows apps that seem to be installed.

param(
    [switch]$NoGui,
    [switch]$PreviewAll
)

Add-Type -AssemblyName PresentationFramework

# --- XAML for the window (grouped by category) ---
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Shortcutter"
    Height="820" Width="1060" MinHeight="620" MinWidth="800"
    ResizeMode="CanResize"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header: logo + app name + short description -->
        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Simple vector logo built with shapes so no external image is required -->
            <Border Width="64" Height="64" CornerRadius="10" Background="#FF2D89EF" VerticalAlignment="Top">
                <Grid>
                    <TextBlock Text="S" Foreground="White" FontWeight="Bold" FontSize="36" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Grid>
            </Border>

            <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
                <TextBlock Text="Shortcutter" FontSize="22" FontWeight="Bold" />
                <TextBlock Text="Click a button to create a desktop shortcut. Only installed apps are shown." FontSize="13" Foreground="#444444" TextWrapping="Wrap"/>
            </StackPanel>
        </Grid>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel Orientation="Vertical" Margin="0,0,0,10">

                <GroupBox Header="Browsers" Margin="0,0,0,10">
                    <WrapPanel x:Name="BrowsersPanel" Margin="5"/>
                </GroupBox>

                <GroupBox Header="Security &amp; Antivirus" Margin="0,0,0,10">
                    <WrapPanel x:Name="SecurityPanel" Margin="5"/>
                </GroupBox>

                <GroupBox Header="Gaming &amp; Launchers" Margin="0,0,0,10">
                    <WrapPanel x:Name="GamingPanel" Margin="5"/>
                </GroupBox>
                <GroupBox Header="Windows &amp; System" Margin="0,0,0,10">
                    <WrapPanel x:Name="WindowsPanel" Margin="5"/>
                </GroupBox>

                <GroupBox Header="Communication" Margin="0,0,0,10">
                    <WrapPanel x:Name="CommPanel" Margin="5"/>
                </GroupBox>

                <GroupBox Header="Tools &amp; Utilities" Margin="0,0,0,10">
                    <WrapPanel x:Name="ToolsPanel" Margin="5"/>
                </GroupBox>

            </StackPanel>
        </ScrollViewer>

        <StackPanel Grid.Row="2"
                    Orientation="Horizontal"
                    HorizontalAlignment="Stretch"
                    Margin="0,10,0,0">
            <TextBlock x:Name="StatusText"
                       Text=""
                       VerticalAlignment="Center"
                       Margin="0,0,10,0"
                       TextWrapping="Wrap"
                       Width="680"/>
            <Button x:Name="CloseButton"
                    Content="Close"
                    Width="100"
                    HorizontalAlignment="Right"/>
        </StackPanel>
    </Grid>
</Window>
"@

# --- Parse XAML to live window ---
$window = [Windows.Markup.XamlReader]::Parse($xaml)

# --- Find controls ---
$BrowsersPanel = $window.FindName("BrowsersPanel")
$SecurityPanel = $window.FindName("SecurityPanel")
$GamingPanel   = $window.FindName("GamingPanel")
$CommPanel     = $window.FindName("CommPanel")
$ToolsPanel    = $window.FindName("ToolsPanel")
$WindowsPanel  = $window.FindName("WindowsPanel")
$StatusText    = $window.FindName("StatusText")
$CloseButton   = $window.FindName("CloseButton")

# --- Desktop path & shortcut creator ---
$DesktopPath = [Environment]::GetFolderPath("Desktop")

function New-DesktopShortcut {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Target,
        [string]$Arguments = "",
        [string]$IconPath  = $null,
        [switch]$WhatIf
    )
    # Expand any environment variables
    $expandedTarget = [Environment]::ExpandEnvironmentVariables($Target)
    $expandedIcon   = if ($IconPath) { [Environment]::ExpandEnvironmentVariables($IconPath) } else { $null }

    # Decide how to launch the target:
    # - URI schemes (like ms-settings:) -> use explorer.exe <uri>
    # - Existing executable/msc/cpl paths -> use directly
    # - control.exe with .cpl arguments -> use control.exe from System32
    $targetPathToUse = $null
    $argsToUse = $Arguments

    # Special-case: if the provided Target is (or points to) control.exe, keep it as control.exe
    # and use the CPL argument directly. This ensures shortcuts like "control mmsys.cpl" are created
    # (no explorer or cmd wrappers that flash a terminal).
    if ($expandedTarget -and ($expandedTarget -match 'control\.exe$' -or $expandedTarget -match '\\control\.exe$' )) {
        $possibleControl = $expandedTarget
        if (-not (Test-Path $possibleControl)) {
            $possibleControl = [Environment]::ExpandEnvironmentVariables('%windir%\System32\control.exe')
        }
        if (Test-Path $possibleControl) {
            $targetPathToUse = (Resolve-Path $possibleControl).Path
            # Keep Arguments as-is (expected to be a .cpl filename like mmsys.cpl)
            $argsToUse = $Arguments
        }
    }

    # Detect URI schemes (ms-settings:, shell:, etc.) but avoid matching Windows drive-letter paths like C:\...
    if ($expandedTarget -and ($expandedTarget -match '^[a-zA-Z][a-zA-Z0-9+\-.]*:' ) -and -not ($expandedTarget -match '^[A-Za-z]:\\')) {
        # URI detected
        # Use explorer.exe <uri> to open settings URIs without launching a terminal window
        $targetPathToUse = [Environment]::ExpandEnvironmentVariables('%windir%\explorer.exe')
        $argsToUse = $expandedTarget
    } elseif (Test-Path $expandedTarget) {
        $targetPathToUse = (Resolve-Path $expandedTarget).Path
    } else {
        # If user provided a Target without a path but with a CPL/MSC argument, prefer control.exe or mmc
        if ($Arguments -and $Arguments -match '\.(cpl|msc)$') {
            $possibleControl = [Environment]::ExpandEnvironmentVariables('%windir%\System32\control.exe')
            if (Test-Path $possibleControl) {
                $targetPathToUse = $possibleControl
            }
        }

        # Last resort: if the expandedTarget looks like an executable name (endswith .exe) but isn't found,
        # try resolving in System32
        if (-not $targetPathToUse -and $expandedTarget -match '\.exe$') {
            $sys32 = [Environment]::ExpandEnvironmentVariables('%windir%\System32')
            $tryPath = Join-Path $sys32 (Split-Path $expandedTarget -Leaf)
            if (Test-Path $tryPath) { $targetPathToUse = $tryPath }
        }

        if (-not $targetPathToUse) {
            $StatusText.Text = "Target not found for $Name. ($Target)"
            return
        }
    }

    # If the arguments refer to a Control Panel .cpl, point the shortcut at control.exe with the CPL as the argument
    if ($argsToUse -and $argsToUse -match '\.cpl(\s|$)') {
        $controlExe = [Environment]::ExpandEnvironmentVariables('%windir%\System32\control.exe')
        if (Test-Path $controlExe) {
            # Keep only the cpl filename (in case it was a full path or had extra whitespace)
            $cplName = $argsToUse -replace '.*?(\S+\.cpl).*','$1'
            $targetPathToUse = $controlExe
            $argsToUse = $cplName

            # If the CPL file exists in System32, use it as the icon
            $cplPath = [Environment]::ExpandEnvironmentVariables("%windir%\System32\$cplName")
            if (Test-Path $cplPath) { $expandedIcon = $cplPath }
        }
    }

    # If the target is an MMC snap-in (.msc), prefer launching via mmc.exe for consistency
    if ($targetPathToUse -and ($targetPathToUse -match '\.msc$')) {
        $mmcExe = [Environment]::ExpandEnvironmentVariables('%windir%\System32\mmc.exe')
        if (Test-Path $mmcExe) {
            $mscFull = $targetPathToUse
            $targetPathToUse = $mmcExe
            $argsToUse = $mscFull
            # prefer the .msc as icon if available
            if (-not $expandedIcon -and (Test-Path $mscFull)) { $expandedIcon = $mscFull }
        }
    }

    # Compute the planned final icon and working directory without yet writing a file.
    $finalIcon = $null
    $finalWD = $null

    # Prefer explicit expandedIcon if provided and exists
    if ($expandedIcon -and (Test-Path $expandedIcon)) {
        $finalIcon = $expandedIcon
    }

    # If target is an exe, prefer its folder as WorkingDirectory and its file as icon (if no explicit icon)
    if ($targetPathToUse -and ($targetPathToUse -match '\.exe$')) {
        try {
            $exePath = $targetPathToUse
            if (Test-Path $exePath) {
                if (-not $finalIcon) { $finalIcon = $exePath }
                $wdCandidate = Split-Path $exePath -Parent
                if ($wdCandidate) { $finalWD = $wdCandidate }
            }
        } catch { }
    }

    # If target is control.exe and argument is a cpl, set WorkingDirectory to System32 and icon to the CPL if available
    if ($targetPathToUse -and ($targetPathToUse -match 'control\.exe$') -and ($argsToUse -and $argsToUse -match '\.cpl')) {
        $sys32 = [Environment]::ExpandEnvironmentVariables('%windir%\System32')
        if (Test-Path $sys32) { $finalWD = $sys32 }
        $cplName = $argsToUse -replace '.*?(\S+\.cpl).*', '$1'
        $cplPath = Join-Path $sys32 $cplName
        if (Test-Path $cplPath) { $finalIcon = $cplPath }
    }

    # If target is explorer.exe launching a URI, set WorkingDirectory to %windir%
    if ($targetPathToUse -and ($targetPathToUse -match 'explorer\.exe$') -and ($argsToUse -and $argsToUse -match '^[a-zA-Z][a-zA-Z0-9+\-.]*:')) {
        $finalWD = [Environment]::ExpandEnvironmentVariables('%windir%')
    }

    # Special-case Discord Update.exe with --processStart: ensure WorkingDirectory is Update.exe folder and prefer Discord.exe icon
    if ($targetPathToUse -and ($targetPathToUse -match 'Update\.exe$') -and ($argsToUse -and $argsToUse -match '--processStart')) {
        try {
            $wd = Split-Path $targetPathToUse -Parent
            if ($wd) { $finalWD = $wd }
            $localDiscord = [Environment]::ExpandEnvironmentVariables('%LocalAppData%\Discord')
            if (Test-Path $localDiscord) {
                $exe = Get-ChildItem -Path (Join-Path $localDiscord 'app-*') -Filter 'Discord.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($exe) { $finalIcon = $exe.FullName }
            }
        } catch { }
    }

    # Build a preview object describing what would be created
    $preview = [PSCustomObject]@{
        Name = $Name
        LnkPath = Join-Path $DesktopPath "$Name.lnk"
        Target = $targetPathToUse
        Arguments = $argsToUse
        WorkingDirectory = $finalWD
        Icon = $finalIcon
    }

    # Write debug info to Desktop log (safe for WhatIf because we don't write .lnk)
    try {
        $debugLog = Join-Path $DesktopPath 'shortcutter-debug.log'
        $debugLine = "Creating shortcut: `n  Path: $($preview.LnkPath)`n  Target: $($preview.Target)`n  Arguments: $($preview.Arguments)`n  Icon: $($preview.Icon)`n  WorkingDir: $($preview.WorkingDirectory)"
        Add-Content -Path $debugLog -Value ("$(Get-Date -Format o) - $debugLine") -ErrorAction SilentlyContinue
    } catch { }

    if ($WhatIf) {
        # Return the preview object for inspection without creating the .lnk
        return $preview
    }

    # Otherwise create the actual shortcut
    $shell = New-Object -ComObject WScript.Shell
    $lnkPath = $preview.LnkPath
    $StatusText.Text = "Creating: $Name..."

    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath = $preview.Target
    if ($preview.Arguments) { $shortcut.Arguments = $preview.Arguments }
    if ($preview.Icon -and (Test-Path $preview.Icon)) { $shortcut.IconLocation = $preview.Icon }
    if ($preview.WorkingDirectory) { $shortcut.WorkingDirectory = $preview.WorkingDirectory }
    $shortcut.Save()

    $StatusText.Text = "Created shortcut: $lnkPath"
    try { Add-Content -Path $debugLog -Value ("$(Get-Date -Format o) - Created shortcut: $lnkPath (Target=$($preview.Target) Args=$($preview.Arguments) WorkingDir=$($shortcut.WorkingDirectory) Icon=$($shortcut.IconLocation))") -ErrorAction SilentlyContinue } catch { }
    return $preview
}

# --- Helper: find executable from list of candidate paths ---
function Find-Executable {
    param([string[]]$CandidatePaths)

    foreach ($raw in $CandidatePaths) {
        if (-not $raw) { continue }
        $expanded = [Environment]::ExpandEnvironmentVariables($raw)

        # Supports wildcards
        if (Test-Path $expanded) {
            try {
                $resolved = Resolve-Path $expanded -ErrorAction Stop | Select-Object -First 1
                if ($resolved -and (Test-Path $resolved.Path)) {
                    return $resolved.Path
                }
            } catch {
                # ignore and continue
            }
        }
    }
    return $null
}


# --- App catalog (many apps), grouped logically ---
$apps = @(
    # --- Browsers ---
    @{
        Name = "Google Chrome"
        Category = "Browsers"
        Paths = @(
            "%ProgramFiles%\Google\Chrome\Application\chrome.exe",
            "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
        )
    }
    @{
        Name = "Mozilla Firefox"
        Category = "Browsers"
        Paths = @(
            "%ProgramFiles%\Mozilla Firefox\firefox.exe",
            "%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe"
        )
    }
    @{
        Name = "Microsoft Edge"
        Category = "Browsers"
        Paths = @(
            "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
            "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
        )
    }
    @{
        Name = "Brave Browser"
        Category = "Browsers"
        Paths = @(
            "%ProgramFiles%\BraveSoftware\Brave-Browser\Application\brave.exe",
            "%ProgramFiles(x86)%\BraveSoftware\Brave-Browser\Application\brave.exe"
        )
    }
    @{
        Name = "Opera GX"
        Category = "Browsers"
        Paths = @(
            "%LocalAppData%\Programs\Opera GX\launcher.exe",
            "%ProgramFiles%\Opera GX\launcher.exe"
        )
    }
    @{
        Name = "Vivaldi"
        Category = "Browsers"
        Paths = @(
            "%LocalAppData%\Vivaldi\Application\vivaldi.exe",
            "%ProgramFiles%\Vivaldi\Application\vivaldi.exe"
        )
    }
    @{
        Name = "Tor Browser"
        Category = "Browsers"
        Paths = @(
            "%ProgramFiles%\Tor Browser\Browser\firefox.exe",
            "%LocalAppData%\Tor Browser\Browser\firefox.exe"
        )
    }

    # --- Security & Antivirus ---
    @{
        Name = "Malwarebytes"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Malwarebytes\Anti-Malware\mbam.exe",
            "%ProgramFiles%\Malwarebytes\Anti-Malware\mbamtray.exe"
        )
    }
    @{
        Name = "Avast Antivirus"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Avast Software\Avast\AvastUI.exe",
            "%ProgramFiles(x86)%\Avast Software\Avast\AvastUI.exe"
        )
    }
    @{
        Name = "AVG Antivirus"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\AVG\Antivirus\AVGUI.exe",
            "%ProgramFiles(x86)%\AVG\Antivirus\AVGUI.exe"
        )
    }
    @{
        Name = "Bitdefender"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Bitdefender\Bitdefender Security\bdsecurity.exe",
            "%ProgramFiles%\Bitdefender\Bitdefender Antivirus Free\bdav.exe"
        )
    }
    @{
        Name = "Kaspersky"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Kaspersky Lab\Kaspersky*\avpui.exe",
            "%ProgramFiles(x86)%\Kaspersky Lab\Kaspersky*\avpui.exe"
        )
    }
    @{
        Name = "ESET Security"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\ESET\ESET Security\egui.exe",
            "%ProgramFiles%\ESET\ESET NOD32 Antivirus\egui.exe"
        )
    }
    @{
        Name = "Norton Security"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Norton Security\Engine\*\uiStub.exe",
            "%ProgramFiles(x86)%\Norton Security\Engine\*\uiStub.exe"
        )
    }
    @{
        Name = "McAfee"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\McAfee\Agent\mcagent.exe",
            "%ProgramFiles(x86)%\McAfee\Agent\mcagent.exe"
        )
    }

    # --- Gaming & Launchers ---
    @{
        Name = "Steam"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\Steam\steam.exe",
            "%ProgramFiles%\Steam\steam.exe"
        )
    }
    @{
        Name = "Epic Games Launcher"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\Epic Games\Launcher\Portal\Binaries\Win64\EpicGamesLauncher.exe"
        )
    }
    @{
        Name = "Battle.net"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\Battle.net\Battle.net.exe"
        )
    }
    @{
        Name = "GOG Galaxy"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\GOG Galaxy\GalaxyClient.exe"
        )
    }
    @{
        Name = "Riot Client"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles%\Riot Games\Riot Client\RiotClientServices.exe"
        )
    }
    @{
        Name = "EA App"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles%\EA\EA Desktop\EA Desktop\EA Desktop.exe",
            "%ProgramFiles(x86)%\EA Games\EA Desktop\EA Desktop.exe"
        )
    }
    @{
        Name = "Ubisoft Connect"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\Ubisoft\Ubisoft Game Launcher\upc.exe"
        )
    }
    @{
        Name = "Rockstar Games Launcher"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles%\Rockstar Games\Launcher\Launcher.exe"
        )
    }
    @{
        Name = "Xbox App"
        Category = "Gaming"
        Paths = @(
            "%LocalAppData%\Microsoft\WindowsApps\XboxApp.exe"
        )
    }

    # --- Communication ---
    @{
        Name = "Discord"
        Category = "Comm"
        # Use Update.exe as the launcher but include the --processStart argument so the shortcut starts Discord.exe
        Target = "%LocalAppData%\Discord\Update.exe"
        Arguments = "--processStart Discord.exe"
        # Also include an executable path for detection fallback
        Paths = @(
            "%LocalAppData%\Discord\app-*\Discord.exe"
        )
    }
    @{
        Name = "Skype"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles(x86)%\Microsoft\Skype for Desktop\Skype.exe"
        )
    }
    @{
        Name = "Telegram Desktop"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles%\Telegram Desktop\Telegram.exe",
            "%LocalAppData%\Telegram Desktop\Telegram.exe"
        )
    }
    @{
        Name = "WhatsApp Desktop"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\WhatsApp\WhatsApp.exe"
        )
    }
    @{
        Name = "Microsoft Teams"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\Microsoft\Teams\current\Teams.exe",
            "%ProgramFiles%\Microsoft\Teams\current\Teams.exe"
        )
    }
    @{
        Name = "Slack"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\slack\slack.exe",
            "%ProgramFiles%\Slack\slack.exe"
        )
    }
    @{
        Name = "Zoom"
        Category = "Comm"
        Paths = @(
            "%AppData%\Zoom\bin\Zoom.exe"
        )
    }
    @{
        Name = "Signal Desktop"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\Programs\signal-desktop\Signal.exe"
        )
    }
    @{
        Name = "Element"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\Programs\Element\element-desktop.exe"
        )
    }

    # --- Tools & Utilities ---
    @{
        Name = "Visual Studio Code"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\Microsoft VS Code\Code.exe",
            "%ProgramFiles%\Microsoft VS Code\Code.exe"
        )
    }
    @{
        Name = "Notepad++"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Notepad++\notepad++.exe",
            "%ProgramFiles(x86)%\Notepad++\notepad++.exe"
        )
    }
    @{
        Name = "7-Zip"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\7-Zip\7zFM.exe",
            "%ProgramFiles(x86)%\7-Zip\7zFM.exe"
        )
    }
    @{
        Name = "VLC Media Player"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\VideoLAN\VLC\vlc.exe",
            "%ProgramFiles(x86)%\VideoLAN\VLC\vlc.exe"
        )
    }
    @{
        Name = "Spotify"
        Category = "Tools"
        Paths = @(
            "%AppData%\Spotify\Spotify.exe",
            "%LocalAppData%\Microsoft\WindowsApps\Spotify.exe"
        )
    }
    @{
        Name = "OBS Studio"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\obs-studio\bin\64bit\obs64.exe",
            "%ProgramFiles(x86)%\obs-studio\bin\32bit\obs32.exe"
        )
    }
    @{
        Name = "GitHub Desktop"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\GitHubDesktop\GitHubDesktop.exe",
            "%ProgramFiles%\GitHub Desktop\GitHubDesktop.exe"
        )
    }
    @{
        Name = "Docker Desktop"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Docker\Docker\Docker Desktop.exe"
        )
    }
    @{
        Name = "FileZilla"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\FileZilla FTP Client\filezilla.exe",
            "%ProgramFiles(x86)%\FileZilla FTP Client\filezilla.exe"
        )
    }
    @{
        Name = "WinSCP"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\WinSCP\WinSCP.exe",
            "%ProgramFiles(x86)%\WinSCP\WinSCP.exe"
        )
    }
    @{
        Name = "ShareX"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\ShareX\ShareX.exe",
            "%ProgramFiles(x86)%\ShareX\ShareX.exe"
        )
    }
    @{
        Name = "Paint.NET"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\paint.net\PaintDotNet.exe"
        )
    }
    @{
        Name = "GIMP"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\GIMP*\bin\gimp-*.exe"
        )
    }
    @{
        Name = "Blender"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Blender Foundation\Blender\blender.exe"
        )
    }
    @{
        Name = "Adobe Acrobat Reader"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            "%ProgramFiles(x86)%\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
        )
    }
    @{
        Name = "Notion"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\Notion\Notion.exe"
        )
    }
    @{
        Name = "Everything Search"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Everything\Everything.exe",
            "%ProgramFiles(x86)%\Everything\Everything.exe"
        )
    }
    
    # --- Windows & System built-ins ---
    @{
        Name = "Calculator"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\calc.exe"
        )
    }
    @{
        Name = "Windows Terminal"
        Category = "Windows"
        Paths = @(
            "%LocalAppData%\Microsoft\WindowsApps\wt.exe",
            "%ProgramFiles%\WindowsApps\wt.exe"
        )
    }
    @{
        Name = "Task Manager"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\Taskmgr.exe"
        )
    }
    @{
        Name = "Registry Editor"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\regedit.exe"
        )
    }
    @{
        Name = "Device Manager"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\devmgmt.msc"
        )
    }
    @{
        Name = "Disk Management"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\diskmgmt.msc"
        )
    }
    @{
        Name = "Services"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\services.msc"
        )
    }
    @{
        Name = "Event Viewer"
        Category = "Windows"
        Paths = @(
            "%windir%\System32\eventvwr.msc"
        )
    }
    @{
        Name = "Control Panel"
        Category = "Windows"
        Target = "%windir%\System32\control.exe"
        Arguments = ""
    }
    @{
        Name = "Sound (classic)"
        Category = "Windows"
        Target = "%windir%\System32\control.exe"
        Arguments = "mmsys.cpl"
    }
    @{
        Name = "Sound (settings)"
        Category = "Windows"
        Target = "ms-settings:sound"
    }
    @{
        Name = "Network Connections"
        Category = "Windows"
        Target = "%windir%\System32\control.exe"
        Arguments = "ncpa.cpl"
    }
    @{
        Name = "Programs & Features"
        Category = "Windows"
        Target = "%windir%\System32\control.exe"
        Arguments = "appwiz.cpl"
    }

    # --- Extra catalog entries (100+) ---
    @{
        Name = "Microsoft Word"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE",
            "%ProgramFiles%\Microsoft Office\Office16\WINWORD.EXE"
        )
    }
    @{
        Name = "Microsoft Excel"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE",
            "%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE"
        )
    }
    @{
        Name = "Microsoft PowerPoint"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE",
            "%ProgramFiles%\Microsoft Office\Office16\POWERPNT.EXE"
        )
    }
    @{
        Name = "Microsoft Outlook"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "%ProgramFiles%\Microsoft Office\Office16\OUTLOOK.EXE"
        )
    }
    @{
        Name = "LibreOffice"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\LibreOffice\program\swriter.exe",
            "%ProgramFiles(x86)%\LibreOffice\program\swriter.exe"
        )
    }
    @{
        Name = "OneDrive"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Microsoft\OneDrive\OneDrive.exe",
            "%ProgramFiles%\Microsoft OneDrive\OneDrive.exe"
        )
    }
    @{
        Name = "Dropbox"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Dropbox\Client\Dropbox.exe",
            "%LocalAppData%\Dropbox\Client\Dropbox.exe"
        )
    }
    @{
        Name = "Google Drive (for desktop)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Google\Drive File Stream\43.0.0.0\GoogleDriveFS.exe",
            "%ProgramFiles%\Google\Drive\googledrivesync.exe",
            "%LocalAppData%\Google\Drive\googledrivesync.exe"
        )
    }
    @{
        Name = "Adobe Creative Cloud"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Adobe\Adobe Creative Cloud\ACC\Creative Cloud.exe",
            "%LocalAppData%\Adobe\Creative Cloud\ACC\Creative Cloud.exe"
        )
    }
    @{
        Name = "Adobe Photoshop"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Adobe\Adobe Photoshop 2023\Photoshop.exe",
            "%ProgramFiles%\Adobe\Adobe Photoshop 2022\Photoshop.exe"
        )
    }
    @{
        Name = "Adobe Illustrator"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Adobe\Adobe Illustrator 2023\Support Files\Contents\Windows\Illustrator.exe"
        )
    }
    @{
        Name = "Lightroom"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Adobe\Adobe Lightroom Classic\lightroom.exe"
        )
    }
    @{
        Name = "HandBrake"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\HandBrake\HandBrake.exe",
            "%ProgramFiles(x86)%\HandBrake\HandBrake.exe"
        )
    }
    @{
        Name = "Plex Media Player"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Plex\Plex Media Player\Plex.exe",
            "%ProgramFiles%\Plex\Plex Media Server\Plex Media Server.exe"
        )
    }
    @{
        Name = "iTunes"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\iTunes\iTunes.exe",
            "%ProgramFiles(x86)%\iTunes\iTunes.exe"
        )
    }
    @{
        Name = "Audacity"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Audacity\audacity.exe",
            "%ProgramFiles(x86)%\Audacity\audacity.exe"
        )
    }
    @{
        Name = "foobar2000"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\foobar2000\foobar2000.exe",
            "%ProgramFiles(x86)%\foobar2000\foobar2000.exe"
        )
    }
    @{
        Name = "WinRAR"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\WinRAR\WinRAR.exe",
            "%ProgramFiles(x86)%\WinRAR\WinRAR.exe"
        )
    }
    @{
        Name = "Rufus"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Rufus\rufus.exe",
            "%LocalAppData%\Programs\Rufus\rufus.exe"
        )
    }
    @{
        Name = "Balena Etcher"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\balenaEtcher\Etcher.exe",
            "%LocalAppData%\Programs\balenaEtcher\Etcher.exe"
        )
    }
    @{
        Name = "CCleaner"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\CCleaner\CCleaner.exe",
            "%ProgramFiles(x86)%\CCleaner\CCleaner.exe"
        )
    }
    @{
        Name = "Sysinternals - Process Explorer"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Sysinternals\procexp.exe",
            "%ProgramFiles(x86)%\Sysinternals\procexp.exe",
            "%ProgramFiles%\Microsoft Sysinternals\procexp.exe"
        )
    }
    @{
        Name = "Process Monitor"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Sysinternals\procmon.exe",
            "%ProgramFiles(x86)%\Sysinternals\procmon.exe"
        )
    }
    @{
        Name = "Autoruns"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Sysinternals\autoruns.exe",
            "%ProgramFiles(x86)%\Sysinternals\autoruns.exe"
        )
    }
    @{
        Name = "PuTTY"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\PuTTY\putty.exe",
            "%ProgramFiles(x86)%\PuTTY\putty.exe"
        )
    }
    @{
        Name = "MobaXterm"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Mobatek\MobaXterm\MobaXterm.exe",
            "%LocalAppData%\Programs\MobaXterm\MobaXterm.exe"
        )
    }
    @{
        Name = "ConEmu"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\ConEmu\ConEmu64.exe",
            "%ProgramFiles(x86)%\ConEmu\ConEmu.exe"
        )
    }
    @{
        Name = "Cmder"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\cmder\Cmder.exe",
            "%LocalAppData%\Programs\cmder\Cmder.exe"
        )
    }
    @{
        Name = "Git (bash)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Git\git-bash.exe",
            "%ProgramFiles%\Git\bin\bash.exe",
            "%ProgramFiles%\Git\cmd\git.exe"
        )
    }
    @{
        Name = "Git GUI"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Git\cmd\git-gui.exe",
            "%ProgramFiles%\Git\git-gui.exe"
        )
    }
    @{
        Name = "Sourcetree"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Atlassian\SourceTree\SourceTree.exe",
            "%ProgramFiles%\Atlassian\SourceTree\SourceTree.exe"
        )
    }
    @{
        Name = "Node.js"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\nodejs\node.exe",
            "%ProgramFiles(x86)%\nodejs\node.exe"
        )
    }
    @{
        Name = "Python"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\Python\Python39\python.exe",
            "%ProgramFiles%\Python39\python.exe",
            "%ProgramFiles%\Python\python.exe"
        )
    }
    @{
        Name = "PyCharm"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\JetBrains\PyCharm Community Edition 2023.2\bin\pycharm64.exe",
            "%ProgramFiles%\JetBrains\PyCharm\bin\pycharm64.exe"
        )
    }
    @{
        Name = "IntelliJ IDEA"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\JetBrains\IntelliJ IDEA Community Edition\bin\idea64.exe",
            "%ProgramFiles%\JetBrains\IntelliJ IDEA\bin\idea64.exe"
        )
    }
    @{
        Name = "WebStorm"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\JetBrains\WebStorm\bin\webstorm64.exe",
            "%ProgramFiles%\JetBrains\WebStorm\bin\webstorm64.exe"
        )
    }
    @{
        Name = "Android Studio"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Android\Android Studio\bin\studio64.exe",
            "%ProgramFiles%\Android\Android Studio\bin\studio.exe"
        )
    }
    @{
        Name = "Visual Studio 2022"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe",
            "%ProgramFiles%\Microsoft Visual Studio\2022\Professional\Common7\IDE\devenv.exe"
        )
    }
    @{
        Name = "SQL Server Management Studio"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Microsoft SQL Server\150\Tools\Binn\ManagementStudio\Ssms.exe",
            "%ProgramFiles(x86)%\Microsoft SQL Server\120\Tools\Binn\ManagementStudio\Ssms.exe"
        )
    }
    @{
        Name = "DBeaver"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\DBeaver\dbeaver.exe",
            "%LocalAppData%\Programs\DBeaver\dbeaver.exe"
        )
    }
    @{
        Name = "HeidiSQL"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\HeidiSQL\heidisql.exe",
            "%ProgramFiles(x86)%\HeidiSQL\heidisql.exe"
        )
    }
    @{
        Name = "VMware Workstation"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\VMware\VMware Workstation\vmware.exe",
            "%ProgramFiles(x86)%\VMware\VMware Workstation\vmware.exe"
        )
    }
    @{
        Name = "VirtualBox"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Oracle\VirtualBox\VirtualBox.exe",
            "%ProgramFiles(x86)%\Oracle\VirtualBox\VirtualBox.exe"
        )
    }
    @{
        Name = "KeePass"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\KeePass Password Safe 2\KeePass.exe",
            "%ProgramFiles(x86)%\KeePass Password Safe 2\KeePass.exe"
        )
    }
    @{
        Name = "KeePassXC"
        Category = "Security"
        Paths = @(
            "%LocalAppData%\Programs\KeePassXC\KeePassXC.exe",
            "%ProgramFiles%\KeePassXC\KeePassXC.exe"
        )
    }
    @{
        Name = "Bitwarden"
        Category = "Security"
        Paths = @(
            "%LocalAppData%\Programs\Bitwarden\Bitwarden.exe",
            "%ProgramFiles%\Bitwarden\Bitwarden.exe"
        )
    }
    @{
        Name = "1Password"
        Category = "Security"
        Paths = @(
            "%LocalAppData%\Programs\1Password\1Password.exe",
            "%ProgramFiles%\1Password\1Password.exe"
        )
    }
    @{
        Name = "NordVPN"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Nord VPN\NordVPN.exe",
            "%ProgramFiles%\NordVPN\NordVPN.exe"
        )
    }
    @{
        Name = "OpenVPN GUI"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\OpenVPN\bin\openvpn-gui.exe",
            "%ProgramFiles(x86)%\OpenVPN\bin\openvpn-gui.exe"
        )
    }
    @{
        Name = "WireGuard"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\WireGuard\wireguard.exe",
            "%ProgramFiles(x86)%\WireGuard\wireguard.exe"
        )
    }
    @{
        Name = "Bitdefender Wallet"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\Bitdefender\Bitdefender Wallet\bdwallet.exe"
        )
    }
    @{
        Name = "NVIDIA GeForce Experience"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\NVIDIA Corporation\NVIDIA GeForce Experience\NVIDIA GeForce Experience.exe"
        )
    }
    @{
        Name = "Intel Graphics Command Center"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Intel\GraphicsCommandCenter\GfxUI.exe",
            "%ProgramFiles%\Intel\Intel Graphics Command Center\GfxUI.exe"
        )
    }
    @{
        Name = "Display Driver Uninstaller"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Display Driver Uninstaller\DDU.exe"
        )
    }
    @{
        Name = "Steam Deck Remote Play"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\Steam\steam.exe"
        )
    }
    @{
        Name = "OBS Studio (Streamlabs)"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\streamlabs\Streamlabs OBS\Streamlabs OBS.exe",
            "%ProgramFiles%\Streamlabs\Streamlabs OBS\Streamlabs OBS.exe"
        )
    }
    @{
        Name = "Spotify (WindowsApps)"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Microsoft\WindowsApps\Spotify.exe"
        )
    }
    @{
        Name = "Notepad (classic)"
        Category = "Tools"
        Paths = @(
            "%windir%\system32\notepad.exe"
        )
    }
    @{
        Name = "Sublime Text"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Sublime Text 3\sublime_text.exe",
            "%ProgramFiles%\Sublime Text\sublime_text.exe"
        )
    }
    @{
        Name = "Atom"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\Atom\atom.exe",
            "%ProgramFiles%\Atom\atom.exe"
        )
    }
    @{
        Name = "Postman"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Postman\Postman.exe",
            "%ProgramFiles%\Postman\Postman.exe"
        )
    }
    @{
        Name = "Insomnia"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\Insomnia\insomnia.exe",
            "%ProgramFiles%\Insomnia\insomnia.exe"
        )
    }
    @{
        Name = "Figma"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\Programs\Figma\Figma.exe"
        )
    }
    @{
        Name = "SketchUp"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\SketchUp\SketchUp 2021\SketchUp.exe"
        )
    }
    @{
        Name = "Blender (CLI)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Blender Foundation\Blender\blender.exe"
        )
    }
    @{
        Name = "GIMP (alternate)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\GIMP 2\bin\gimp-2.10.exe"
        )
    }
    @{
        Name = "Everything (Admin)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Everything\Everything.exe"
        )
    }
    @{
        Name = "Vagrant"
        Category = "Tools"
        Paths = @(
            "%Hash%\Vagrant\bin\vagrant.exe",
            "%ProgramFiles%\HashiCorp\Vagrant\bin\vagrant.exe"
        )
    }
    @{
        Name = "Terraform"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\Terraform\terraform.exe",
            "%ProgramFiles%\HashiCorp\terraform.exe"
        )
    }
    @{
        Name = "Chocolatey"
        Category = "Tools"
        Paths = @(
            "C:\\ProgramData\\chocolatey\\bin\\choco.exe"
        )
    }
    @{
        Name = "Scoop"
        Category = "Tools"
        Paths = @(
            "%UserProfile%\\scoop\\shims\\scoop.ps1",
            "%UserProfile%\\scoop\\shims\\python.exe"
        )
    }
    @{
        Name = "VLC (alternate)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\VideoLAN\\VLC\\vlc.exe"
        )
    }
    @{
        Name = "Skype for Business"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles%\\Microsoft Office\\root\\Office16\\lync.exe",
            "%ProgramFiles%\\Microsoft Office\\Office16\\lync.exe"
        )
    }
    @{
        Name = "Thunderbird"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles%\\Mozilla Thunderbird\\thunderbird.exe",
            "%LocalAppData%\\Programs\\Thunderbird\\thunderbird.exe"
        )
    }
    @{
        Name = "Mail (Windows Mail)"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles%\\WindowsApps\\microsoft.windowscommunicationsapps_8wekyb3d8bbwe\\Mail.exe"
        )
    }
    @{
        Name = "Zoom (alternate)"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles%\\Zoom\\bin\\Zoom.exe",
            "%AppData%\\Zoom\\bin\\Zoom.exe"
        )
    }
    @{
        Name = "BitTorrent"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\uTorrent\\uTorrent.exe",
            "%ProgramFiles%\\BitTorrent\\BitTorrent.exe"
        )
    }
    @{
        Name = "qBittorrent"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\qBittorrent\\qbittorrent.exe",
            "%ProgramFiles(x86)%\\qBittorrent\\qbittorrent.exe"
        )
    }
    @{
        Name = "VLC (portable)"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\\Programs\\VLC\\vlc.exe"
        )
    }
    @{
        Name = "Steam (Big Picture)"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\\Steam\\steam.exe"
        )
    }
    @{
        Name = "GIMP Plugin Registry"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\GIMP 2\\bin\\gimp-2.10.exe"
        )
    }
    @{
        Name = "OBS Studio (alternate)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\obs-studio\\bin\\64bit\\obs64.exe"
        )
    }
    @{
        Name = "Microsoft SQL Server Configuration Manager"
        Category = "Windows"
        Paths = @(
            "%windir%\\System32\\SQLServerManager15.msc",
            "%windir%\\System32\\SQLServerManager14.msc"
        )
    }
    @{
        Name = "Snagit"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\TechSmith\\Snagit 2021\\SnagitEditor.exe",
            "%ProgramFiles%\\TechSmith\\Snagit 2022\\SnagitEditor.exe"
        )
    }
    @{
        Name = "Greenshot"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\Greenshot\\Greenshot.exe",
            "%ProgramFiles(x86)%\\Greenshot\\Greenshot.exe"
        )
    }
    @{
        Name = "ShareX (alternate)"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\ShareX\\ShareX.exe"
        )
    }
    @{
        Name = "Bitdefender Security Center"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\\Bitdefender\\Agent\\bdservicehost.exe"
        )
    }
    @{
        Name = "Malwarebytes Tray"
        Category = "Security"
        Paths = @(
            "%ProgramFiles%\\Malwarebytes\\Anti-Malware\\mbamtray.exe"
        )
    }
    @{
        Name = "Zoom Rooms"
        Category = "Comm"
        Paths = @(
            "%ProgramFiles%\\ZoomRooms\\zoomrooms.exe"
        )
    }
    @{
        Name = "Steam (webhelper)"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\\Steam\\bin\\cef\\cef.win64\\steamwebhelper.exe",
            "%ProgramFiles%\\Steam\\bin\\cef\\cef.win64\\steamwebhelper.exe"
        )
    }
    @{
        Name = "Epic Games (web)"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles%\\Epic Games\\Launcher\\Portal\\Binaries\\Win64\\EpicWebHelper.exe"
        )
    }
    @{
        Name = "GOG Galaxy (helper)"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles%\\GOG Galaxy\\GalaxyClient.exe"
        )
    }
    @{
        Name = "Unity Hub"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\Unity Hub\\Unity Hub.exe",
            "%LocalAppData%\\Programs\\Unity Hub\\Unity Hub.exe"
        )
    }
    @{
        Name = "Unreal Engine Launcher"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\Epic Games\\Launcher\\Engine\\Binaries\\Win64\\UnrealEditor.exe"
        )
    }
    @{
        Name = "Blitz (game client)"
        Category = "Gaming"
        Paths = @(
            "%LocalAppData%\\Programs\\Blitz\\Blitz.exe"
        )
    }
    @{
        Name = "SteamVR"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles(x86)%\\Steam\\steamapps\\common\\SteamVR\\bin\\win64\\vrstartup.exe"
        )
    }
    @{
        Name = "Epic Games (launcher alt)"
        Category = "Gaming"
        Paths = @(
            "%ProgramFiles%\\Epic Games\\Launcher\\Portal\\Binaries\\Win64\\EpicGamesLauncher.exe"
        )
    }
    @{
        Name = "OBS Studio (portable)"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\\Programs\\obs-studio\\bin\\64bit\\obs64.exe"
        )
    }
    @{
        Name = "Chromium"
        Category = "Browsers"
        Paths = @(
            "%ProgramFiles%\\Chromium\\Application\\chrome.exe",
            "%LocalAppData%\\Chromium\\Application\\chrome.exe"
        )
    }
    @{
        Name = "Brackets"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\Brackets\\Brackets.exe",
            "%LocalAppData%\\Programs\\Brackets\\Brackets.exe"
        )
    }
    @{
        Name = "CMake GUI"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\CMake\\bin\\cmake-gui.exe"
        )
    }
    @{
        Name = "Anaconda Navigator"
        Category = "Tools"
        Paths = @(
            "%LocalAppData%\\Continuum\\Anaconda3\\Scripts\\anaconda-navigator-script.py",
            "%ProgramData%\\Anaconda3\\Scripts\\anaconda-navigator-script.py"
        )
    }
    @{
        Name = "RStudio"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\RStudio\\bin\\rstudio.exe"
        )
    }
    @{
        Name = "MATLAB"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\MATLAB\\R2023a\\bin\\matlab.exe"
        )
    }
    @{
        Name = "SAS"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\SASHome\\SASFoundation\\9.4\\sas.exe"
        )
    }
    @{
        Name = "Tableau"
        Category = "Tools"
        Paths = @(
            "%ProgramFiles%\\Tableau\\Tableau Desktop\\bin\\tableau.exe"
        )
    }
    @{
        Name = "Slack (alternate)"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\\slack\\slack.exe"
        )
    }
    @{
        Name = "Signal (alt)"
        Category = "Comm"
        Paths = @(
            "%LocalAppData%\\Programs\\signal-desktop\\Signal.exe"
        )
    }
    @{
        Name = "Bitwarden (alt)"
        Category = "Security"
        Paths = @(
            "%LocalAppData%\\Programs\\Bitwarden\\Bitwarden.exe"
        )
    }
)

# --- Shared click handler for all app buttons ---
$buttonClickHandler = {
    $info = $this.Tag
    if (-not $info) { return }

    New-DesktopShortcut -Name $info.Name -Target $info.Target -Arguments $info.Arguments -IconPath $info.IconPath
}

# --- Create buttons only for installed apps ---
foreach ($app in $apps) {
    # Support two kinds of app records:
    #  - { Paths = @(...) } : Find-Executable as before
    #  - { Target = '...', Arguments = '...' } : use directly (useful for control.exe and ms-settings: URIs)

    $target = $null
    $arguments = $null
    $icon = $null

    if ($app.ContainsKey('Target')) {
        $target    = $app.Target
        if ($app.ContainsKey('Arguments')) { $arguments = $app.Arguments }
        if ($app.ContainsKey('IconPath'))   { $icon = $app.IconPath }

        # If Target is a filesystem path but doesn't exist, try falling back to Paths detection (if provided).
        $expandedTargetCheck = [Environment]::ExpandEnvironmentVariables($target)
        $isUriLike = ($expandedTargetCheck -and ($expandedTargetCheck -match '^[a-zA-Z][a-zA-Z0-9+\-.]*:' ) -and -not ($expandedTargetCheck -match '^[A-Za-z]:\\'))
        if (-not $isUriLike -and $expandedTargetCheck -and -not (Test-Path $expandedTargetCheck) -and $app.ContainsKey('Paths')) {
            $exe = Find-Executable -CandidatePaths $app.Paths
            if ($exe) { $target = $exe; $icon = $exe }
        }
    } elseif ($app.ContainsKey('Paths')) {
        $exe = Find-Executable -CandidatePaths $app.Paths
        if (-not $exe) { continue }   # not installed -> skip
        $target = $exe
        $icon   = $exe
    } else {
        continue
    }

    $button          = New-Object System.Windows.Controls.Button
    $button.Content  = $app.Name
    $button.Margin   = [System.Windows.Thickness]::new(5)
    $button.Padding  = [System.Windows.Thickness]::new(10,5,10,5)
    $button.Tag      = [PSCustomObject]@{
        Name     = $app.Name
        Target   = $target
        Arguments= $arguments
        IconPath = $icon
    }

    $button.Add_Click($buttonClickHandler)

    switch ($app.Category) {
        "Browsers" { $BrowsersPanel.Children.Add($button) | Out-Null }
        "Security" { $SecurityPanel.Children.Add($button) | Out-Null }
        "Gaming"   { $GamingPanel.Children.Add($button)   | Out-Null }
        "Windows"  { $WindowsPanel.Children.Add($button)  | Out-Null }
        "Comm"     { $CommPanel.Children.Add($button)     | Out-Null }
        "Tools"    { $ToolsPanel.Children.Add($button)    | Out-Null }
        default     { $ToolsPanel.Children.Add($button)    | Out-Null }
    }
}

# --- Close button ---
$CloseButton.Add_Click({
    $window.Close() | Out-Null
})

# If requested, run a PreviewAll mode (no .lnk files created) and print planned values.
if ($PreviewAll) {
    Write-Output "Previewing shortcuts (no files will be created)."
    foreach ($app in $apps) {
        $target = $null
        $arguments = $null
        $icon = $null

        if ($app.ContainsKey('Target')) {
            $target    = $app.Target
            if ($app.ContainsKey('Arguments')) { $arguments = $app.Arguments }
            if ($app.ContainsKey('IconPath'))   { $icon = $app.IconPath }

            $expandedTargetCheck = [Environment]::ExpandEnvironmentVariables($target)
            $isUriLike = ($expandedTargetCheck -and ($expandedTargetCheck -match '^[a-zA-Z][a-zA-Z0-9+\-.]*:' ) -and -not ($expandedTargetCheck -match '^[A-Za-z]:\\'))
            if (-not $isUriLike -and $expandedTargetCheck -and -not (Test-Path $expandedTargetCheck) -and $app.ContainsKey('Paths')) {
                $exe = Find-Executable -CandidatePaths $app.Paths
                if ($exe) { $target = $exe; $icon = $exe }
            }
        } elseif ($app.ContainsKey('Paths')) {
            $exe = Find-Executable -CandidatePaths $app.Paths
            if (-not $exe) { continue }
            $target = $exe
            $icon   = $exe
        } else {
            continue
        }

        $preview = New-DesktopShortcut -Name $app.Name -Target $target -Arguments $arguments -IconPath $icon -WhatIf
        if ($preview) {
            $preview | Format-List | Out-String | Write-Output
        } else {
            Write-Output "Skipped: $($app.Name) (no target found)"
        }
    }
    return
}

# --- Show window ---
if (-not $NoGui) { $window.ShowDialog() | Out-Null }
