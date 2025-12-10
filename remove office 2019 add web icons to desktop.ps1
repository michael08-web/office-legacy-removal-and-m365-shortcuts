Write-Output '===== Office Cleanup + O365 Web Shortcuts Script Starting ====='

# -------------------------------
# 0. Ensure local icon folder & copy icons from NAS share
# -------------------------------
$iconSourceFolder = '\\fileserver\share\Office Icons'
$iconFolder       = 'C:\ProgramData\OfficeWebIcons'

if (-not (Test-Path $iconFolder)) {
    try {
        New-Item -Path $iconFolder -ItemType Directory -Force | Out-Null
        Write-Output ('Created local icon folder: ' + $iconFolder)
    } catch {
        Write-Output ('WARNING: Could not create ' + $iconFolder + '. ' + $_.Exception.Message)
    }
}

if (Test-Path $iconSourceFolder) {
    try {
        $sourcePattern = Join-Path $iconSourceFolder '*.ico'
        Copy-Item -Path $sourcePattern -Destination $iconFolder -Force -ErrorAction SilentlyContinue
        Write-Output ('Copied .ico files from "' + $iconSourceFolder + '" to "' + $iconFolder + '" (if any).')
    } catch {
        Write-Output ('WARNING: Failed to copy icons from NAS share. ' + $_.Exception.Message)
    }
} else {
    Write-Output ('Icon source folder "' + $iconSourceFolder + '" not found; continuing without copying icons.')
}

# -------------------------------
# 1. Detect existing Microsoft 365 Apps (Click-to-Run)
# -------------------------------
$CTRConfig = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$M365Found = $false

if (Test-Path $CTRConfig) {
    try {
        $ProductProps = Get-ItemProperty -Path $CTRConfig -ErrorAction Stop
        $ProductIds   = $ProductProps.ProductReleaseIds

        if ($ProductIds) {
            foreach ($id in $ProductIds) {
                if ($id -match 'O365|M365|ProPlus|Business|MicrosoftOffice|Microsoft 365') {
                    $M365Found = $true
                    break
                }
            }
        }
    } catch {
        # detection failure -> assume not found
    }
}

if ($M365Found) {
    Write-Output 'Microsoft 365 Apps detected.'
} else {
    Write-Output 'No Microsoft 365 Apps detected.'
}

# -------------------------------
# 2. Cleanup old Office desktop shortcuts
#    - Always clean user desktops
#    - Only clean Public Desktop if M365 is NOT installed
# -------------------------------
Write-Output 'Cleaning up old Office shortcuts from desktops...'

$publicDesktop = 'C:\Users\Public\Desktop'
$desktopRoots  = @()

# User desktops
try {
    $userDirs = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue
    foreach ($u in $userDirs) {
        $desk = Join-Path $u.FullName 'Desktop'
        $desktopRoots += $desk
    }
} catch {}

# Public desktop ONLY if M365 is NOT installed
if (-not $M365Found) {
    $desktopRoots += $publicDesktop
} else {
    Write-Output 'M365 detected: leaving Public Desktop shortcuts untouched.'
}

# Any .lnk with these patterns in the name gets removed
$patterns = @('*Word*.lnk','*Excel*.lnk','*PowerPoint*.lnk','*Outlook*.lnk','*Publisher*.lnk')

foreach ($root in $desktopRoots) {
    if (Test-Path $root) {
        try {
            Get-ChildItem -Path $root -Recurse -Include $patterns -File -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        } catch {
            # ignore per-file errors
        }
    }
}

Write-Output 'Old Office shortcuts cleaned where applicable.'

# -------------------------------
# 3. If Microsoft 365 Apps are installed, stop here
# -------------------------------
if ($M365Found) {
    Write-Output 'Microsoft 365 Apps present – skipping Office uninstall and web shortcut creation.'
    Write-Output '===== Script Complete (user desktop cleanup + icon copy only) ====='
    exit 0
}

# -------------------------------
# 4. Office Deployment Tool path (for Office 2019 removal)
# -------------------------------
$setupExe         = 'C:\Office2019\setup.exe'
$odtFolder        = 'C:\Office2019'
$uninstallXmlPath = Join-Path $odtFolder 'Uninstall-Office2019.xml'

if (-not (Test-Path $setupExe)) {
    Write-Output 'ERROR: C:\Office2019\setup.exe not found. Cannot uninstall Office 2019. Exiting.'
    Write-Output '===== Script Complete (no uninstall performed) ====='
    exit 1
}

# -------------------------------
# 5. Create uninstall XML (Remove All) without here-strings
# -------------------------------
if (-not (Test-Path $uninstallXmlPath)) {
    Write-Output ('Creating uninstall config XML at ' + $uninstallXmlPath)

    $xmlLines = @(
        '<Configuration>',
        '  <Remove All="TRUE" />',
        '  <Display Level="None" AcceptEULA="TRUE" />',
        '</Configuration>'
    )
    $xmlLines | Set-Content -Path $uninstallXmlPath -Encoding UTF8
} else {
    Write-Output ('Uninstall XML already exists at ' + $uninstallXmlPath)
}

# -------------------------------
# 6. Uninstall Office 2019 (handled by ODT)
# -------------------------------
Write-Output 'Running Office Deployment Tool to remove Office 2019...'

$arguments = '/configure "' + $uninstallXmlPath + '"'

$proc = Start-Process -FilePath $setupExe -ArgumentList $arguments -Wait -PassThru -NoNewWindow

if ($proc.ExitCode -ne 0) {
    Write-Output ('WARNING: Office uninstall returned exit code ' + $proc.ExitCode)
} else {
    Write-Output 'Office 2019 uninstall command completed.'
}

# -------------------------------
# 7. Helper: icon selection for shortcuts
# -------------------------------
function Get-ShortcutIconLocation {
    param(
        [Parameter(Mandatory = $true)] $Item,
        [Parameter(Mandatory = $false)] [string]$EdgePath,
        [Parameter(Mandatory = $false)] [string]$IconFolder
    )

    if ($IconFolder -and $Item.Icon) {
        $icoPath = Join-Path $IconFolder $Item.Icon
        if (Test-Path $icoPath) {
            return ($icoPath + ',0')
        }
    }

    if ($EdgePath -and (Test-Path $EdgePath)) {
        return ($EdgePath + ',0')
    }

    return ($env:SystemRoot + '\system32\SHELL32.dll,1')
}

# -------------------------------
# 8. Create Office 365 Web Shortcuts on Public Desktop
# -------------------------------
Write-Output 'Creating Office 365 web shortcuts on Public Desktop...'

if (-not (Test-Path $publicDesktop)) {
    New-Item -Path $publicDesktop -ItemType Directory -Force | Out-Null
}

# Try to locate Edge
$edgePaths = @(
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
)

$edgePath = $null
foreach ($p in $edgePaths) {
    if (Test-Path $p) { $edgePath = $p; break }
}

# Define shortcuts as objects
$shortcutDefinitions = @()
$shortcutDefinitions += New-Object PSObject -Property @{ Name='Word Online';       Url='https://www.office.com/launch/word';       Icon='word.ico' }
$shortcutDefinitions += New-Object PSObject -Property @{ Name='Excel Online';      Url='https://www.office.com/launch/excel';      Icon='excel.ico' }
$shortcutDefinitions += New-Object PSObject -Property @{ Name='PowerPoint Online'; Url='https://www.office.com/launch/powerpoint'; Icon='powerpoint.ico' }
$shortcutDefinitions += New-Object PSObject -Property @{ Name='Outlook Web';       Url='https://outlook.office.com/';              Icon='outlook.ico' }
$shortcutDefinitions += New-Object PSObject -Property @{ Name='OneNote Online';    Url='https://www.office.com/launch/onenote';    Icon='onenote.ico' }
$shortcutDefinitions += New-Object PSObject -Property @{ Name='OneDrive Online';   Url='https://onedrive.live.com/';               Icon='onedrive.ico' }
$shortcutDefinitions += New-Object PSObject -Property @{ Name='SharePoint Online'; Url='https://www.microsoft365.com/sharepoint/'; Icon='sharepoint.ico' }

# Create .lnk shortcuts if Edge exists, else .url
if ($edgePath) {
    Write-Output ('Edge found at ' + $edgePath + '. Creating .LNK shortcuts with icons...')

    $shell = New-Object -ComObject WScript.Shell

    foreach ($item in $shortcutDefinitions) {
        $lnkPath = Join-Path $publicDesktop ($item.Name + '.lnk')
        Write-Output ('Creating .lnk: ' + $lnkPath)

        $iconLocation = Get-ShortcutIconLocation -Item $item -EdgePath $edgePath -IconFolder $iconFolder

        $shortcut = $shell.CreateShortcut($lnkPath)
        $shortcut.TargetPath       = $edgePath
        $shortcut.Arguments        = $item.Url
        $shortcut.IconLocation     = $iconLocation
        $shortcut.WorkingDirectory = (Split-Path $edgePath)
        $shortcut.Save()
    }
}
else {
    Write-Output 'Edge not found. Creating .URL shortcuts with icons...'

    foreach ($item in $shortcutDefinitions) {
        $urlPath = Join-Path $publicDesktop ($item.Name + '.url')
        Write-Output ('Creating .url: ' + $urlPath)

        $iconLocation = Get-ShortcutIconLocation -Item $item -EdgePath $null -IconFolder $iconFolder
        $iconParts = $iconLocation -split ',', 2
        $iconFile  = $iconParts[0]
        $iconIndex = if ($iconParts.Count -gt 1) { $iconParts[1] } else { '0' }

        $urlContent = '[InternetShortcut]' + "`r`n" +
                      'URL=' + $item.Url + "`r`n" +
                      'IconFile=' + $iconFile + "`r`n" +
                      'IconIndex=' + $iconIndex + "`r`n"

        Set-Content -Path $urlPath -Value $urlContent -Encoding ASCII
    }
}

Write-Output 'Office 365 web shortcuts created.'
Write-Output '===== Script Complete ====='