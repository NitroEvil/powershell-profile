# Ensure the script can run with elevated privileges
$currentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ($currentUser.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
    Write-Warning "Please run this script as an Administrator!"
    break
}

# Check for internet connectivity before proceeding
try {
    Test-Connection -ComputerName "www.google.com" -Count 1 -ErrorAction Stop | Out-Null
} catch {
    Write-Warning "Internet connection is required but not available. Please check your connection."
    break
}

# Profile creation or update
if (-not (Test-Path -Path $PROFILE -PathType Leaf)) {
    try {
        # Detect Version of PowerShell & Create Profile directories if they do not exist.
        $profilePath = ""
        if ($PSVersionTable.PSEdition -eq "Core") { 
            $profilePath = [Environment]::GetFolderPath("MyDocuments") + "\Powershell"
        } elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = [Environment]::GetFolderPath("MyDocuments") + "\WindowsPowerShell"
        }

        if (-not (Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory"
        }

        $options = @{
            Uri     = "https://github.com/NitroEvil/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1"
            OutFile = $PROFILE
        }
        Invoke-RestMethod @options
        Write-Host "The profile @ [$PROFILE] has been created."
        Write-Host "If you want to add any persistent components, please do so at [$profilePath\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    } catch {
        Write-Error "Failed to create or update the profile. Error: $_"
    }
} else {
    try {
        Get-Item -Path $PROFILE | Move-Item -Destination "oldprofile.ps1" -Force
        $options = @{
            Uri     = "https://github.com/NitroEvil/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1"
            OutFile = $PROFILE
        }
        Invoke-RestMethod @options
        Write-Host "The profile @ [$PROFILE] has been created and old profile removed."
        Write-Host "Please back up any persistent components of your old profile to [$HOME\Documents\PowerShell\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    } catch {
        Write-Error "Failed to backup and update the profile. Error: $_"
    }
}

# OhMyPosh Install
try {
    winget install -e --accept-source-agreements --accept-package-agreements JanDeDobbeleer.OhMyPosh
} catch {
    Write-Error "Failed to install Oh My Posh. Error: $_"
}

# Font Install
$fontName = "MesloLGS Nerd Font"
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
$fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
try {
    if ($fontFamilies -notcontains "MesloLGS Nerd Font") {
        $fontDownloadName = "Meslo"
        $options = @{
            Uri     = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/$fontDownloadName.zip"
            OutFile = ".\$fontDownloadName.zip"
        }
        Invoke-RestMethod @options

        Expand-Archive -Path ".\$fontDownloadName.zip" -DestinationPath ".\$fontDownloadName" -Force
        $destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
        Get-ChildItem -Path ".\$fontDownloadName" -Recurse -Filter "*.ttf" | ForEach-Object {
            If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                $destination.CopyHere($_.FullName, 0x10)
            }
        }

        Remove-Item -Path ".\$fontDownloadName" -Recurse -Force
        Remove-Item -Path ".\$fontDownloadName.zip" -Force
    }
} catch {
    Write-Error "Failed to download or install the Cascadia Code font. Error: $_"
}

# Final check and message to the user
if ((Test-Path -Path $PROFILE) -and (winget list --name "OhMyPosh" -e) -and ($fontFamilies -contains $fontName)) {
    Write-Host "Setup completed successfully. Please restart your PowerShell session to apply changes."
} else {
    Write-Warning "Setup completed with errors. Please check the error messages above."
}

# Terminal Icons Install
try {
    Install-Module -Name Terminal-Icons -Repository PSGallery -Force
} catch {
    Write-Error "Failed to install Terminal Icons module. Error: $_"
}

# zoxide Install
try {
    winget install -e --id ajeetdsouza.zoxide
    Write-Host "zoxide installed successfully."
} catch {
    Write-Error "Failed to install zoxide. Error: $_"
}
