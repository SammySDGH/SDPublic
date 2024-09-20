<#
.SYNOPSIS
    Automated Windows 11 Device Setup Script (Online/Offline)
.DESCRIPTION
    This script automates the process of setting up a new Windows 11 PC. 
    It will automatically detect whether an internet connection is available and run the appropriate method (Online or Offline). 
    Optionally, a -Method parameter can be used to manually select the method.
.NOTES
    Author: Sammy Kastanja
    Date: 09-09-2024
    Version: 3.8.1
    Requires: Administrative privileges
#>

[CmdletBinding()]
param (
    [string]$Method = ""
)

$SCRIPT_VERSION = "3.8.1"
$ASCII_LOGO = @"
       _____            _       __   ____             __
      / ___/____  _____(_)___ _/ /  / __ \___  ____ _/ /
      \__ \/ __ \/ ___/ / __  / /  / / / / _ \/ __  / /
     ___/ / /_/ / /__/ / /_/ / /  / /_/ /  __/ /_/ / /  
    /____/\____/\___/_/\__,_/_/  /_____/\___/\__,_/_/   
       
        Device Setup Script V$SCRIPT_VERSION by Sammy Kastanja

"@

$networkPath = "\\10.1.6.246\Scripts\DeviceSetup"
$sourcePath = if ($PSScriptRoot) { $PSScriptRoot } else { "" }

$unattendXML_Online = "$networkPath\unattend.xml"
$unattendXML_Offline = if ($sourcePath) { Join-Path -Path $sourcePath -ChildPath "unattend.xml" } else { "" }
$newUnattendXML = "C:\Windows\System32\Sysprep\unattend.xml"
$ninjaInstaller_Online = "$networkPath\NinjaAgentInstaller.msi"
$ninjaInstaller_Offline = if ($sourcePath) { Join-Path -Path $sourcePath -ChildPath "NinjaAgentInstaller.msi" } else { "" }
$sysprepPath = "C:\Windows\System32\Sysprep\Sysprep.exe"

function Test-InternetConnection {
    try { Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet } catch { $false }
}

function Set-LoadingAnimation {
    param ($Message, [System.Diagnostics.Process]$Process)
    $spinnerChars = @("   ", ".  ", ".. ", "...")
    $currentIndex = 0
    while (-not $Process.HasExited) {
        Write-Host -NoNewline "`r$Message$($spinnerChars[$currentIndex])" -ForegroundColor Cyan
        $currentIndex = ($currentIndex + 1) % $spinnerChars.Length
        Start-Sleep -Milliseconds 200
    }
    Write-Host "`rFinished $Message!" -ForegroundColor Green
}

function Confirm-Input {
    param ([string]$Message)
    while ($true) {
        $keyinput = Read-Host "$Message (Y/N)"
        if ($keyinput -eq "Y" -or $keyinput -eq "y") { return $true }
        elseif ($keyinput -eq "N" -or $keyinput -eq "n") { return $false }
        else { Write-Host "Invalid input. Please enter 'Y' or 'N'." -ForegroundColor Red }
    }
}

function Get-PcNameFromUser {
    $newPcName = Read-Host "Enter the new PC name"
    if (Confirm-Input -Message "Is the new PC name correct? ($newPcName)") { return $newPcName }
    else { Write-Host "Please re-enter the new PC name." -ForegroundColor Yellow; return Get-PcNameFromUser }
}

function Get-Locale {
    [CmdletBinding()]
    param()

    $locales = @(
        @{
            Number = 1
            Language = 'Dutch'
            InputLocale = '0413:00000413;0409:00000409;0407:00000407'
            SystemLocale = 'nl-NL'
            UILanguage = 'nl-NL'
            UserLocale = 'nl-NL'
        },
        @{
            Number = 2
            Language = 'English'
            InputLocale = '0409:00000409;0407:00000407;080c:0000080c'
            SystemLocale = 'en-US'
            UILanguage = 'en-US'
            UserLocale = 'en-US'
        },
        @{
            Number = 3
            Language = 'German'
            InputLocale = '0407:00000407;0409:00000409;080c:0000080c'
            SystemLocale = 'de-DE'
            UILanguage = 'de-DE'
            UserLocale = 'de-DE'
        },
        @{
            Number = 4
            Language = 'Belgian'
            InputLocale = '080c:0000080c;0409:00000409;0407:00000407'
            SystemLocale = 'fr-BE'
            UILanguage = 'fr-BE'
            UserLocale = 'fr-BE'
        },
        @{
            Number = 5
            Language = 'French'
            InputLocale = '040c:0000040c;0409:00000409;0407:00000407'
            SystemLocale = 'fr-FR'
            UILanguage = 'fr-FR'
            UserLocale = 'fr-FR'
        }
    )

    while ($true) {
        Write-Host "Please select a language:"
        foreach ($locale in $locales) {
            Write-Host "$($locale.Number). $($locale.Language)"
        }

        $selection = Read-Host "Enter the number corresponding to your choice (1-5)"
        if ($selection -match '^[1-5]$') {
            $selectedLocale = $locales | Where-Object { $_.Number -eq [int]$selection }
            if ($selectedLocale) {
                if (Confirm-Input -Message "You have selected $($selectedLocale.Language). Is this correct?") {
                    return $selectedLocale
                }
            }
        } else {
            Write-Host "Invalid selection. Please enter a number between 1 and 5." -ForegroundColor Red
        }
    }
}

function Install-NinjaAgent {
    param ([string]$Installer)
    if (Confirm-Input -Message "Install NinjaOne Agent?") {
        try {
            $process = Start-Process "msiexec.exe" -ArgumentList "/i `"$Installer`" /quiet" -NoNewWindow -PassThru
            Set-LoadingAnimation -Message "Installing NinjaOne Agent" -Process $process
            $process.WaitForExit()
        } catch { Write-Host "Failed to install NinjaOne Agent: $_" -ForegroundColor Red; throw }
    } else { Write-Host "Skipping NinjaOne Agent installation." -ForegroundColor Yellow }
}

function Update-UnattendXml {
    param (
        [string]$PcName,
        [string]$UnattendXmlPath,
        [hashtable]$Locale
    )
    try {
        [xml]$xml = Get-Content -Path $UnattendXmlPath -Raw

        # Update the ComputerName
        $shellSetupComponents = $xml.unattend.settings | Where-Object {
            $_.component.'@name' -eq 'Microsoft-Windows-Shell-Setup'
        }
        foreach ($settings in $shellSetupComponents) {
            foreach ($component in $settings.component) {
                if ($component.ComputerName) {
                    $component.ComputerName = $PcName
                    Write-Host "Updated ComputerName in unattend.xml" -ForegroundColor Green
                }
            }
        }

        # Update Locale Settings
        $intlComponents = $xml.unattend.settings | Where-Object {
            $_.component.'@name' -eq 'Microsoft-Windows-International-Core'
        }
        foreach ($settings in $intlComponents) {
            foreach ($component in $settings.component) {
                if ($component.InputLocale -or $component.SystemLocale -or $component.UILanguage -or $component.UserLocale) {
                    $component.InputLocale = $Locale.InputLocale
                    $component.SystemLocale = $Locale.SystemLocale
                    $component.UILanguage = $Locale.UILanguage
                    $component.UserLocale = $Locale.UserLocale
                    Write-Host "Updated locale settings in unattend.xml" -ForegroundColor Green
                }
            }
        }

        # Save the updated unattend.xml
        $xml.Save($newUnattendXML)
        Write-Host "Updated unattend.xml with new PC name and locale settings!" -ForegroundColor Green
    } catch {
        Write-Host "Failed to update unattend.xml: $_" -ForegroundColor Red
        throw
    }
}


function Start-Sysprep {
    try {
        $process = Start-Process -FilePath $sysprepPath -ArgumentList "/generalize", "/oobe", "/reboot", "/quiet", "/unattend:$newUnattendXML" -WindowStyle Hidden -PassThru
        Set-LoadingAnimation -Message "Running Sysprep. Please be patient, system will reboot automatically" -Process $process
        $process.WaitForExit()
    } catch { Write-Error "Failed to start Sysprep: $_"; throw }
}

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Main {
    Write-Host $ASCII_LOGO -ForegroundColor Cyan

    # Check for administrative privileges
    if (-not (Test-AdminPrivileges)) {
        throw "This script requires administrative privileges. Please run as Administrator."
    }

    # Determine the method if not specified
    if (-not $Method) {
        $Method = if (Test-InternetConnection) { "Online" } else { "Offline" }
    }
    Write-Host "Running in $Method mode." -ForegroundColor Green

    # Get the new PC name from the user
    $pcName = Get-PcNameFromUser

    # Get the locale settings from the user
    $locale = Get-Locale

    # Execute based on the selected method
    if ($Method -eq "Online") {
        # Install NinjaOne Agent from the network
        Install-NinjaAgent -Installer $ninjaInstaller_Online

        # Update unattend.xml with the new PC name and locale settings
        Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Online -Locale $locale
    } elseif ($Method -eq "Offline") {
        # Check if offline resources are available
        if ($unattendXML_Offline -and $ninjaInstaller_Offline) {
            # Install NinjaOne Agent from local storage
            Install-NinjaAgent -Installer $ninjaInstaller_Offline

            # Update unattend.xml with the new PC name and locale settings
            Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Offline -Locale $locale
        } else {
            Write-Host "Offline mode not available. No USB drive detected." -ForegroundColor Red
            # Fallback to Online mode
            $Method = "Online"
            Install-NinjaAgent -Installer $ninjaInstaller_Online
            Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Online -Locale $locale
        }
    } else {
        Write-Host "Invalid method specified. Please use 'Online' or 'Offline'." -ForegroundColor Red
        return
    }

    # Start Sysprep with the updated unattend.xml
    Start-Sysprep
}


try { Main } catch { Write-Host "An error occurred during setup: $_" -ForegroundColor Red; throw }