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
    Version: 3.8.2
    Requires: Administrative privileges
#>

[CmdletBinding()]
param (
    [string]$Method = ""
)

$SCRIPT_VERSION = "3.8.2"
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
    try {
        Invoke-WebRequest -Uri "http://www.msftconnecttest.com/connecttest.txt" -UseBasicParsing -TimeoutSec 5 > $null
        return $true
    } catch {
        return $false
    }
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
    $regex = '^[^\\/:*?"<>|]{1,15}$'
    while ($true) {
        $newPcName = Read-Host "Enter the new PC name (Max 15 chars, no special chars \/:*?""<>|)"
        if ($newPcName -match $regex) {
            if (Confirm-Input -Message "Is the new PC name correct? ($newPcName)") {
                return $newPcName
            }
        } else {
            Write-Host "Invalid PC name. Please adhere to the naming conventions." -ForegroundColor Red
        }
    }
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
            Write-Host "$($locale.Number). $($locale.Language)" -ForegroundColor Yellow
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
        if (-not (Test-Path $Installer)) {
            Write-Host "Installer not found at $Installer" -ForegroundColor Red
            return
        }
        try {
            $process = Start-Process "msiexec.exe" -ArgumentList "/i `"$Installer`" /quiet" -NoNewWindow -PassThru
            Set-LoadingAnimation -Message "Installing NinjaOne Agent" -Process $process
            $process.WaitForExit()
        } catch {
            Write-Host "Failed to install NinjaOne Agent: $_" -ForegroundColor Red
            throw
        }
    } else {
        Write-Host "Skipping NinjaOne Agent installation." -ForegroundColor Yellow
    }
}

function Update-UnattendXml {
    param (
        [string]$PcName,
        [string]$UnattendXmlPath,
        [hashtable]$Locale,
        [string]$NewUnattendXmlPath
    )
    try {
        # Load the unattend.xml file as XML
        [xml]$xml = Get-Content -Path $UnattendXmlPath -Raw

        # Update the ComputerName
        $computerNameNode = $xml.SelectSingleNode("//component[@name='Microsoft-Windows-Shell-Setup']/ComputerName")
        if ($computerNameNode) {
            $computerNameNode.InnerText = $PcName
            Write-Host "Updated ComputerName in unattend.xml" -ForegroundColor Green
        } else {
            # Create the ComputerName node if it doesn't exist
            $shellSetupComponent = $xml.SelectSingleNode("//component[@name='Microsoft-Windows-Shell-Setup']")
            if ($shellSetupComponent) {
                $newElement = $xml.CreateElement("ComputerName")
                $newElement.InnerText = $PcName
                $shellSetupComponent.AppendChild($newElement) | Out-Null
                Write-Host "Added ComputerName to unattend.xml" -ForegroundColor Green
            } else {
                Write-Host "Microsoft-Windows-Shell-Setup component not found." -ForegroundColor Red
            }
        }

        # Update Locale Settings
        $intlComponent = $xml.SelectSingleNode("//component[@name='Microsoft-Windows-International-Core']")
        if ($intlComponent) {
            foreach ($localeSetting in 'InputLocale', 'SystemLocale', 'UILanguage', 'UserLocale') {
                $localeNode = $intlComponent.SelectSingleNode($localeSetting)
                if ($localeNode) {
                    $localeNode.InnerText = $Locale.$localeSetting
                } else {
                    # Create the locale node if it doesn't exist
                    $newElement = $xml.CreateElement($localeSetting)
                    $newElement.InnerText = $Locale.$localeSetting
                    $intlComponent.AppendChild($newElement) | Out-Null
                }
            }
            Write-Host "Updated locale settings in unattend.xml" -ForegroundColor Green
        } else {
            Write-Host "Microsoft-Windows-International-Core component not found." -ForegroundColor Red
        }

        # Save the updated unattend.xml
        $xml.Save($NewUnattendXmlPath)
        Write-Host "Saved updated unattend.xml to $NewUnattendXmlPath" -ForegroundColor Green
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
    } catch {
        Write-Error "Failed to start Sysprep: $_"
        throw
    }
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
        Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Online -Locale $locale -NewUnattendXmlPath $newUnattendXML
    } elseif ($Method -eq "Offline") {
        # Check if offline resources are available
        if ((Test-Path $unattendXML_Offline) -and (Test-Path $ninjaInstaller_Offline)) {
            # Install NinjaOne Agent from local storage
            Install-NinjaAgent -Installer $ninjaInstaller_Offline

            # Update unattend.xml with the new PC name and locale settings
            Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Offline -Locale $locale -NewUnattendXmlPath $newUnattendXML
        } else {
            Write-Host "Offline mode not available. Offline resources not found." -ForegroundColor Red
            # Fallback to Online mode
            $Method = "Online"
            Install-NinjaAgent -Installer $ninjaInstaller_Online
            Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Online -Locale $locale -NewUnattendXmlPath $newUnattendXML
        }
    } else {
        Write-Host "Invalid method specified. Please use 'Online' or 'Offline'." -ForegroundColor Red
        return
    }

    # Start Sysprep with the updated unattend.xml
    Start-Sysprep
}

try {
    Main
} catch {
    Write-Host "An error occurred during setup: $_" -ForegroundColor Red
    throw
}
