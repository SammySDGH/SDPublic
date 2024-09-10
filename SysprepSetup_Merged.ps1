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
    Version: 3.8
    Requires: Administrative privileges
#>

[CmdletBinding()]
param (
    [string]$Method = ""
)

$SCRIPT_VERSION = "3.8"
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
    param ([string]$PcName, [string]$UnattendXmlPath)
    try {
        $unattendContent = Get-Content -Path $UnattendXmlPath -Raw
        $updatedUnattendContent = $unattendContent -replace "<ComputerName>PC-SD</ComputerName>", "<ComputerName>$PcName</ComputerName>"
        Set-Content -Path $newUnattendXML -Value $updatedUnattendContent -Force
        Write-Host "Updated unattend.xml with new PC name!" -ForegroundColor Green
    } catch { Write-Host "Failed to update unattend.xml: $_" -ForegroundColor Red; throw }
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
    if (-not (Test-AdminPrivileges)) { throw "This script requires administrative privileges. Please run as Administrator." }
    if (-not $Method) { $Method = if (Test-InternetConnection) { "Online" } else { "Offline" } }
    Write-Host "Running in $Method mode." -ForegroundColor Green
    $pcName = Get-PcNameFromUser
    if ($Method -eq "Online") {
        Install-NinjaAgent -Installer $ninjaInstaller_Online
        Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Online
    } elseif ($Method -eq "Offline") {
        if ($unattendXML_Offline -and $ninjaInstaller_Offline) {
            Install-NinjaAgent -Installer $ninjaInstaller_Offline
            Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Offline
        } else {
            Write-Host "Offline mode not available. No USB drive detected." -ForegroundColor Red
            $Method = "Online"
            Install-NinjaAgent -Installer $ninjaInstaller_Online
            Update-UnattendXml -PcName $pcName -UnattendXmlPath $unattendXML_Online
        }
    }
    Start-Sysprep
}

try { Main } catch { Write-Host "An error occurred during setup: $_" -ForegroundColor Red; throw }
