<#
.SYNOPSIS
Set Display Scaling based on the monitor's supported scaling range.

.DESCRIPTION
This script sets the display scaling percentage using WinAPI calls. 
It first checks the available scaling options for the monitor and only applies the scaling if it's supported by the monitor.

.PARAMETER Scaling
The desired display scaling value. 
Example values: "100%", "125%", "150%", "175%", "200%", "225%", "250%", "275%", "300%"

.EXAMPLE
PS> .\SetDisplayScaling.ps1 -Scaling "150%"

.NOTES
Author: Sammy Kastanja
Version: 2.4
#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("100%", "125%", "150%", "175%", "200%", "225%", "250%", "275%", "300%")]
    [string]$Scaling
)

# Map scaling values to their corresponding WinAPI values
$scalingMap = @{
    "100%" = 0
    "125%" = 1
    "150%" = 2
    "175%" = 3
    "200%" = 4
    "225%" = 5
    "250%" = 6
    "275%" = 7
    "300%" = 8
}

# Function to get the supported scaling options for the monitor
function Get-SupportedScalingOptions {
    $scalingOptions = @()

    # Use WMI with Get-WmiObject to get monitor resolutions and supported scaling
    $monitorInfo = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue
    if ($monitorInfo) {
        $maxResolution = ($monitorInfo.MaxHorizontalImageSize * $monitorInfo.MaxVerticalImageSize)
        
        # Higher resolution monitors support higher scaling
        if ($maxResolution -ge 3840 * 2160) {  # 4K or higher
            $scalingOptions += "100%", "125%", "150%", "175%", "200%", "225%", "250%", "275%", "300%"
        } elseif ($maxResolution -ge 2560 * 1440) {  # 1440p (2K)
            $scalingOptions += "100%", "125%", "150%", "175%", "200%"
        } elseif ($maxResolution -ge 1920 * 1080) {  # 1080p
            $scalingOptions += "100%", "125%", "150%"
        } else {
            $scalingOptions += "100%", "125%"  # Lower resolutions
        }
    } else {
        # Default to basic scaling options without outputting warnings
        $scalingOptions += "100%", "125%", "150%"
    }

    # Return supported scaling options
    return $scalingOptions
}

# Retrieve the corresponding WinAPI value for the selected scaling
$logPixels = $scalingMap[$Scaling]

# Get the list of supported scaling percentages
$supportedScalings = Get-SupportedScalingOptions

if (-not $supportedScalings.Contains($Scaling)) {
    Write-Error "The scaling option $Scaling% is not supported by the current monitor. Supported options: $($supportedScalings -join ', ')"
    exit 1
}

# Function to set display scaling using WinAPI call
function Set-Scaling {
    param($logPixels)

    $source = @'
    [DllImport("user32.dll", EntryPoint = "SystemParametersInfo")]
    public static extern bool SystemParametersInfo(
                  uint uiAction,
                  uint uiParam,
                  uint pvParam,
                  uint fWinIni);
'@
    $apicall = Add-Type -MemberDefinition $source -Name WinAPICall -Namespace SystemParamInfo -PassThru
    $apicall::SystemParametersInfo(0x009F, $logPixels, $null, 1) | Out-Null
}

# Apply the scaling
Set-Scaling -logPixels $logPixels