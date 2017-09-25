# // ***********************************************************************************************************
# // Copyright (c) Florian VALENTE.  All rights reserved.
# // Microsoft Deployment Toolkit Solution Accelerator
# //
# // Version:   1.0 (14/02/26)
# // 
# // Purpose:   Install HP SPP
# // Usage:     InstallSPP.ps1
# // Remark:    HP SPP sources must be in "Sources" sub-folder
# // This script is provided "AS IS" with no warranties
# // ***********************************************************************************************************

$ErrorActionPreference = 'Stop'

# Get Script root
$PSScriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path

# Set HP SPP Default local Installation Path
$DefaultInstallPath = "$env:TEMP\SPP"

# Set HP SUM Setup file
$Setup = $tsenv:Architecture + "\hpsum_bin_" + $tsenv:Architecture + ".exe"
# Set HP SUM Parameters
$Parameters = "/s /override_existing_connection /use_snmp /allow_update_to_bundle /allow_non_bundle_components /current_credential /continue_on_error FailedDependencies"
# Set HP SPP Version
$Version = "2014.02"


# Function to copy SPP under local location
Function Copy-SourcesLocally {
    If (Test-Path $DefaultInstallPath) {
        Remove-Item -Path $DefaultInstallPath -Recurse -Force | Out-Null
        If (!$?) {
            Write-Error "ERROR: Folder $DefaultInstallPath can't be removed!"
        }
        Else {
            Write-Host "Folder $DefaultInstallPath removed."
        }
    }

    # Create install folder
    New-Item -Path $DefaultInstallPath -ItemType directory -Force | Out-Null

    # Copy items
    Write-Host "Copying $setupPath to $DefaultInstallPath..."
    Copy-Item "$setupPath\*" -Destination $DefaultInstallPath -Recurse -Force
    If (!$?) {
        Write-Error "ERROR: HP SPP files cannot be copied locally!"
    }
    Else {
        Write-Host "HP SPP Installation files copied successfully!"
    }
}


# Function to install HP SPP
Function Install-SPP {
    Write-Host "Setup: $Setup"
    Write-Host "Parameters: $Parameters"

    Write-Host "Installing HP SPP $Version..."
    Write-Host "Summary logs are available in $($env:SystemDrive)\cpqsystem\hp\log"
    Write-Host "Detailed logs are available in $($env:SystemDrive)\cpqsystem\hp\log\localhost"
    $objStatus = Start-Process -FilePath "$DefaultInstallPath\$Setup" -ArgumentList "$Parameters" -PassThru -Wait
    If ($objStatus.ExitCode -eq 0) {
        Write-Host "HP SPP installed successfully."
    }
    ElseIf ($objStatus.ExitCode -eq 1) {
        Write-Host "HP SPP installed successfully but a reboot is required"
    }
    ElseIf ($objStatus.ExitCode -eq 3) {
        Write-Host "HP SPP installed successfully but the component was current or not required."
    }
    ElseIf ($objStatus.ExitCode -eq -1) {
        Write-Warning "HP SPP installed with errors! Consult logs in in $($env:SystemDrive)\cpqsystem\hp\log"
    }
    ElseIf ($objStatus.ExitCode -eq -2) {
        Write-Warning "HP SPP not installed because bad input parameter was encountered! Consult logs in in $($env:SystemDrive)\cpqsystem\hp\log"
    }
    ElseIf ($objStatus.ExitCode -eq -3) {
        Write-Warning "HP SPP installed with components installation errors! Consult logs in in $($env:SystemDrive)\cpqsystem\hp\log"
    }
    Else {
        Clean-Installation
        Write-Error "HP SPP not installed successfully! (Err: $($objStatus.ExitCode), Desc: $($objStatus.Description))"
    }
}


# Function to clean HP SPP Install folder in case of error or at the end of installation
Function Clean-Installation {
    Remove-Item -Path "$DefaultInstallPath" -Recurse -Force
    If (!$?) {
        Write-Error "ERROR: Installation cleanup failed!"
    }
    Else {
        Write-Host "Installation cleaned successfully"
    }
}



########
# MAIN #
########
If ($tsenv:Make -eq "HP") {
    # Set path of the Setup file
    $Global:setupPath = "$PSScriptRoot\Sources"

    # Test if Setup path exists
    If (!(Test-Path "$setupPath\$Setup")) {
        Write-Error "ERROR: $setupPath folder or $Setup not exists! Installation aborted."
    }

    # Copy sources locally
    Copy-SourcesLocally

    # Install HP SPP
    Install-SPP

    # Cleanup the installation
    Clean-Installation

    # Sleep needed to really wait the end of the installation
    Start-Sleep 60
}
Else {
    Write-Warning "The Manufacturer detected is $($tsenv:Make) and not HP! So SPP $Version can't be installed"
}
