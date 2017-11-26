#Requires -version 5.1
<#
    .SYNOPSIS 
    Get a list of sound devices

    .DESCRIPTION
    Get a list of sound devices

    .NOTES
    Get a list of sound devices
                
    .EXAMPLE
    .\Get-SoundDevice3CX.ps1

#>
# Version 0.5/26.11.2017
# Copyright ©2017 Bruno Holliger, Switzerland

# List all audio devices
Get-PnpDevice | Where-Object {$_.InstanceId -like 'SWD\MMDEVAPI\*' -and $_.Status -eq 'OK' -and $_.Present -eq $true -and $_.PNPClass -eq 'AudioEndpoint'} | Select-Object FriendlyName | Out-GridView

Write-Host 'Hit enter to close the window.'
Read-Host