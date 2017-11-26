#Requires -version 5.1
<#
    .SYNOPSIS 
    Sets a new device in 3CX

    .DESCRIPTION
    Sets a new device in 3CX

    .NOTES
    3CX seems to be unable to automatically set already used devices or use devices more appropriate than the primary sound device.
                
    .EXAMPLE
    .\Set-SoundDevice3CX.ps1

#>
# Version 0.5/26.11.2017
# Copyright ©2017 Bruno Holliger, Switzerland

# ****************************************************************************
# Configuration

# Mask for finding a microphone device
$microphoneDeviceMask = 'Microphone*Logitech*USB*Headset*'

# Mask for finding a speaker device
$speakerDeviceMask = 'Speakers*Logitech*USB*Headset*'

# Path to xml config file
$xml = "$env:APPDATA\3CXPhone for Windows\3CXPhone.xml"

# /Configuration
# ****************************************************************************

[xml]$config = Get-Content $xml -Encoding UTF8

$microphone = Get-PnpDevice | Where-Object {$_.InstanceId -like 'SWD\MMDEVAPI\*' -and $_.Status -eq 'OK' -and $_.Present -eq $true -and $_.FriendlyName -like $microphoneDeviceMask -and $_.PNPClass -eq 'AudioEndpoint'} | Select-Object FriendlyName, @{Name='DeviceGUID'; Expression={$_.InstanceID.Replace('SWD\MMDEVAPI\', '')}} | Select-Object -First 1
$speaker = Get-PnpDevice | Where-Object {$_.InstanceId -like 'SWD\MMDEVAPI\*' -and $_.Status -eq 'OK' -AND $_.Present -eq $true -and $_.FriendlyName -like $speakerDeviceMask -and $_.PNPClass -eq 'AudioEndpoint'} | Select-Object FriendlyName, @{Name='DeviceGUID'; Expression={$_.InstanceID.Replace('SWD\MMDEVAPI\', '')}} | Select-Object -First 1

Write-Host 'Setting the following devices:' -ForegroundColor Green
Write-Host "Speaker: $($speaker.FriendlyName)" -ForegroundColor Green
Write-Host "Microphone: $($microphone.FriendlyName)" -ForegroundColor Green

If ($speaker -and $microphone -and ($config.Accounts.General.MicrophoneDeviceGuid -ne $microphone.DeviceGUID.ToString().ToLower() -or $config.Accounts.General.SpeakerDeviceGuid -ne $speaker.DeviceGUID.ToString().ToLower())) {
    Write-Host "Closing 3CX"

    Get-Process 3CXWin8Phone | Foreach-Object { $_.CloseMainWindow() | Out-Null } | Out-Null
        
    $process = Get-Process 3CXWin8Phone

    $i = 0
    While ($process) {
        Write-Host Waiting for program to close...
        Start-Sleep -Milliseconds 500
        $process = Get-Process 3CXWin8Phone -ErrorAction SilentlyContinue
        if (++$i -ge 6) {
            $process = $false
        }
    }

    # If it's still here, kill it
    If (Get-Process 3CXWin8Phone -ErrorAction SilentlyContinue) {
        Get-Process 3CXWin8Phone | Stop-Process
    }

    # Set the devices and save the config file
    $config.Accounts.General.MicrophoneDeviceGuid = $microphone.DeviceGUID.ToString().ToLower()
    $config.Accounts.General.SpeakerDeviceGuid = $speaker.DeviceGUID.ToString().ToLower()

    $config.OuterXml | Set-Content $xml -Encoding UTF8

    Write-Host "Restarting 3CX"
    Start-Process "$env:ProgramData\3CXPhone for Windows\PhoneApp\3CXWin8Phone.exe" -WindowStyle Minimized
} Else {
    Write-Host 'No need to change a device.'
}
