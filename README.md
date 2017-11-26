# Set-SoundDevice3CX
Powershell Script to set an appropriate sound device for 3CX softphone client.

## System requirements
- This script runs on Windows 10.
- Script execution policy needs to be at least [RemoteSigned](https://technet.microsoft.com/en-us/library/ee176961.aspx)

## Installation
1. Download the script from the repository.
2. Store it in a folder.

## Configuration of the script
1. Start the script Get-SoundDevice3CX.ps1
2. Pick the device names for microphone and speakers you would like to use.
3. Create a mask that fits all your devices.
   Example: This mask: Speakers\*Logitech\*USB\*Headset\* works for: Speakers (6- Logitech USB Headset)
4. Put the masks for microphone and speaker into the config section of the script.

## Configuration of the task scheduler
1. Copy the file path to clipboard.
2. Start task scheduler.
3. Create a new task.
4. Give it an appropriate name.
5. Configure for: Windows 10.
6. Change to "Triggers".
7. Add the following triggers:
   - At logon
   - On workstation unlock
8. Switch to "Actions"
9. Enter 'powershell.exe' in "Program/script".
10. Enter '-file "c:\mytools\Set-SoundDevice3CX\Set-SoundDevice3CX.ps1"'.
11. Start in only the path to the ps1 file.
12. Hit Ok.
13. Hit Ok again.

Lock your workstation and check if the device is correctly set in 3CX.

Have fun!
