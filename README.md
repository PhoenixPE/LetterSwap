# LetterSwap

LetterSwap allows you to swap drive letters and/or synchronize letters of disks on based on the registry of the guest OS.

This is useful for mounting a boot drive in WinPE eg. `Y:`

The utility forcefully changes drive letters even if they are already used (letters abxyz are ignored), therefore it is desirable to launch letterSwap as soon as possible, for example through RunOnceEx.

This is a fork of the open source LetterSwap written by Nikzzzz and released on [TheOven.org](https://old.theoven.org/index0435.html?topic=93.0).

## Usage

USAGE: `LetterSwap.exe [/Swap <DriveLetter1>: <DriveLetter2>:] [/HideLetter|/MountAll] [/Auto|/Manual|/WinDir <Path>] [/BootDrive <NewLetter>:[\<TagFile>]] [/SetLetter <NewLetter>:\TagFile] [/Wait <Seconds>] [/Save] [/RestartExplorer] [/IgnoreLetter <Letters>] [/IgnoreCD] [/Log <LogFile>|con:]`

### Arguments
|Arg|Description|
|---|---|
|/HideLetter                         | Hides inactive removable and CDROM disks.|
|/MountAll                           | Mount inactive disks to the first available drive letter.|
|/Swap <DriveLetter1> <DriveLetter2> | Swap the specified drive letters. Ex. LetterSwap.exe /Swap D: E:|
|/Auto                               | Find the first guest OS.|
|/Manual                             | Display a dialog prompting to select the guest OS Windows directory.|
|/WinDir <Path>                      | Specify the directory of the guest OS. (Ex. D:\Windows)|
|/BootDrive <NewLetter>:             | Assigns the boot disk the specified drive letter.|
|/BootDrive <NewLetter>:\<TagFile>   | Search for <TagFile> and assign the boot disk the specified drive letter. Ex. Letterswap.exe /Auto /BootDrive Y:\USB.Y|
|/SetLetter <NewLetter>:\<TagFile>   | Search for <TagFile> and assign the disk the specified drive letter. Ex. Letterswap.exe /SetLetter Z:\File.tag|
|/Wait <Seconds>                     | Used in conjunction with /BootDrive or /SetLetter to define the number of seconds to wait for drives to become available.|
|/Save                               | When used in conjunction with /Auto, /Manual, or /WinDir save the Guest and Source drive letters to the registry (HKLM\SOFTWARE\LetterSwap).|
|/RestartExplorer                    | Restart Explorer.exe after letter change.|
|/IgnoreLetter <Letter>[...]         | When used in conjunction with /Auto, /Manual, or /WinDir ignore the specified drive letters. (The system drive and letters y,z are always ignored.) Ex. Letterswap.exe /IgnoreLetter abde|
|/IgnoreCD                           | When used in conjunction with /Auto, /Manual, or /WinDir ignore all driveletters belonging to CDROM drives.|
|/Log <LogFile>`|`con:               | Output to the specified log file or console.|
