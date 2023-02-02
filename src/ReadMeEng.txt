LetterSwap.exe - allows to synchronize letters of disks on the basis of the register of guest system.
So probably to appropriate to a Boot Disk letter Y: (it is searched on the marker file).
Syntax :
LetterSwap.exe [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/BootDrive NewLetter:\DriveMarkerFile] [/log LogFile] [/Letter RegExp] [/wait 10]
/? - help
WinDir - direct instructions of guest system, for example d:\windows
/Auto - finds the first guest system
/Manual - produces request
/bootdrive y:\MarkerFile appropriates to a disk with the marker file letter Y:, it is convenient, if on it there are the programs, demanding an absolute way.
/HideLetter - hides inactive disks
/MountAll - shows inactive disks

Example:
LetterSwap.exe /HideLetter /auto /bootdrive y:\MarkerFile /Log "%temp%\LSLog.txt"

The utility forcedly changes disks even if they are already used (letters abxyz are ignored), therefore it is desirable for launching as soon as possible, for example through RunOnceEx.
