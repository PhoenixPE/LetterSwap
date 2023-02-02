Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagGUID = "ulong Data1;ushort Data2;ushort Data3;byte Data4[8]"
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $__WINVER = __Ver()
Global Const $tagNOTIFYICONDATA = 'dword Size;hwnd hWnd;uint ID;uint Flags;uint CallbackMessage;ptr hIcon;wchar Tip[128];dword State;dword StateMask;wchar Info[256];uint Version;wchar InfoTitle[64];dword InfoFlags;'
Global Const $tagPRINTDLG = 'align 2;dword_ptr Size;hwnd hOwner;ptr hDevMode;ptr hDevNames;hwnd hDC;dword Flags;ushort FromPage;ushort ToPage;ushort MinPage;ushort MaxPage;' & __Iif(@AutoItX64, 'uint', 'ushort') & ' Copies;ptr hInstance;lparam lParam;ptr PrintHook;ptr SetupHook;ptr PrintTemplateName;ptr SetupTemplateName;ptr hPrintTemplate;ptr hSetupTemplate;'
Func _WinAPI_DeleteVolumeMountPoint($sPath)
Local $Ret = DllCall('kernel32.dll', 'int', 'DeleteVolumeMountPointW', 'wstr', $sPath)
If(@error) Or(Not $Ret[0]) Then
Return SetError(1, 0, 0)
EndIf
Return 1
EndFunc
Func _WinAPI_GetVolumeNameForVolumeMountPoint($sPath)
Local $Ret = DllCall('kernel32.dll', 'int', 'GetVolumeNameForVolumeMountPointW', 'wstr', $sPath, 'wstr', '', 'dword', 80)
If(@error) Or(Not $Ret[0]) Then
Return SetError(1, 0, '')
EndIf
Return $Ret[2]
EndFunc
Func _WinAPI_SetVolumeMountPoint($sPath, $GUID)
Local $Ret = DllCall('kernel32.dll', 'int', 'SetVolumeMountPointW', 'wstr', $sPath, 'wstr', $GUID)
If(@error) Or(Not $Ret[0]) Then
Return SetError(1, 0, 0)
EndIf
Return 1
EndFunc
Func __Iif($fTest, $iTrue, $iFalse)
If $fTest Then
Return $iTrue
Else
Return $iFalse
EndIf
EndFunc
Func __Ver()
Local $tOSVI, $Ret
$tOSVI = DllStructCreate('dword Size;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128]')
DllStructSetData($tOSVI, 'Size', DllStructGetSize($tOSVI))
$Ret = DllCall('kernel32.dll', 'int', 'GetVersionExW', 'ptr', DllStructGetPtr($tOSVI))
If(@error) Or(Not $Ret[0]) Then
Return SetError(1, 0, 0)
EndIf
Return BitOR(BitShift(DllStructGetData($tOSVI, 'MajorVersion'), -8), DllStructGetData($tOSVI, 'MinorVersion'))
EndFunc
Opt('TrayIconHide', 1)
$sAboot = "                @Nikzzzz 30.12.2011"
$sHelp = @ScriptName & " [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/BootDrive NewLetter:\DriveMarkerFile] [/log LogFile] [/IgnoreLetter Letters] [/wait 10]" & @CRLF
Global $sReg = "HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices"
If $CmdLine[0] = 0 Then
EndIf
Dim $aMountHost[1][2]
Dim $aMountGuest[1][2]
$sIgnoreLetter = 'yz'
$sLogFile = ''
$sHiveSystemGuest = ''
$sBootDrive = ''
$sNewBootDrive = ''
$s = ""
$i = 1
$sWait = 1000
$sLoopWait = 1
While $i <= $CmdLine[0]
$vTemp = StringLower($CmdLine[$i])
Switch $vTemp
Case "/auto"
$aDrives = DriveGetDrive("FIXED")
For $k = 1 To $aDrives[0]
If $aDrives[$k] <> EnvGet("SystemDrive") And FileExists($aDrives[$k] & '\windows\system32\config\system') Then
$sHiveSystemGuest = $aDrives[$k] & '\windows'
ExitLoop
EndIf
Next
Case "/manual"
$sHiveSystemGuest = FileSelectFolder("Select OS  (Example: d:\Windows)", 1)
Case "/bootdrive"
If $i < $CmdLine[0] Then
$i += 1
$sNewBootDrive = StringLeft($CmdLine[$i], 2)
$sMarkerFile = StringMid($CmdLine[$i], 3)
EndIf
Case "/log"
If $i < $CmdLine[0] Then
$i += 1
$sLogFile = $CmdLine[$i]
EndIf
Case "/IgnoreLetter"
If $i < $CmdLine[0] Then
$i += 1
$sIgnoreLetter = $CmdLine[$i]
EndIf
Case "/HideLetter"
LetterClean('Removable;CDROM')
Case "/MountAll"
MountAll()
Case "/wait"
If $i < $CmdLine[0] Then
$i += 1
If Number($CmdLine[$i]) > 0 Then $sLoopWait = Number($CmdLine[$i])
EndIf
Case "/?", "/help"
MsgBox(4096, @ScriptName & $sAboot, $sHelp)
Exit
Case Else
$sHiveSystemGuest = $CmdLine[$i]
EndSwitch
$i += 1
WEnd
LogOut("Command line:" & @CRLF & $CmdLineRaw & @CRLF)
If FileExists($sHiveSystemGuest & '\system32\config\system') Then
MountGet($sReg, $aMountHost)
RunWait('reg.exe load hklm\GuestSYSTEM "' & $sHiveSystemGuest & '\system32\config\system"', '', @SW_HIDE)
For $l = 0 To 9
$sKey = "HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevice"
If $l = 0 Then
$sKey &= 's'
Else
$sKey &= $l
EndIf
MountGet($sKey, $aMountGuest)
Next
RunWait('reg.exe unload hklm\GuestSYSTEM', '', @SW_HIDE)
For $i = 1 To UBound($aMountHost, 1) - 1
$sLetterHost = $aMountHost[$i][2]
If $sLetterHost <> '' And StringInStr($sIgnoreLetter & StringLeft(EnvGet('SystemDrive'), 1), $sLetterHost) = 0 Then
$sLetterGuest = ''
For $j = 1 To UBound($aMountGuest, 1) - 1
If $aMountHost[$i][0] = $aMountGuest[$j][0] Then
$sLetterGuest = $aMountGuest[$j][2]
ExitLoop
EndIf
Next
If($sLetterHost = $sLetterGuest) Or($sLetterGuest = '') Then ContinueLoop
LogOut("Swap letter " & $sLetterHost & ': <> ' & $sLetterGuest & ':' & @CRLF)
$iLetterHostSwap = 0
For $k = 1 To UBound($aMountHost, 1) - 1
If $aMountHost[$k][2] = $sLetterGuest Then
$iLetterHostSwap = $k
ExitLoop
EndIf
Next
If $iLetterHostSwap <> 0 Then
$aMountHost[$iLetterHostSwap][2] = $sLetterHost
EndIf
$aMountHost[$i][2] = $sLetterGuest
MountSwap($sLetterHost & ':\', $sLetterGuest & ':\')
EndIf
Next
EndIf
If $sNewBootDrive <> '' Then
$aDrives = DriveGetDrive("All")
For $j = 1 To $sLoopWait
For $k = 1 To $aDrives[0]
If StringInStr("ab" & StringLeft(EnvGet('SystemDrive'), 1), StringLeft($aDrives[$k], 1)) = 0 And(DriveStatus($aDrives[$k]) <> "NOTREADY") Then
If FileExists($aDrives[$k] & $sMarkerFile) Then
$sBootDrive = $aDrives[$k]
ExitLoop 2
EndIf
EndIf
Next
Sleep($sWait)
Next
EndIf
If $sBootDrive <> '' Then
LogOut("Found BootDrive " & $sBootDrive & @CRLF)
LogOut("Swap letter " & $sBootDrive & ' <> ' & $sNewBootDrive & @CRLF)
MountSwap($sBootDrive & '\', $sNewBootDrive & '\')
EndIf
Func Conv($bStr)
If BinaryMid($bStr, 3, 4) = Binary("0x3f003f00") Then
Return BinaryToString($bStr, 2)
Else
Return String($bStr)
EndIf
EndFunc
Func FreeLetter()
Local $sFreeLetter = '', $i
For $i = Asc("c") To Asc("z")
If DriveGetType(Chr($i) & ':\') = '' Then
$sFreeLetter = Chr($i) & ":\"
ExitLoop
EndIf
Next
Return $sFreeLetter
EndFunc
Func LetterClean($sDriveType)
Local $i, $sFreeLetter, $asMount[1][3], $sTempLetter
MountGet($sReg, $asMount)
For $i = 1 To UBound($asMount) - 1
If Not $asMount[$i][2] Then
$sTempLetter = FreeLetter()
_WinAPI_SetVolumeMountPoint($sTempLetter, $asMount[$i][1])
If Not @error Then $asMount[$i][2] = StringLeft($sTempLetter, 1)
EndIf
If StringInStr($sDriveType, DriveGetType($asMount[$i][2]) & ':\') = 0 And $asMount[$i][2] And StringInStr('ab', $asMount[$i][2]) = 0 Then
Switch DriveStatus($asMount[$i][2] & ':\')
Case 'READY', 'UNKNOWN'
Case Else
If $asMount[$i][2] Then
_WinAPI_DeleteVolumeMountPoint($asMount[$i][2] & ':\')
EndIf
EndSwitch
EndIf
Next
EndFunc
Func LogOut($sStr)
If $sLogFile = '' Then Return
FileWrite($sLogFile, $sStr)
EndFunc
Func MountAll()
Local $i, $sFreeLetter, $asMount[1][3]
MountGet($sReg, $asMount)
For $i = 1 To UBound($asMount) - 1
If Not $asMount[$i][2] Then
$sTempLetter = FreeLetter()
_WinAPI_SetVolumeMountPoint($sTempLetter, $asMount[$i][1])
If Not @error Then $asMount[$i][2] = StringLeft($sTempLetter, 1)
EndIf
If DriveStatus($asMount[$i][2] & ':\') <> 'INVALID' And $asMount[$i][2] Then
If Not $asMount[$i][2] Then
$sFreeLetter = FreeLetter()
_WinAPI_SetVolumeMountPoint($sFreeLetter, $asMount[$i][1])
EndIf
EndIf
Next
EndFunc
Func MountGet($sReg, ByRef $asMount)
Local $i, $i1, $sValueName, $sLetter, $sValueData
For $i = 1 To 999
$sValueName = RegEnumVal($sReg, $i)
If @error <> 0 Then ExitLoop
$sValueData = Conv(RegRead($sReg, $sValueName))
$fFound = False
For $i1 = 1 To UBound($asMount) - 1
If $asMount[$i1][0] = $sValueData Then
$fFound = True
ExitLoop
EndIf
Next
If Not $fFound Then
$i1 = UBound($asMount)
ReDim $asMount[$i1 + 1][3]
$asMount[$i1][0] = $sValueData
EndIf
$sLetter = StringRegExpReplace($sValueName, "(?i)\\DosDevices\\([a-z]):", "\1")
If @extended <> 0 Then $asMount[$i1][2] = $sLetter
$sVolume = StringRegExpReplace($sValueName, '(?i)\\\?(\?\\Volume{[0-9a-f-]*})', "\1")
If @extended <> 0 Then $asMount[$i1][1] = '\\' & $sVolume & '\'
Next
EndFunc
Func MountSwap($sDrive1, $sDrive2)
Local $sGuid1 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive1)
Local $sGuid2 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive2)
If @error Then $sGuid2 = 0
_WinAPI_DeleteVolumeMountPoint($sDrive1)
If $sGuid2 Then
_WinAPI_DeleteVolumeMountPoint($sDrive2)
_WinAPI_SetVolumeMountPoint($sDrive1, $sGuid2)
EndIf
_WinAPI_SetVolumeMountPoint($sDrive2, $sGuid1)
EndFunc
