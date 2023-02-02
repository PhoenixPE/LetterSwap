#NoTrayIcon
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" & ((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func __WINVER()
	Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
	DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
	Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
	If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
	Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc   ;==>__WINVER
Func _WinAPI_DeleteVolumeMountPoint($sMountedPath)
	Local $aRet = DllCall('kernel32.dll', 'bool', 'DeleteVolumeMountPointW', 'wstr', $sMountedPath)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aRet[0]
EndFunc   ;==>_WinAPI_DeleteVolumeMountPoint
Func _WinAPI_GetVolumeNameForVolumeMountPoint($sMountedPath)
	Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVolumeNameForVolumeMountPointW', 'wstr', $sMountedPath, 'wstr', '', 'dword', 80)
	If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, '')
	Return $aRet[2]
EndFunc   ;==>_WinAPI_GetVolumeNameForVolumeMountPoint
Func _WinAPI_SetVolumeMountPoint($sFilePath, $sGUID)
	Local $aRet = DllCall('kernel32.dll', 'bool', 'SetVolumeMountPointW', 'wstr', $sFilePath, 'wstr', $sGUID)
	If @error Then Return SetError(@error, @extended, False)
	Return $aRet[0]
EndFunc   ;==>_WinAPI_SetVolumeMountPoint
Opt('MustDeclareVars', 1)
Opt('TrayIconHide', 1)
Opt('ExpandEnvStrings', 1)
Global $sAboot = "                (c)Nikzzzz 27.10.2017"
Global $sHelp = @ScriptName & " [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/Save] [/BootDrive NewLetter:\TagFile] [/RestartExplorer] [/log LogFile] [/IgnoreLetter Letters] [/Swap Drive: Drive:] [/wait 10]" & @CRLF
Global $sReg = "HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices"
If $CmdLine[0] = 0 Then
	MsgBox(4096, @ScriptName & $sAboot, $sHelp)
	Exit
EndIf
Global $aMountHost[1][2], $aMountGuest[1][2], $sIgnoreLetter = 'yz', $sLogFile = '', $sHiveSystemGuest = '', $sBootDrive = '', $sKey, $sMarkerFile
Global $sNewBootDrive = '', $s = "", $i = 1, $sWait = 100, $fSave = False, $sGuest = '', $sRestartExplorer = False, $sSourceDrive = StringLeft(EnvGet('SourceDrive'), 1)
Local $vTemp, $aDrives, $sLetterGuest, $sLetterHost
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
		Case "/Swap"
			If ($i + 2) <= $CmdLine[0] Then
				MountSwap(StringLeft($CmdLine[$i + 1], 1) & ':\', StringLeft($CmdLine[$i + 2], 1) & ':\')
				$i += 2
			EndIf
		Case "/HideLetter"
			LetterClean('Removable;CDROM')
		Case "/MountAll"
			MountAll()
		Case "/Save"
			$fSave = True
		Case "/wait"
			If $i < $CmdLine[0] Then
				$i += 1
				If Number($CmdLine[$i]) > 0 Then $sWait = Number($CmdLine[$i]) * 10
			EndIf
		Case "/?", "/help"
			MsgBox(4096, @ScriptName & $sAboot, $sHelp)
			Exit
		Case "/RestartExplorer"
			$sRestartExplorer = True
		Case "/IgnoreCD"
			$aDrives = DriveGetDrive("CDROM")
			For $k = 1 To $aDrives[0]
				$sIgnoreLetter &= $aDrives[$k]
			Next
		Case Else
			$sHiveSystemGuest = $CmdLine[$i]
	EndSwitch
	$i += 1
WEnd
LogOut("----- Start " & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & "  Command Line: " & @ScriptName & $CmdLineRaw & @CRLF & @CRLF)
$sIgnoreLetter &= StringLeft(EnvGet('SystemDrive'), 1)
If FileExists($sHiveSystemGuest & '\system32\config\system') Then
	$sGuest = StringLeft($sHiveSystemGuest, 2)
	MountGet($sReg, $aMountHost)
	Test('...... Host:  ' & $sReg, $aMountHost)
	RunWait('reg.exe load hklm\GuestSYSTEM "' & $sHiveSystemGuest & '\system32\config\system"', '', @SW_HIDE)
	$sKey = "HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevices"
	MountGet($sKey, $aMountGuest)
	RunWait('reg.exe unload hklm\GuestSYSTEM', '', @SW_HIDE)
	Test('...... Guest:  ' & $sReg, $aMountGuest)
	For $i = 1 To UBound($aMountGuest, 1) - 1
		MountGet($sReg, $aMountHost)
		$sLetterGuest = $aMountGuest[$i][2]
		If StringInStr($sIgnoreLetter, $sLetterGuest) Then ContinueLoop
		$sLetterHost = ''
		For $j = 1 To UBound($aMountHost, 1) - 1
			If $aMountHost[$j][0] = $aMountGuest[$i][0] Then
				$sLetterHost = $aMountHost[$j][2]
				ExitLoop
			EndIf
		Next
		If ($sLetterHost = $sLetterGuest) Or ($sLetterHost = '') Then ContinueLoop
		MountSwap($sLetterHost & ':\', $sLetterGuest & ':\')
		LogOut("Swap letter " & $sLetterHost & ': <> ' & $sLetterGuest & ':' & @CRLF)
	Next
	If $fSave Then
		RegWrite('HKLM\SOFTWARE\LetterSwap', 'Guest', 'REG_SZ', $sGuest)
		RegWrite('HKLM\SOFTWARE\LetterSwap', 'SourceDrive', 'REG_SZ', $sSourceDrive & ':')
	EndIf
EndIf
If $sNewBootDrive <> '' Then
	For $j = 1 To $sWait
		$aDrives = DriveGetDrive("All")
		For $k = 1 To $aDrives[0]
			If StringInStr("ab" & StringLeft(EnvGet('SystemDrive'), 1), StringLeft($aDrives[$k], 1)) = 0 And (DriveStatus($aDrives[$k]) <> "NOTREADY") Then
				If FileExists($aDrives[$k] & $sMarkerFile) Then
					$sBootDrive = $aDrives[$k]
					ExitLoop 2
				EndIf
			EndIf
		Next
		Sleep(100)
	Next
EndIf
If $sBootDrive <> '' Then
	LogOut("Found BootDrive " & $sBootDrive & @CRLF)
	MountSwap($sBootDrive & '\', $sNewBootDrive & '\')
	LogOut("Swap letter " & $sBootDrive & ' <> ' & $sNewBootDrive & @CRLF)
EndIf
If $sRestartExplorer Then
	While ProcessExists("Explorer.exe")
		ProcessClose("Explorer.exe")
		Sleep(500)
	WEnd
	Run("Explorer.exe")
EndIf
LogOut(@CRLF)
LogOut("----- Finish  " & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & @CRLF)
MountGet($sReg, $aMountHost)
Test('...... Host:  ' & $sReg, $aMountHost)
LogOut(@CRLF)
Exit
Func Conv($bStr)
	Local $sStr = '', $sRet = '', $sTemp
	If BinaryMid($bStr, 3, 4) = StringToBinary("??", 2) Then
		$sStr = BinaryToString($bStr, 2)
		$sTemp = StringRegExpReplace($sStr, '(?i)\\\?\?\\STORAGE#Partition#S(........).*', '\1')
		If @extended Then
			$sRet = _Reverse($sTemp)
			$sTemp = StringRegExpReplace($sStr, '(?i)\\\?\?\\STORAGE#Partition#S........_O(.*)_.*', '\1')
			If Mod(StringLen($sTemp), 2) Then $sTemp = '0' & $sTemp
			$sRet &= _Reverse($sTemp)
			While StringLen($sRet) < 24
				$sRet &= '00'
			WEnd
			$sRet = '0x' & StringUpper($sRet)
		Else
			$sRet = $sStr
		EndIf
	Else
		$sRet = String($bStr)
	EndIf
	Return $sRet
EndFunc   ;==>Conv
Func _Reverse($sStr)
	Local $sRet = '', $i
	While StringLen($sStr)
		$sRet &= StringRight($sStr, 2)
		$sStr = StringTrimRight($sStr, 2)
	WEnd
	Return $sRet
EndFunc   ;==>_Reverse
Func FreeLetter()
	Local $sFreeLetter = '', $i
	For $i = Asc("c") To Asc("z")
		If DriveGetType(Chr($i) & ':\') = '' Then
			$sFreeLetter = Chr($i) & ":\"
			ExitLoop
		EndIf
	Next
	Return $sFreeLetter
EndFunc   ;==>FreeLetter
Func LetterClean($sdrivetype)
	Local $i, $sFreeLetter, $asMount[1][3], $aRet[1]
	MountGet($sReg, $asMount)
	For $i = 1 To UBound($asMount) - 1
		If StringInStr($sdrivetype, DriveGetType($asMount[$i][1])) And (StringInStr("ab", $asMount[$i][2]) = 0) Then
			$aRet = DllCall('kernel32.dll', 'bool', 'GetVolumeInformationW', 'wstr', $asMount[$i][1], 'wstr', '', 'dword', 4096, 'dword*', 0, 'dword*', 0, 'dword*', 0, 'wstr', '', 'dword', 4096)
			If @error Or Not $aRet[0] Then
				If $asMount[$i][2] Then _WinAPI_DeleteVolumeMountPoint($asMount[$i][2] & ":\")
			Else
				If Not $asMount[$i][2] Then
					$sFreeLetter = FreeLetter()
					If $sFreeLetter = "" Then Exit 1
					_WinAPI_SetVolumeMountPoint($sFreeLetter, $asMount[$i][1])
				EndIf
			EndIf
		EndIf
	Next
EndFunc   ;==>LetterClean
Func LogOut($sStr)
	If $sLogFile = '' Then Return
	FileWrite($sLogFile, $sStr)
EndFunc   ;==>LogOut
Func Test($sStr, ByRef $asMount)
	Local $i
	LogOut($sStr & @CRLF)
	For $i = 1 To UBound($asMount) - 1
		LogOut('"' & $asMount[$i][2] & '"  "' & $asMount[$i][1] & '" "' & $asMount[$i][0] & '"' & @CRLF)
	Next
	LogOut(@CRLF)
EndFunc   ;==>Test
Func MountAll()
	Local $i, $sFreeLetter, $asMount[1][3], $aRet[1]
	MountGet($sReg, $asMount)
	For $i = 1 To UBound($asMount) - 1
		$aRet = DllCall('kernel32.dll', 'bool', 'GetVolumeInformationW', 'wstr', $asMount[$i][1], 'wstr', '', 'dword', 4096, 'dword*', 0, 'dword*', 0, 'dword*', 0, 'wstr', '', 'dword', 4096)
		If Not @error Then
			If Not $asMount[$i][2] Then
				$sFreeLetter = FreeLetter()
				If $sFreeLetter = "" Then Exit 1
				_WinAPI_SetVolumeMountPoint($sFreeLetter, $asMount[$i][1])
			EndIf
		EndIf
	Next
EndFunc   ;==>MountAll
Func MountGet($sReg, ByRef $asMount)
	Local $i, $i1, $sValueName, $sLetter, $sValueData, $fFound, $sVolume
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
	$i = 1
	While $i < UBound($asMount)
		If $asMount[$i][2] = '' Then
			For $i1 = $i + 1 To UBound($asMount) - 1
				$asMount[$i1 - 1][0] = $asMount[$i1][0]
				$asMount[$i1 - 1][1] = $asMount[$i1][1]
				$asMount[$i1 - 1][2] = $asMount[$i1][2]
			Next
			ReDim $asMount[UBound($asMount) - 1][3]
		Else
			$i += 1
		EndIf
	WEnd
EndFunc   ;==>MountGet
Func MountSwap($sDrive1, $sDrive2)
	SetError(0)
	If $sDrive1 = $sDrive2 Then Return
	Local $sGuid1 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive1)
	Local $sGuid2 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive2)
	If @error Then $sGuid2 = 0
	_WinAPI_DeleteVolumeMountPoint($sDrive1)
	If $sGuid2 Then
		If Not _WinAPI_DeleteVolumeMountPoint($sDrive2) Then
			SetError(1)
			Return
		EndIf
		_WinAPI_SetVolumeMountPoint($sDrive1, $sGuid2)
	EndIf
	_WinAPI_SetVolumeMountPoint($sDrive2, $sGuid1)
	SetError(0)
EndFunc   ;==>MountSwap
