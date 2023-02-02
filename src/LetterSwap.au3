#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_type=a3x
#AutoIt3Wrapper_Icon=LetterSwap.ico
#AutoIt3Wrapper_Outfile=d:\__Proect\LetterSwapAu3\LetterSwap.a3x
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=LetterSwap.exe
#AutoIt3Wrapper_Res_Description=LetterSwap.exe
#AutoIt3Wrapper_Res_Fileversion=2019.2.8.18
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=2018.2.8
#AutoIt3Wrapper_Res_LegalCopyright=(c)Nikzzzz
#AutoIt3Wrapper_Run_After=%scitedir%\CheckSum\CheckSumPe.exe /c "%out%"
#AutoIt3Wrapper_Run_After=%scitedir%\CheckSum\signtool.exe sign /f "%scitedir%\CheckSum\Sert\Sert.pfx" "%out%"
#AutoIt3Wrapper_Run_After=%scitedir%\CheckSum\CheckSumPe.exe /c "%outx64%"
#AutoIt3Wrapper_Run_After=%scitedir%\CheckSum\signtool.exe sign /f "%scitedir%\CheckSum\Sert\Sert.pfx" "%outx64%"
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinAPIFiles.au3>
#include <Reg.au3>
#include <Security.au3>

Opt('MustDeclareVars', 1)
Opt('TrayIconHide', 1)
Opt('ExpandEnvStrings', 1)

Global $sAboot = "                (c)Nikzzzz 08.02.2019"
Global $sHelp = @ScriptName & " [/HideLetter|/MountAll] [/Auto|/Manual|WinDir] [/Save] [/BootDrive NewLetter:[\TagFile]] [/SetLetter NewLetter:\TagFile] [/RestartExplorer] [/log [LogFile|con:]] [/IgnoreLetter Letters] [/Swap Drive: Drive:] [/wait 10]" & @CRLF


If $CmdLine[0] = 0 Then
	MsgBox(4096, @ScriptName & $sAboot, $sHelp)
	Exit
EndIf
Global $sHostKey = "HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices"
Global $sGuestKey = "HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevices"
Global $aMountHost[1][2], $aMountGuest[1][2], $sIgnoreLetter = 'yz', $sLogFile = '', $sSystemGuest = '', $sBootDrive = '', $sGuestKey, $sTagFile = '', $sTagFile1 = '',$iLetterClean=0,$iMountAll=1
Global $sNewBootDrive = '', $s = "", $i = 1, $iWait = 100, $fSave = False, $sGuest = '', $sRestartExplorer = False, $sHostDrive = StringLeft(EnvGet('SourceDrive'), 1)
Local $vTemp, $aDrives, $sLetterGuest, $sLetterHost, $sNewDrive1 = ''
While $i <= $CmdLine[0]
	$vTemp = StringLower($CmdLine[$i])
	Switch $vTemp
		Case "/MountAll"
			$iMountAll=1
		Case "/auto"
			$aDrives = DriveGetDrive("FIXED")
			For $k = 1 To $aDrives[0]
				If $aDrives[$k] <> EnvGet("SystemDrive") And FileExists($aDrives[$k] & '\windows\system32\config\system') Then
					$sSystemGuest = $aDrives[$k] & '\windows'
					ExitLoop
				EndIf
			Next
		Case "/manual"
			$sSystemGuest = FileSelectFolder("Select OS  (Example: d:\Windows)", 1)
		Case "/bootdrive"
			If $i < $CmdLine[0] Then
				$i += 1
				$sNewBootDrive = StringLeft($CmdLine[$i], 2)
				$sTagFile = StringMid($CmdLine[$i], 4)
			EndIf
		Case "/setletter"
			If $i < $CmdLine[0] Then
				$i += 1
				$sNewDrive1 = StringLeft($CmdLine[$i], 2)
				$sTagFile1 = StringMid($CmdLine[$i], 4)
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
				_MountSwap($CmdLine[$i + 1], $CmdLine[$i + 2])
				$i += 2
			EndIf
		Case "/HideLetter"
			$iLetterClean=1
		Case "/Save"
			$fSave = True
		Case "/wait"
			If $i < $CmdLine[0] Then
				$i += 1
				If Number($CmdLine[$i]) > 0 Then $iWait = Number($CmdLine[$i]) * 10
			EndIf
		Case "/?", "/help"
			MsgBox(4096, @ScriptName & $sAboot, $sHelp)
			Exit
		Case "/RestartExplorer"
			$sRestartExplorer = True
		Case "/IgnoreCD"
			$aDrives = DriveGetDrive("CDROM")
			For $k = 1 To $aDrives[0]
				$sIgnoreLetter &= StringLeft($aDrives[$k], 1)
			Next
		Case Else
			$sSystemGuest = $CmdLine[$i]
	EndSwitch
	$i += 1
WEnd
_LogOutN("----- Start " & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & "  Command Line: " & @ScriptName & $CmdLineRaw & @CRLF)
$sIgnoreLetter &= StringLeft(EnvGet('SystemDrive'), 1)
If $iMountAll Then _MountAll()
if $iLetterClean Then _LetterClean('Removable;CDROM')
If $sSystemGuest <> '' And FileExists($sSystemGuest & '\system32\config\system') Then
	_LogOutN('Hosts System : ' & EnvGet('SystemRoot'))
	_LogOutN('Guest System : ' & $sSystemGuest & @CRLF)
	$sGuest = StringLeft($sSystemGuest, 2)
	_MountGet($sHostKey, $aMountHost)
	_MountPrint('...... Host:  ' & $sHostKey, $aMountHost)
	_RegLoadHive($sSystemGuest & '\system32\config\system', 'hklm\GuestSYSTEM')
	_MountGet($sGuestKey, $aMountGuest)
	_RegUnLoadHive('hklm\GuestSYSTEM')
	_MountPrint('...... Guest:  ' & $sHostKey, $aMountGuest)
	For $i = 1 To UBound($aMountGuest, 1) - 1
		_MountGet($sHostKey, $aMountHost)
		$sLetterGuest = $aMountGuest[$i][1]
		If StringInStr($sIgnoreLetter, $sLetterGuest) Then ContinueLoop
		$sLetterHost = ''
		For $i1 = 1 To UBound($aMountHost, 1) - 1
			If $aMountHost[$i1][0] = $aMountGuest[$i][0] Then
				$sLetterHost = $aMountHost[$i1][1]
				ExitLoop
			EndIf
		Next
		If $sLetterHost = '' Then ContinueLoop
		If StringInStr($sIgnoreLetter, $sLetterHost) Then ContinueLoop
		_MountSwap($sLetterHost, $sLetterGuest)
	Next

	If $fSave Then
		RegWrite('HKLM\SOFTWARE\LetterSwap', 'Guest', 'REG_SZ', StringLeft($sSystemGuest, 2))
		RegWrite('HKLM\SOFTWARE\LetterSwap', 'SourceDrive', 'REG_SZ', $sHostDrive & ':')
	EndIf

EndIf

If $sNewDrive1 <> '' And $sTagFile1 <> '' Then
	While $iWait > 0
		$aDrives = DriveGetDrive('all')
		For $i = 1 To UBound($aDrives) - 1
			If Not FileExists($aDrives[$i] & '\' & $sTagFile1) Then ContinueLoop
			_LogOutN('Found TagFile : "' & $aDrives[$i] & '\' & $sTagFile1 & '"')
			_MountSwap($aDrives[$i], $sNewDrive1)
			ExitLoop -2
		Next
		$iWait -= 1
		Sleep(100)
	WEnd
EndIf

If $sNewBootDrive <> '' Then
	While $iWait > 0
		$sBootDrive = _GetBootDrive($sTagFile)
		If $sBootDrive <> '' Then
			_LogOutN('Found BootDrive : "' & $sBootDrive & '"')
			_MountSwap($sBootDrive, $sNewBootDrive)
			ExitLoop
		EndIf
		$iWait -= 1
		Sleep(100)
	WEnd
EndIf

If $sRestartExplorer Then
	_LogOutN('Restart Explorer')
	While ProcessExists("Explorer.exe")
		ProcessClose("Explorer.exe")
		Sleep(500)
	WEnd
	Run("Explorer.exe")
EndIf

_LogOutN('-----------------------------------------------------------------------------------------------------')
_MountGet($sHostKey, $aMountHost)
_MountPrint('...... Host:  ' & $sHostKey, $aMountHost)
_LogOutN('-----------------------------------------------------------------------------------------------------')
_LogOutN("----- Finish  " & @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF)

Exit

Func _GetBootDrive($sTagFile)
	Local $vDriveList, $i, $i1
	Local $sStartOpt = RegRead('HKLM\SYSTEM\CurrentControlSet\Control', 'SystemStartOptions')
	Local $vTmp = StringRegExp($sStartOpt, '(?i)MININT\s.*RDPATH=(?:.*\)(\w+)\((\d+)\))?(\\.*)', 2)
	If @error Then Return ''
	If $vTmp[3] = '' Then Return ''
	Switch $vTmp[1]
		Case 'CDROM'
			$vDriveList = 'CDROM'
		Case 'PARTITION'
			$vDriveList = 'REMOVABLE,FIXED,NETWORK'
		Case Else
			$vDriveList = 'REMOVABLE,FIXED,NETWORK,CDROM'
	EndSwitch
	$vDriveList = StringSplit($vDriveList, ',', 2)
	For $i = 0 To UBound($vDriveList) - 1
		Local $asDriveLetter = DriveGetDrive($vDriveList[$i])
		For $i1 = 1 To UBound($asDriveLetter) - 1
			If StringInStr('a:b:' & EnvGet('SystemDrive'), $asDriveLetter[$i1]) > 0 Then ContinueLoop
			If $vTmp[1] = 'PARTITION' Then
				If _GetPart($asDriveLetter[$i1]) <> $vTmp[2] Then ContinueLoop
			EndIf
			If $sTagFile <> '' Then
				If Not FileExists($asDriveLetter[$i1] & '\' & $sTagFile) Then ContinueLoop
			EndIf
			If FileExists($asDriveLetter[$i1] & $vTmp[3]) Then Return $asDriveLetter[$i1]
		Next
	Next
	Return ''
EndFunc   ;==>_GetBootDrive

Func _GetPart($sDrive)
	Local $aDriveNumber = _WinAPI_GetDriveNumber($sDrive)
	If IsArray($aDriveNumber) Then Return $aDriveNumber[2]
	Return ''
EndFunc   ;==>_GetPart

Func _Conv($bStr)
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
EndFunc   ;==>_Conv

Func _Reverse($sStr)
	Local $sRet = '', $i
	While StringLen($sStr)
		$sRet &= StringRight($sStr, 2)
		$sStr = StringTrimRight($sStr, 2)
	WEnd
	Return $sRet
EndFunc   ;==>_Reverse

Func _FreeLetter()
	Local $sFreeLetter = '', $i
	For $i = Asc("c") To Asc("z")
		If DriveGetType(Chr($i) & ':\') = '' Then
			$sFreeLetter = Chr($i) & ":\"
			ExitLoop
		EndIf
	Next
	Return $sFreeLetter
EndFunc   ;==>_FreeLetter

Func _LetterClean($sdrivetype)
  _MountAll()
	Local $i, $sFreeLetter, $asMount[1][3], $aRet[1]
	_MountGet($sHostKey, $asMount)
	For $i = 1 To UBound($asMount) - 1
		If StringInStr($sdrivetype, DriveGetType($asMount[$i][1])) And (StringInStr("ab", $asMount[$i][1]) = 0) Then
			$aRet = DllCall('kernel32.dll', 'bool', 'GetVolumeInformationW', 'wstr', $asMount[$i][1], 'wstr', '', 'dword', 4096, 'dword*', 0, 'dword*', 0, 'dword*', 0, 'wstr', '', 'dword', 4096)
			If @error Or Not $aRet[0] Then
				If $asMount[$i][1] Then _WinAPI_DeleteVolumeMountPoint($asMount[$i][1] & ":\")
        _LogOutN('UnMount ' & $asMount[$i][1])
			EndIf
		EndIf
	Next
EndFunc   ;==>_LetterClean

Func _LogOut($sStr = '')
	Switch $sLogFile
		Case ''
		Case 'con:'
			ConsoleWrite($sStr)
		Case Else
			FileWrite($sLogFile, $sStr)
	EndSwitch
EndFunc   ;==>_LogOut

Func _LogOutN($sStr = '')
	_LogOut($sStr & @CRLF)
EndFunc   ;==>_LogOutN

Func _MountPrint($sStr, ByRef $asMount)
	Local $i
	_LogOutN($sStr)
	For $i = 1 To UBound($asMount) - 1
		_LogOutN('"' & $asMount[$i][1] & ':"  "' & $asMount[$i][0] & '"')
	Next
	_LogOutN()
EndFunc   ;==>_MountPrint

Func _MountAll()
	Local $i, $sFreeLetter, $sGuids = ',', $sGuid
	Local $aDrives = DriveGetDrive('all')
	For $i = 1 To UBound($aDrives) - 1
		$sGuids &= _WinAPI_GetVolumeNameForVolumeMountPoint($aDrives[$i] & '\') & ','
	Next
	For $i = 1 To 999
		$sGuid = '\\?\Volume' & RegEnumKey('HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\CPC\Volume', $i) & '\'
		If @error Then ExitLoop
		If StringInStr($sGuids, $sGuid) > 0 Then ContinueLoop
		$sFreeLetter = _FreeLetter()
		_WinAPI_SetVolumeMountPoint($sFreeLetter, $sGuid)
		$sGuids &= $sGuid & ','
		_LogOutN('Mount ' & $sFreeLetter & ' ' & $sGuid)
	Next
EndFunc   ;==>_MountAll

Func _MountGet($sHostKey, ByRef $asMount)
	Local $i, $i1 = 0, $sValueName, $sLetter, $sValueData, $fFound, $sVolume
	ReDim $asMount[1][2]
	For $i = 1 To 999
		$sValueName = RegEnumVal($sHostKey, $i)
		If @error <> 0 Then ExitLoop
		$sValueData = RegRead($sHostKey, $sValueName)
		If @error Then ContinueLoop
		$sValueData = _Conv($sValueData)
		$sLetter = StringRegExpReplace($sValueName, "(?i)\\DosDevices\\([a-z]):", "\1")
		If @extended <> 0 Then
			$i1 = UBound($asMount)
			ReDim $asMount[$i1 + 1][2]
			$asMount[$i1][0] = $sValueData
			$asMount[$i1][1] = $sLetter
		EndIf
	Next
EndFunc   ;==>_MountGet

Func _MountSwap($sDrive1, $sDrive2)
	$sDrive1 = StringLeft($sDrive1, 1) & ':\'
	$sDrive2 = StringLeft($sDrive2, 1) & ':\'
	If $sDrive1 = $sDrive2 Then Return 1
	Local $sGuid1 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive1)
	Local $sGuid2 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive2)
	While 1
		If $sGuid1 Then
			_WinAPI_DeleteVolumeMountPoint($sDrive1)
			If @error Then
				$sGuid1 = ''
				$sGuid2 = ''
				ExitLoop
			EndIf
		EndIf
		If $sGuid2 Then
			_WinAPI_DeleteVolumeMountPoint($sDrive2)
			If @error Then
				$sGuid2 = ''
				ExitLoop
			EndIf
		EndIf
		If $sGuid1 Then
			_WinAPI_SetVolumeMountPoint($sDrive2, $sGuid1)
			If @error Then
				ExitLoop
			EndIf
		EndIf
		If $sGuid2 Then
			_WinAPI_SetVolumeMountPoint($sDrive1, $sGuid2)
			If @error Then
				_WinAPI_DeleteVolumeMountPoint($sDrive2)
				ExitLoop
			EndIf
		EndIf
		_LogOutN('Swap letter "' & StringLeft($sDrive1, 2) & '" <> "' & StringLeft($sDrive2, 2) & '"')
		Return 1
	WEnd
	If $sGuid1 Then _WinAPI_SetVolumeMountPoint($sDrive1, $sGuid1)
	If $sGuid1 Then _WinAPI_SetVolumeMountPoint($sDrive2, $sGuid2)
	_LogOutN('Swap letter "' & StringLeft($sDrive1, 2) & '" <> "' & StringLeft($sDrive2, 2) & '" - Error')
	Return 0
EndFunc   ;==>_MountSwap

