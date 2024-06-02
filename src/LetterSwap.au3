#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=LetterSwap.ico
#AutoIt3Wrapper_Outfile=..\bin\x86\LetterSwap.exe
#AutoIt3Wrapper_Outfile_x64=..\bin\x64\LetterSwap.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Comment=LetterSwap.exe
#AutoIt3Wrapper_Res_Description=LetterSwap.exe
#AutoIt3Wrapper_Res_Fileversion=2024.5.27.79
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=2024.5.27
#AutoIt3Wrapper_Res_LegalCopyright=(c) Nikzzzz, Homes32 & Contributors
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; AutoIt 3.3.16.1

#include <WinAPIFiles.au3>
#include ".\Reg.au3"
#include ".\SecurityEx.au3"

Opt('MustDeclareVars', 1)
Opt('TrayIconHide', 1)
Opt('ExpandEnvStrings', 1)

Global $sAbout = "LetterSwap v" & FileGetVersion(@ScriptFullPath) & " (c) Nikzzzz, Homes32 & Contributors"
Global $sUsage = "USAGE: " & @ScriptName & " [/Swap <DriveLetter1>: <DriveLetter2>:] [/HideLetter|/MountAll] [/Auto|/Manual|/WinDir <Path>] [/BootDrive <NewLetter>:[\<TagFile>]] [/SetLetter <NewLetter>:\TagFile] [/Wait <Seconds>] [/Save] [/RestartExplorer] [/IgnoreLetter <Letters>] [/IgnoreCD] [/Log <LogFile>|con:]" & @CRLF
Global $sHelp = @CRLF & "Swap drive letters and/or synchronize letters of disks on based on the registry of the guest OS." & @CRLF _
		& "" & @CRLF _
		& $sUsage & @CRLF _
		& "  /HideLetter                          Hides inactive removable and CDROM disks." & @CRLF _
		& "  /MountAll                            Mount inactive disks to the first available drive letter." & @CRLF _
		& "  /Swap <DriveLetter1> <DriveLetter2>  Swap the specified drive letters." & @CRLF _
		& "                                         Ex. LetterSwap.exe /Swap D: E:" & @CRLF _
		& "  /Auto                                Find the first guest OS." & @CRLF _
		& "  /Manual                              Display a dialog prompting to select the guest OS Windows directory." & @CRLF _
		& "  /WinDir <Path>                       Specify the directory of the guest OS. (Ex. D:\Windows)" & @CRLF _
		& "  /BootDrive <NewLetter>:              Assigns the boot disk the specified drive letter." & @CRLF _
		& "  /BootDrive <NewLetter>:\<TagFile>    Search for <TagFile> and assign the boot disk the specified drive letter." & @CRLF _
		& "                                         Ex. Letterswap.exe /Auto /BootDrive Y:\USB.Y" & @CRLF _
		& "  /SetLetter <NewLetter>:\<TagFile>    Search for <TagFile> and assign the disk the specified drive letter." & @CRLF _
		& "                                         Ex. Letterswap.exe /SetLetter Z:\File.tag" & @CRLF _
		& "  /Wait <Seconds>                      Used in conjunction with /BootDrive or /SetLetter to define the number of" & @CRLF _
        & "                                       seconds to wait for drives to become available." & @CRLF _
		& "  /Save                                When used in conjunction with /Auto, /Manual, or /WinDir save the Guest" & @CRLF _
		& "                                       and Source drive letters to the registry (HKLM\SOFTWARE\LetterSwap)." & @CRLF _
		& "  /RestartExplorer                     Restart Explorer.exe after letter change." & @CRLF _
		& "  /IgnoreLetter <Letter>[...]          When used in conjunction with /Auto, /Manual, or /WinDir ignore the specified" & @CRLF _
		& "                                       drive letters. (The system drive and letters y,z are always ignored.)" & @CRLF _
		& "                                         Ex. Letterswap.exe /IgnoreLetter abde" & @CRLF _
		& "  /IgnoreCD                            When used in conjunction with /Auto, /Manual, or /WinDir ignore all drive" & @CRLF _
		& "                                       letters belonging to CDROM drives." & @CRLF _
		& "  /Log <LogFile>|con:                  Output to the specified log file or console." & @CRLF _
		& " " & @CRLF

If $CmdLine[0] = 0 Then
	ConsoleWrite(@CRLF & $sAbout & @CRLF & @CRLF & $sUsage & @CRLF)
	Exit
EndIf

Global $sHostKey = "HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices"
Global $sGuestKey = "HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevices"
Global $aMountHost[1][2], $aMountGuest[1][2], $sIgnoreLetter = '', $sLogFile = '', $sSystemGuest = '', $sBootDrive = '', $sGuestKey, $sTagFile = '', $sTagFile1 = '', $iLetterClean = 0, $iMountAll = 0
Global $sNewBootDrive = '', $s = "", $i = 1, $iWait0 = 100, $iWait, $fSave = False, $sGuest = '', $sRestartExplorer = False, $sHostDrive = StringLeft(EnvGet('SourceDrive'), 1)
Local $aDrives, $sLetterGuest, $sLetterHost, $sNewDrive1 = ''

; Process Cmdline
While $i <= $CmdLine[0]
	Switch $CmdLine[$i]
		Case "/?", "/h"
			ConsoleWrite(@CRLF & $sAbout & @CRLF & $sHelp & @CRLF)
			Exit
		Case "/HideLetter"
			$iLetterClean = 1
		Case "/MountAll"
			$iMountAll = 1
		Case "/Auto"
			$aDrives = DriveGetDrive("FIXED")
			For $k = 1 To $aDrives[0]
				If $aDrives[$k] <> EnvGet("SystemDrive") And FileExists($aDrives[$k] & '\windows\system32\config\system') Then
					$sSystemGuest = $aDrives[$k] & '\windows'
					ExitLoop
				EndIf
			Next
		Case "/Manual"
			$sSystemGuest = FileSelectFolder("Select the OS directory (Example: d:\Windows)", 1)
		Case "/WinDir"
			$i += 1
			$sSystemGuest = $CmdLine[$i]
		Case "/BootDrive"
			If $i < $CmdLine[0] Then
				$i += 1
				$sNewBootDrive = StringLeft($CmdLine[$i], 2)
				$sTagFile = StringMid($CmdLine[$i], 4)
			EndIf
		Case "/SetLetter"
			If $i < $CmdLine[0] Then
				$i += 1
				$sNewDrive1 = StringLeft($CmdLine[$i], 2)
				$sTagFile1 = StringMid($CmdLine[$i], 4)
			EndIf
		Case "/Swap"
			If ($i + 2) <= $CmdLine[0] Then
				_MountSwap($CmdLine[$i + 1], $CmdLine[$i + 2])
				$i += 2
			EndIf
		Case "/IgnoreLetter", "/IgnoreLetters"
			If $i < $CmdLine[0] Then
				$i += 1
				$sIgnoreLetter = StringUpper($CmdLine[$i])
			EndIf
		Case "/IgnoreCD"
			$aDrives = DriveGetDrive("CDROM")
			For $k = 1 To $aDrives[0]
				$sIgnoreLetter &= StringUpper(StringLeft($aDrives[$k], 1))
			Next
		Case "/Log"
			If $i < $CmdLine[0] Then
				$i += 1
				$sLogFile = $CmdLine[$i]
			EndIf
		Case "/RestartExplorer"
			$sRestartExplorer = True
		Case "/Save"
			$fSave = True
		Case "/Wait"
			If $i < $CmdLine[0] Then
				$i += 1
				If Number($CmdLine[$i]) >= 0 Then $iWait0 = Number($CmdLine[$i]) * 10
			EndIf
		Case Else
			If $sLogFile = "" Then
				ConsoleWrite(@CRLF & $sUsage & @CRLF)
				ConsoleWrite("Command Line: " & @ScriptName & " " & $CmdLineRaw & @CRLF & "Invalid Argument: " & $CmdLine[$i] & @CRLF)
			Else
				_LogOutN("Command Line: " & @ScriptName & " " & $CmdLineRaw & @CRLF & "Invalid Argument: " & $CmdLine[$i])
			EndIf
			Exit -1
	EndSwitch
	$i += 1
WEnd

_LogOutN("LetterSwap v" & FileGetVersion(@ScriptFullPath) & " Started " & @MDAY & "-" & @MON & "-" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC)
_LogOutN("Command Line: " & @ScriptName & " " & $CmdLineRaw & @CRLF)

; Ignore the SystemDrive
$sIgnoreLetter &= StringLeft(EnvGet('SystemDrive'), 1)

; /MountAll
If $iMountAll Then _MountAll()

; /HideLetter
If $iLetterClean Then _LetterClean('Removable;CDROM')

; /Auto /Manual /WinDir
If $sSystemGuest <> '' And FileExists($sSystemGuest & '\system32\config\system') Then
	_LogOutN('Host  System : ' & EnvGet('SystemRoot'))
	_LogOutN('Guest System : ' & $sSystemGuest & @CRLF)
	$sGuest = StringLeft($sSystemGuest, 2)
	_LogOutN("Host Volume Information:")
	_MountGet($sHostKey, $aMountHost)
	_MountPrint('...... Host:  ' & $sHostKey, $aMountHost)
	_RegLoadHive($sSystemGuest & '\system32\config\system', 'HKLM\GuestSYSTEM')
	_MountGet($sGuestKey, $aMountGuest)
	_LogOutN("Guest Volume Information:")
	_RegUnLoadHive('HKLM\GuestSYSTEM')
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

	; /Save
	If $fSave Then
		RegWrite('HKLM\SOFTWARE\LetterSwap', 'Guest', 'REG_SZ', StringLeft($sSystemGuest, 2))
		RegWrite('HKLM\SOFTWARE\LetterSwap', 'SourceDrive', 'REG_SZ', $sHostDrive & ':')
	EndIf
EndIf

; /SetLetter
If $sNewDrive1 <> '' And $sTagFile1 <> '' Then
	$iWait = $iWait0
	While $iWait >= 0
		$aDrives = DriveGetDrive('all')
		For $i = 1 To UBound($aDrives) - 1
			If Not FileExists($aDrives[$i] & '\' & $sTagFile1) Then ContinueLoop
			_LogOutN('Found TagFile : "' & $aDrives[$i] & '\' & $sTagFile1 & '"')
			_MountSwap($aDrives[$i], $sNewDrive1)
			ExitLoop 2
		Next
		$iWait -= 1
		Sleep(100)
	WEnd
EndIf

; /Bootdrive
If $sNewBootDrive <> '' Then
	$iWait = $iWait0
	While $iWait >= 0
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

; /RestartExplorer
If $sRestartExplorer Then
	_LogOutN('Restart Explorer')
	While ProcessExists("Explorer.exe")
		ProcessClose("Explorer.exe")
		Sleep(500)
	WEnd
	Run("Explorer.exe")
EndIf

; Display current Mount Points now that we are finished processing
_MountGet($sHostKey, $aMountHost)
_LogOutN(@CRLF & "Current Host Volume Information:")
_MountPrint('...... Host:  ' & $sHostKey, $aMountHost)

_LogOutN("LetterSwap Finished " & @MDAY & "-" & @MON & "-" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF)
Exit 0 ; Done!

Func _GetBootDrive($sTagFile)
	Local $vDriveList, $i, $i1, $hF, $bData
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
			If FileExists($asDriveLetter[$i1] & $vTmp[3]) Then
				If FileExists(EnvGet('SystemDrive') & '\$WIMDESC') Then
					$hF = FileOpen($asDriveLetter[$i1] & $vTmp[3], 16)
					FileSetPos($hF, FileGetSize($asDriveLetter[$i1] & $vTmp[3]) - 32768, 0)
					$bData = FileRead($hF)
					FileClose($hF)
					If StringInStr(StringMid($bData, 3), StringMid(_FileRead(EnvGet('SystemDrive') & '\$WIMDESC', 16), 3)) = 0 Then ContinueLoop
				EndIf
				Return $asDriveLetter[$i1]
			EndIf
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
		Case '', 'con:'
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
	Local $i, $sFreeLetter, $WMIService, $WMIVolumes, $WMIMountPoints, $collectOutPut = ''

	$WMIService = ObjGet("winmgmts:\\.\root\cimv2")
	$WMIVolumes = $WMIService.ExecQuery("Select * from Win32_Volume Where DriveType=3 or DriveType=5")

	If IsObj($WMIVolumes) Then
		For $Volume In $WMIVolumes
			$collectOutPut = $collectOutPut & "DriveLetter: " & $Volume.DriveLetter & "    Name: " & $Volume.Name & "    DeviceID: " & $Volume.DeviceID & @CRLF
			If IsKeyword($Volume.DriveLetter) Or $Volume.DriveLetter = "" Then
				$sFreeLetter = _FreeLetter()
				$Volume.AddMountPoint($sFreeLetter)
				_LogOutN('Mount ' & $sFreeLetter & ' ' & $Volume.DeviceID)
			EndIf
		Next
		_LogOutN("Volume Information (MountAll)" & @CRLF & "------------------------------" & @CRLF & $collectOutPut)
	EndIf
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

Func _FileRead($sFile, $iMode = 0)
	Local $vData
	If $sFile = 'con:' Or $sFile = 'con' Then
		$vData = ConsoleRead()
	Else
		Local $hF = FileOpen($sFile, $iMode)
		Local $vData = FileRead($hF)
		FileClose($hF)
	EndIf
	Return $vData
EndFunc   ;==>_FileRead
