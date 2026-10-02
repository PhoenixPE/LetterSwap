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
#AutoIt3Wrapper_Res_Fileversion=2026.8.17.89
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductVersion=2026.08.17.87
#AutoIt3Wrapper_Res_LegalCopyright=(c) Nikzzzz, Homes32 & Contributors
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; AutoIt 3.3.16.1

#include <WinAPIFiles.au3>
#include <WinAPIError.au3>
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
		& "  /Auto                                Find the first guest Windows OS." & @CRLF _
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
		& "  /IgnoreLetter <Letter>[...]          When used in conjunction with /Auto, /Manual, /WinDir, or MountAll ignore the specified" & @CRLF _
		& "                                       drive letters. (The system drive, A:, B: are always ignored.)" & @CRLF _
		& "                                         Ex. Letterswap.exe /IgnoreLetter def" & @CRLF _
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
Global $iAutoDetectSystemGuest = 0, $sNewBootDrive = '', $s = "", $i = 1, $iWait0 = 100, $iWait, $fSave = False, $sGuest = '', $sRestartExplorer = False, $sHostDrive = StringLeft(EnvGet('SourceDrive'), 1)
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
			$iAutoDetectSystemGuest = 1
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

Local $hRunTimer = TimerInit()
_LogOutN("===== LetterSwap v" & FileGetVersion(@ScriptFullPath) & " - Started " & _NowStamp() & " =====")
_LogOutN("Command Line: " & @ScriptName & " " & $CmdLineRaw)

; Ignore the SystemDrive
$sIgnoreLetter &= StringLeft(EnvGet('SystemDrive'), 1)

; /MountAll
If $iMountAll Then _MountAll()

; /HideLetter
If $iLetterClean Then _LetterClean('Removable;CDROM')

; /Auto
If $iAutoDetectSystemGuest Then
	$aDrives = DriveGetDrive("FIXED")
	For $k = 1 To $aDrives[0]
		If $aDrives[$k] <> EnvGet("SystemDrive") And FileExists($aDrives[$k] & '\windows\system32\config\system') Then
			$sSystemGuest = $aDrives[$k] & '\windows'
			ExitLoop
		EndIf
	Next
EndIf

; If we found a valid Guest Windows OS (or the user supplied a valid guest os via /Manual and/or /WinDir)
; get the list of mounted drives from it's registry and mount them.
If $sSystemGuest <> '' And FileExists($sSystemGuest & '\system32\config\system') Then
	_LogOutN('Auto Host/Guest Sync...')
	_LogOutN('  Host  System : ' & EnvGet('SystemRoot'))
	_LogOutN('  Guest System : ' & $sSystemGuest)
	$sGuest = StringLeft($sSystemGuest, 2)
	_MountGet($sHostKey, $aMountHost)
	_MountPrint('Host Volumes (' & $sHostKey & ')', $aMountHost)
	_RegLoadHive($sSystemGuest & '\system32\config\system', 'HKLM\GuestSYSTEM')
	_MountGet($sGuestKey, $aMountGuest)
	_RegUnLoadHive('HKLM\GuestSYSTEM')
	_MountPrint('Guest Volumes (' & $sGuestKey & ')', $aMountGuest)
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
	_LogOutN('  Searching for tag file "' & $sTagFile1 & '" (up to ' & ($iWait0 / 10) & 's)...')
	$iWait = $iWait0
	While $iWait >= 0
		$aDrives = DriveGetDrive('all')
		For $i = 1 To UBound($aDrives) - 1
			If Not FileExists($aDrives[$i] & '\' & $sTagFile1) Then ContinueLoop
			_LogOutN('  Found "' & $aDrives[$i] & '\' & $sTagFile1 & '" -> assigning ' & $sNewDrive1)
			_MountSwap($aDrives[$i], $sNewDrive1)
			ExitLoop 2
		Next
		$iWait -= 1
		Sleep(100)
	WEnd
	If $iWait < 0 Then _LogOutN('  Tag file not found, giving up.')
EndIf

; /Bootdrive
If $sNewBootDrive <> '' Then
	_LogOutN('  Searching for the boot drive (up to ' & ($iWait0 / 10) & 's)...')
	$iWait = $iWait0
	While $iWait >= 0
		$sBootDrive = _GetBootDrive($sTagFile)
		If $sBootDrive <> '' Then
			_LogOutN('  Found boot drive "' & $sBootDrive & '" -> assigning ' & $sNewBootDrive)
			_MountSwap($sBootDrive, $sNewBootDrive)
			ExitLoop
		EndIf
		$iWait -= 1
		Sleep(100)
	WEnd
	If $iWait < 0 Then _LogOutN('  Boot drive not found, giving up.')
EndIf

; /RestartExplorer
If $sRestartExplorer Then
	_LogOutN('Restarting Explorer')
	While ProcessExists("Explorer.exe")
		ProcessClose("Explorer.exe")
		Sleep(500)
	WEnd
	Run("Explorer.exe")
EndIf

; Display current Mount Points now that we are finished processing
_MountGet($sHostKey, $aMountHost)
_MountPrint('Final Host Volumes (' & $sHostKey & ')', $aMountHost)
_LogOutN()
_LogOutN("===== LetterSwap Finished " & _NowStamp() & " (elapsed " & StringFormat("%.1f", TimerDiff($hRunTimer) / 1000) & "s) =====" & @CRLF)
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

Func _GetFreeDriveLetter()
	Local $sFreeLetter = '', $i
	For $i = Asc("c") To Asc("z") ; Skip drives A: and B: as they are traditionally 'reserved' for floppy/ramdrive.
		If DriveGetType(Chr($i) & ':\') = '' Then
			$sFreeLetter = Chr($i) & ":\"
			ExitLoop
		EndIf
	Next
	Return $sFreeLetter
EndFunc   ;==>_GetFreeDriveLetter

Func _LetterClean($sdrivetype)
	_MountAll()
	Local $i, $asMount[1][3]
	_MountGet($sHostKey, $asMount)
	For $i = 1 To UBound($asMount) - 1
		If StringInStr($sdrivetype, DriveGetType($asMount[$i][1])) And (StringInStr("ab", $asMount[$i][1]) = 0) Then
			_WinAPI_GetVolumeInformation($asMount[$i][1] & ':\')
			If @error Then
				_WinAPI_DeleteVolumeMountPoint($asMount[$i][1] & ':\')
				_LogOutN('  Removing mount point ' & $asMount[$i][1] & ': (no media present)')
			EndIf
		EndIf
	Next
EndFunc   ;==>_LetterClean

Func _NowStamp()
	Return StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC)
EndFunc   ;==>_NowStamp

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

; Logs the real Win32 error for the WinAPI call that just failed. Must be called
; immediately after that call, before any other WinAPI activity, since GetLastError()
; reflects whichever call ran most recently.
Func _LogWinApiError($sAction)
	Local $iErr = _WinAPI_GetLastError()
	Local $sMsg = _WinAPI_GetLastErrorMessage()
	If $sMsg = '' Then $sMsg = 'no further details'
	_LogOutN('  ERROR: ' & $sAction & ' failed. Returned: ' & $iErr & ' (' & $sMsg & ')')
EndFunc   ;==>_LogWinApiError

Func _MountPrint($sTitle, ByRef $asMount)
	Local $i
	_LogOutN($sTitle)
	If UBound($asMount) <= 1 Then
		_LogOutN('  (none)')
		Return
	EndIf
	For $i = 1 To UBound($asMount) - 1
		_LogOutN('  ' & $asMount[$i][1] & ':   ' & $asMount[$i][0])
	Next
EndFunc   ;==>_MountPrint

Func _MountAll()
	Local $sFreeLetter, $WMIService, $WMIVolumes

	$WMIService = ObjGet("winmgmts:\\.\root\cimv2")
	$WMIVolumes = $WMIService.ExecQuery("Select * from Win32_Volume Where DriveType=3 or DriveType=5")

	If Not IsObj($WMIVolumes) Then Return

	_LogOutN('Mounting all drives...')
	For $Volume In $WMIVolumes
		If IsKeyword($Volume.DriveLetter) Or $Volume.DriveLetter = "" Then
			$sFreeLetter = _GetFreeDriveLetter()
			$Volume.AddMountPoint($sFreeLetter)
			_LogOutN('  Mounted ' & $sFreeLetter & '  ' & $Volume.DeviceID & '  (' & $Volume.Name & ')')
		Else
			_LogOutN('  ' & $Volume.DriveLetter & '  ' & $Volume.DeviceID & '  (' & $Volume.Name & ')')
		EndIf
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

; NOTE: _WinAPI_DeleteVolumeMountPoint() and _WinAPI_SetVolumeMountPoint()
; (WinAPIFiles.au3) don't set @error when the underlying Win32 call genuinely
; fails - @error there would only reflect a DllCall-marshalling problem,
; not a real API failure. We must check the return value directly.
Func _MountSwap($sDrive1, $sDrive2)
	$sDrive1 = StringLeft($sDrive1, 1) & ':\'
	$sDrive2 = StringLeft($sDrive2, 1) & ':\'
	If $sDrive1 = $sDrive2 Then Return 1
	Local $sLetter1 = StringLeft($sDrive1, 2), $sLetter2 = StringLeft($sDrive2, 2)
	Local $sGuid1 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive1) ; volume currently at drive1
	Local $sGuid2 = _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive2) ; volume currently at drive2

	_LogOutN('Swapping "' & $sLetter1 & '" <-> "' & $sLetter2 & '"...')

	While 1
		If $sGuid1 Then
			If Not _WinAPI_DeleteVolumeMountPoint($sDrive1) Then
				_LogWinApiError('Freeing ' & $sLetter1)
				$sGuid1 = ''
				$sGuid2 = ''
				ExitLoop
			EndIf
		EndIf
		If $sGuid2 Then
			If Not _WinAPI_DeleteVolumeMountPoint($sDrive2) Then
				_LogWinApiError('Freeing ' & $sLetter2)
				$sGuid2 = ''
				ExitLoop
			EndIf
		EndIf
		If $sGuid1 Then
			If Not _WinAPI_SetVolumeMountPoint($sDrive2, $sGuid1) Then
				_LogWinApiError('Mounting ' & $sLetter1 & ' onto ' & $sLetter2)
				ExitLoop
			ElseIf _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive2) <> $sGuid1 Then
				_LogOutN('  ERROR: Mounting ' & $sLetter1 & ' onto ' & $sLetter2 & ' reported success, but ' & $sLetter2 & ' does not match the expected volume.')
				ExitLoop
			EndIf
		EndIf
		If $sGuid2 Then
			If Not _WinAPI_SetVolumeMountPoint($sDrive1, $sGuid2) Then
				_LogWinApiError('Mounting ' & $sLetter2 & ' onto ' & $sLetter1)
				_WinAPI_DeleteVolumeMountPoint($sDrive2)
				ExitLoop
			ElseIf _WinAPI_GetVolumeNameForVolumeMountPoint($sDrive1) <> $sGuid2 Then
				_LogOutN('  ERROR: Mounting ' & $sLetter2 & ' onto ' & $sLetter1 & ' reported success, but ' & $sLetter1 & ' does not match the expected volume.')
				_WinAPI_DeleteVolumeMountPoint($sDrive2)
				ExitLoop
			EndIf
		EndIf
		_LogOutN('  SUCCESS: "' & $sLetter1 & '" <-> "' & $sLetter2 & '"')
		Return 1
	WEnd
	If $sGuid1 Then _WinAPI_SetVolumeMountPoint($sDrive1, $sGuid1) ; best-effort restore
	If $sGuid2 Then _WinAPI_SetVolumeMountPoint($sDrive2, $sGuid2) ; best-effort restore
	_LogOutN('  Swap failed - attempted to restore original letters.')
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
