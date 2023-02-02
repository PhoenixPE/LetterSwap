#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ReMount1.ico
#AutoIt3Wrapper_Res_Comment=LetterSwap.exe
#AutoIt3Wrapper_Res_Description=LetterBootSwap.exe
#AutoIt3Wrapper_Res_Fileversion=1.1.0.12
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=@Nikzzzz
#Tidy_Parameters=/sfc
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/striponly
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt('TrayIconHide', 1)

$sAboot = "                @Nikzzzz 24.03.2011"
$sHelp = @ScriptName & " [WinDir | /Auto | /Manual] [/BootDrive:y MarkerFile] [/log LogFile]" & @CRLF

If $CmdLine[0] = 0 Then
	MsgBox(4096, @ScriptName & $sAboot, $sHelp)
	Exit
EndIf

Dim $aMountHost[1][2]
Dim $aMountGuest[1][2]
$sLogFile = ''
$sHiveSystemGuest = ""
$sBootDrive = ""
$s = ""
$i = 1
While $i <= $CmdLine[0]
	$vTemp = StringLower($CmdLine[$i])
	Switch $vTemp
		Case "/auto"
			$aDrives = DriveGetDrive("FIXED")
			For $k = 1 To $aDrives[0]
				If FileExists($aDrives[$k] & '\windows\system32\config\system') Then $sHiveSystemGuest = $aDrives[$k] & '\windows'
			Next
		Case "/manual"
			$sHiveSystemGuest = FileSelectFolder("Select OS  (Example: d:\Windows)", 1)
		Case "/bootdrive:y"
			If $i < $CmdLine[0] Then
				$i += 1
				$sMarkerFile = $CmdLine[$i]
				$aDrives = DriveGetDrive("All")
				For $k = 1 To $aDrives[0]
					If FileExists($aDrives[$k] & '\' & $sMarkerFile) Then $sBootDrive = $aDrives[$k]
				Next
			EndIf
		Case "/log"
			If $i < $CmdLine[0] Then
				$i += 1
				$sLogFile = $CmdLine[$i]
			EndIf
		Case Else
			$sHiveSystemGuest = $CmdLine[$i]
	EndSwitch
	$i += 1
WEnd

LogOut("Command line:" & @CRLF & $CmdLineRaw & @CRLF)

FileInstall('ReMount.exe', @TempDir & '\$$rm$$.exe')
If $sBootDrive <> '' Then
	LogOut("Found BootDrive " & $sBootDrive & @CRLF)
	LogOut("Swap letter " & $sBootDrive & ' <> y:' & @CRLF)
	RunWait('"' & @TempDir & '\$$rm$$.exe" ' & $sBootDrive & ' y: -f', '', @SW_HIDE)
EndIf

If FileExists($sHiveSystemGuest & '\system32\config\system') Then
	LogOut("Host system:" & @CRLF)
	For $i = 1 To 999
		$sValueName = RegEnumVal("HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices", $i)
		If @error <> 0 Then ExitLoop
		$sLetter = StringRegExpReplace($sValueName, "(?i)\\DosDevices\\([c-w]):", "\1")
		If @extended = 0 Then ContinueLoop
		ReDim $aMountHost[UBound($aMountHost, 1) + 1][2]
		$aMountHost[UBound($aMountHost, 1) - 1][0] = $sLetter
		$aMountHost[UBound($aMountHost, 1) - 1][1] = Conv(RegRead("HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices", $sValueName))
		LogOut($sLetter & ' - ' & $aMountHost[UBound($aMountHost, 1) - 1][1] & @CRLF)
	Next
	LogOut("Guest system: " & $sHiveSystemGuest & @CRLF)
	RunWait('reg.exe load hklm\GuestSYSTEM "' & $sHiveSystemGuest & '\system32\config\system"', '', @SW_HIDE)
	For $l = 0 To 9
		$sKey = "HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevice"
		If $l = 0 Then
			$sKey &= 's'
		Else
			$sKey &= $l
		EndIf
		$sValueName = RegEnumVal($sKey, 1)
		If @error <> 0 Then ContinueLoop
		LogOut("Key: " & StringReplace($sKey, 'GuestSYSTEM', 'SYSTEM') & @CRLF)
		For $i = 1 To 999
			$sValueName = RegEnumVal($sKey, $i)
			If @error <> 0 Then ExitLoop
			$sLetter = StringRegExpReplace($sValueName, "(?i)\\DosDevices\\([c-w]):", "\1")
			If @extended = 0 Then ContinueLoop
			ReDim $aMountGuest[UBound($aMountGuest, 1) + 1][2]
			$aMountGuest[UBound($aMountGuest, 1) - 1][0] = $sLetter
			$aMountGuest[UBound($aMountGuest, 1) - 1][1] = Conv(RegRead($sKey, $sValueName))
			LogOut($sLetter & ' - ' & $aMountGuest[UBound($aMountGuest, 1) - 1][1] & @CRLF)
		Next
		RunWait('reg.exe unload hklm\GuestSYSTEM', '', @SW_HIDE)
	Next
	For $i = 1 To UBound($aMountHost, 1) - 1
		$sLetterHost = $aMountHost[$i][0]
		$sLetterGuest = ''
		For $j = 1 To UBound($aMountGuest, 1) - 1
			If $aMountHost[$i][1] = $aMountGuest[$j][1] Then
				$sLetterGuest = $aMountGuest[$j][0]
				ExitLoop
			EndIf
		Next
		If $aMountHost[$i][0] = $sLetterGuest Or $sLetterGuest = '' Then ContinueLoop
		LogOut("Swap letter " & $sLetterHost & ': <> ' & $sLetterGuest & ':' & @CRLF)

		For $k = 1 To UBound($aMountHost, 1) - 1
			$iLetterHostSwap = 0
			If $aMountHost[$k][0] = $sLetterGuest Then
				$iLetterHostSwap = $k
				ExitLoop
			EndIf
		Next

		If $iLetterHostSwap <> 0 Then
			RunWait('"' & @TempDir & '\$$rm$$.exe" ' & $sLetterHost & ': ' & $sLetterGuest & ': -s -f', '', @SW_HIDE)
			$aMountHost[$i][0] = $sLetterGuest
			$aMountHost[$k][0] = $sLetterHost
		Else
			$aMountHost[$i][0] = $sLetterGuest
			RunWait('"' & @TempDir & '\$$rm$$.exe" ' & $sLetterHost & ': ' & $sLetterGuest & ': -f', '', @SW_HIDE)
		EndIf
	Next
EndIf

FileDelete(@TempDir & '\$$rm$$.exe')

Func Conv($bStr)
	If BinaryMid($bStr, 3, 4) = Binary("0x3f003f00") Then
		Return BinaryToString($bStr, 2)
	Else
		Return String($bStr)
	EndIf
EndFunc   ;==>Conv

Func LogOut($sStr)
	If $sLogFile = '' Then Return
	FileWrite($sLogFile, $sStr)
EndFunc   ;==>LogOut