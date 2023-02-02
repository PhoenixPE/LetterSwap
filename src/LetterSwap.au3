#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_icon=ReMount1.ico
#AutoIt3Wrapper_Res_Comment=LetterSwap.exe
#AutoIt3Wrapper_Res_Description=LetterBootSwap.exe
#AutoIt3Wrapper_Res_Fileversion=1.0.0.6
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=@Nikzzzz
#Tidy_Parameters=/sfc
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/striponly
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt('TrayIconHide', 1)

$sAboot = "                @Nikzzzz 16.03.2011"
$sHelp = @ScriptName & "   [/?] [WinDir | /Auto | /Manual] [/BootDrive:y MarkerFile]" & @CRLF

Dim $aMountHost[1][2]
Dim $aMountGuest[1][2]

$sHiveSystemGuest = ""
$sBootDrive = ""
$s = ""
$i = 1
While $i <= $CmdLine[0]
	$vTemp = StringLower($CmdLine[$i])
	Switch $vTemp
		Case "/?"
			MsgBox(4096, @ScriptName & $sAboot, $sHelp)
		Case "/auto"
			$aDrives = DriveGetDrive("FIXED")
			For $k = 1 To $aDrives[0]
				If FileExists($aDrives[$k] & '\windows\system32\config\system') Then $sHiveSystemGuest = $aDrives[$k] & '\windows'
			Next
		Case "/manual"
			$sHiveSystemGuest = FileSelectFolder("Select OS", 1)
		Case "/bootdrive:y"
			If $i < $CmdLine[0] Then
				$i += 1
				$sMarkerFile = $CmdLine[$i]
				$aDrives = DriveGetDrive("All")
				For $k = 1 To $aDrives[0]
					If FileExists($aDrives[$k] & '\' & $sMarkerFile) Then $sBootDrive = $aDrives[$k]
				Next
			EndIf
		Case Else
			$sHiveSystemGuest = $CmdLine[$i]
	EndSwitch
	$i += 1
WEnd

FileInstall('ReMount.exe', @TempDir & '\$$rm$$.exe')
If $sBootDrive <> '' Then
	RunWait('"' & @TempDir & '\$$rm$$.exe" ' & $sBootDrive & ' y: -f', '', @SW_HIDE)
EndIf

If FileExists($sHiveSystemGuest & '\system32\config\system') Then
	For $i = 1 To 999
		$sValueName = RegEnumVal("HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices", $i)
		If @error <> 0 Then ExitLoop
		$sLetter = StringRegExpReplace($sValueName, "(?i)\\DosDevices\\([c-w]):", "\1")
		If @extended = 0 Then ContinueLoop
		ReDim $aMountHost[UBound($aMountHost, 1) + 1][2]
		$aMountHost[UBound($aMountHost, 1) - 1][0] = $sLetter
		$aMountHost[UBound($aMountHost, 1) - 1][1] = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices", $sValueName)
	Next

	RunWait('reg.exe load hklm\GuestSYSTEM "' & $sHiveSystemGuest & '\system32\config\system"', '', @SW_HIDE)
	For $i = 1 To 999
		$sValueName = RegEnumVal("HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevices", $i)
		If @error <> 0 Then ExitLoop
		$sLetter = StringRegExpReplace($sValueName, "(?i)\\DosDevices\\([c-w]):", "\1")
		If @extended = 0 Then ContinueLoop
		ReDim $aMountGuest[UBound($aMountGuest, 1) + 1][2]
		$aMountGuest[UBound($aMountGuest, 1) - 1][0] = $sLetter
		$aMountGuest[UBound($aMountGuest, 1) - 1][1] = RegRead("HKEY_LOCAL_MACHINE\GuestSYSTEM\MountedDevices", $sValueName)
	Next
	RunWait('reg.exe unload hklm\GuestSYSTEM', '', @SW_HIDE)
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
