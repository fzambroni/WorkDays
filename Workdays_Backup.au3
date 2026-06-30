#include-once
; ======================================================================================================================
; WorkDays shared backup library
; Silent helpers used by WorkDays and the Outlook Agent.
; Backup format intentionally stays compatible with the original WorkDays restore routine: key=value lines only.
; ======================================================================================================================

Func _WD_Backup_DefaultFolder($sBaseDir = "")
	If $sBaseDir = "" Then $sBaseDir = @ScriptDir
	Return $sBaseDir & "\Backup"
EndFunc   ;==>_WD_Backup_DefaultFolder

Func _WD_Backup_FileName($sPrefix = "Backup")
	Return $sPrefix & "_" & @YEAR & "_" & @MON & "_" & @MDAY & "_" & @HOUR & @MIN & @SEC & ".bkp"
EndFunc   ;==>_WD_Backup_FileName

Func _WD_Backup_Create($sRegRoot, $sBackupFile = "", $sBackupDir = "", $sPrefix = "Backup")
	If $sRegRoot = "" Then Return SetError(1, 0, "")

	If $sBackupFile = "" Then
		If $sBackupDir = "" Then $sBackupDir = _WD_Backup_DefaultFolder()
		DirCreate($sBackupDir)
		$sBackupFile = $sBackupDir & "\" & _WD_Backup_FileName($sPrefix)
	Else
		Local $iSlash = StringInStr($sBackupFile, "\", 0, -1)
		If $iSlash > 0 Then DirCreate(StringLeft($sBackupFile, $iSlash - 1))
	EndIf

	Local $hFile = FileOpen($sBackupFile, 10)
	If $hFile = -1 Then Return SetError(2, @error, "")

	Local $i, $r, $d
	Local $sName, $sValue, $sYear, $sMonth, $sDay

	; Root settings used by WorkDays.
	For $i = 1 To 10000
		$sName = RegEnumVal($sRegRoot, $i)
		If @error Then ExitLoop
		$sValue = RegRead($sRegRoot, $sName)
		FileWriteLine($hFile, $sName & "=" & StringReplace($sValue, @CRLF, " /n"))
	Next

	; Calendar data only: YYYY\MM\DD.
	For $i = 1 To 10000
		$sYear = RegEnumKey($sRegRoot, $i)
		If @error Then ExitLoop
		If Not StringRegExp($sYear, "^\d{4}$") Then ContinueLoop

		For $r = 1 To 10000
			$sMonth = RegEnumKey($sRegRoot & "\" & $sYear, $r)
			If @error Then ExitLoop
			If Not StringRegExp($sMonth, "^\d{2}$") Then ContinueLoop

			For $d = 1 To 10000
				$sDay = RegEnumVal($sRegRoot & "\" & $sYear & "\" & $sMonth, $d)
				If @error Then ExitLoop
				If Not StringRegExp($sDay, "^\d{2}$") Then ContinueLoop
				$sValue = RegRead($sRegRoot & "\" & $sYear & "\" & $sMonth, $sDay)
				FileWriteLine($hFile, $sYear & "\" & $sMonth & "\" & $sDay & "=" & StringReplace($sValue, @CRLF, " /n"))
			Next
		Next
	Next

	FileClose($hFile)
	Return $sBackupFile
EndFunc   ;==>_WD_Backup_Create

Func _WD_Backup_CountRegistryDateRecords($sRegRoot, $sStartISO = "", $sEndISO = "")
	Local $iCount = 0
	Local $i, $r, $d
	Local $sYear, $sMonth, $sDay, $sISO, $sValue

	For $i = 1 To 10000
		$sYear = RegEnumKey($sRegRoot, $i)
		If @error Then ExitLoop
		If Not StringRegExp($sYear, "^\d{4}$") Then ContinueLoop

		For $r = 1 To 10000
			$sMonth = RegEnumKey($sRegRoot & "\" & $sYear, $r)
			If @error Then ExitLoop
			If Not StringRegExp($sMonth, "^\d{2}$") Then ContinueLoop

			For $d = 1 To 10000
				$sDay = RegEnumVal($sRegRoot & "\" & $sYear & "\" & $sMonth, $d)
				If @error Then ExitLoop
				If Not StringRegExp($sDay, "^\d{2}$") Then ContinueLoop
				$sISO = $sYear & "-" & $sMonth & "-" & $sDay
				If $sStartISO <> "" And $sISO < $sStartISO Then ContinueLoop
				If $sEndISO <> "" And $sISO > $sEndISO Then ContinueLoop
				$sValue = RegRead($sRegRoot & "\" & $sYear & "\" & $sMonth, $sDay)
				If String($sValue) <> "" Then $iCount += 1
			Next
		Next
	Next

	Return $iCount
EndFunc   ;==>_WD_Backup_CountRegistryDateRecords

Func _WD_Backup_CountFileDateRecords($sBackupFile)
	Local $hFile = FileOpen($sBackupFile, 0)
	If $hFile = -1 Then Return SetError(1, @error, 0)

	Local $iCount = 0
	Local $sLine, $iEq, $sKey
	While 1
		$sLine = FileReadLine($hFile)
		If @error = -1 Then ExitLoop
		$iEq = StringInStr($sLine, "=")
		If $iEq = 0 Then ContinueLoop
		$sKey = StringLeft($sLine, $iEq - 1)
		If StringRegExp($sKey, "^\d{4}\\\d{2}\\\d{2}$") Then $iCount += 1
	WEnd

	FileClose($hFile)
	Return $iCount
EndFunc   ;==>_WD_Backup_CountFileDateRecords
