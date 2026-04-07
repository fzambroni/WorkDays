#include-once

Func GenerateWorkdaysReportHTML($Year, $Full)

	Local $oObject = WKHtmlToX() ; Mantém o conversor original, que já funciona no seu ambiente

	Local $RegistryBase = "HKEY_CURRENT_USER\Software\WorkDays\" & $Year
	Local $OutputPath = "Workdays_Report_" & @MON & "_" & @MDAY & "_" & @YEAR & ".pdf"
	Local $OutputPathTemp = "Workdays_Report_" & @MON & "_" & @MDAY & "_" & @YEAR & ".html"
	Local $sHtmlPath = @ScriptDir & "\" & $OutputPathTemp
	Local $sPdfPath = @ScriptDir & "\" & $OutputPath

	Local $hFile = FileOpen($sHtmlPath, 2)
	If $hFile = -1 Then
		MsgBox(16, "Error", "Failed to create HTML file.")
		Return SetError(1, 0, 0)
	EndIf

	Local $CatNames[9] = ["OnSite", "Remote", "Holiday", "PTO", "Travel", "Sick", "Other", "Blank", "Weekends"]
	Local $Colors[9] = [ _
			_GetReportColorHTML("Color_OnSite", 0x00CC66), _
			_GetReportColorHTML("Color_Remote", 0x0080FF), _
			_GetReportColorHTML("Color_holiday", 0xFFFFCC), _
			_GetReportColorHTML("Color_PTO", 0x66FFFF), _
			_GetReportColorHTML("Color_Travel", 0xFF8000), _
			_GetReportColorHTML("Color_Sick", 0xFF6666), _
			"#DDDDDD", _
			_GetReportColorHTML("Color_Blank", 0xFFFFFF), _
			_GetReportColorHTML("Color_Weekend", 0xA0A0A0) _
	]
	Local $FontColors[9] = [ _
			_GetReportFontColorHTML("Font_OnSite"), _
			_GetReportFontColorHTML("Font_Remote"), _
			_GetReportFontColorHTML("Font_holiday"), _
			_GetReportFontColorHTML("Font_PTO"), _
			_GetReportFontColorHTML("Font_Travel"), _
			_GetReportFontColorHTML("Font_Sick"), _
			"#000000", _
			_GetReportFontColorHTML("Font_Blank"), _
			_GetReportFontColorHTML("Font_Weekend") _
	]

	Local $CategoryCount[4][9] = [[0]]
	Local $CategoryNotes[4][9]
	Local $QuarterStats[4][7]

	Local $TotalDays = 0
	Local $WorkDays = 0
	Local $RealOnSite = 0

	For $m = 12 To 1 Step -1
		Local $MonthKey = StringFormat("%02d", $m)
		Local $FullKey = $RegistryBase & "\" & $MonthKey
		Local $q = Int(($m - 1) / 3)
		Local $i = 1

		While 1
			Local $Day = RegEnumVal($FullKey, $i)
			If @error Then ExitLoop

			Local $RawVal = RegRead($FullKey, $Day)
			If @error Then
				$i += 1
				ContinueLoop
			EndIf

			Local $DateStr = $Year & "/" & $MonthKey & "/" & $Day
			Local $CatLetter = StringUpper(StringLeft($RawVal, 1))
			Local $Note = StringTrimLeft($RawVal, 1)
			If $Note = $RawVal Then $Note = ""

			Local $CatIndex = 6 ; Other por padrão
			If $CatLetter = "O" Then $CatIndex = 0
			If $CatLetter = "R" Then $CatIndex = 1
			If $CatLetter = "H" Then $CatIndex = 2
			If $CatLetter = "P" Then $CatIndex = 3
			If $CatLetter = "T" Then $CatIndex = 4
			If $CatLetter = "S" Then $CatIndex = 5
			If $CatLetter = "B" Or $RawVal = "" Then $CatIndex = 7
			If $CatLetter = "W" Then $CatIndex = 8

			If $CatIndex = 6 And $Note = "" Then
				$i += 1
				ContinueLoop
			EndIf

			$CategoryCount[$q][$CatIndex] += 1

			If $Note <> "" Then
				If StringInStr($Note, "/n", 0, 1) Then
					Local $Note_Splited = StringSplit($Note, "/n", 1)
					For $Count_Note = 1 To $Note_Splited[0]
						Local $sChunk = _HtmlAsciiEntityEncode($Note_Splited[$Count_Note])

						If $Count_Note = 1 Then
							If $Note_Splited[$Count_Note] <> "" Then
								$CategoryNotes[$q][$CatIndex] = "<li><b>" & _HtmlAsciiEntityEncode($DateStr) & ":</b> " & $sChunk & "</li>" & $CategoryNotes[$q][$CatIndex]
							EndIf
						Else
							If $Note_Splited[$Count_Note] <> "" Then
								$CategoryNotes[$q][$CatIndex] = "<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> " & $sChunk & "<br>" & $CategoryNotes[$q][$CatIndex]
							EndIf
						EndIf
					Next
				Else
					$CategoryNotes[$q][$CatIndex] = "<li><b>" & _HtmlAsciiEntityEncode($DateStr) & ":</b> " & _HtmlAsciiEntityEncode($Note) & "</li>" & $CategoryNotes[$q][$CatIndex]
				EndIf
			Else
				If $CatIndex <> 7 And $CatIndex <> 8 Then
					$CategoryNotes[$q][$CatIndex] = "<li>" & _HtmlAsciiEntityEncode($DateStr) & "</li>" & $CategoryNotes[$q][$CatIndex]
				EndIf
			EndIf

			$QuarterStats[$q][0] += 1

			If $CatLetter = "O" Or $CatLetter = "R" Or $CatLetter = "T" Or $CatLetter = "B" Or $CatLetter = "" Then
				$QuarterStats[$q][1] += 1
				$WorkDays += 1
			EndIf

			If $CatLetter = "O" Or $CatLetter = "T" Then
				$QuarterStats[$q][4] += 1
				$RealOnSite += 1
			EndIf

			$TotalDays += 1
			$i += 1
		WEnd
	Next

	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop

		Local $Expected = Ceiling(($QuarterStats[$q][1] / 5) * 3)
		Local $Actual = $QuarterStats[$q][4]

		$QuarterStats[$q][3] = $Expected
		$QuarterStats[$q][5] = $Expected - $Actual

		If $Actual > 0 Then
			$QuarterStats[$q][2] = Round(($Expected / $Actual), 2)
		Else
			$QuarterStats[$q][2] = 0
		EndIf
	Next

	Local $ExpectedTotal = Ceiling(($WorkDays / 5) * 3)
	Local $Ratio = 0
	If $WorkDays > 0 Then $Ratio = Round($RealOnSite / ($WorkDays / 5), 2)

	If $Full = 1 Then
		FileWriteLine($hFile, "<html><head><meta charset=""utf-8""><title>Workdays Report - DETAILED - " & $Year & "</title>")
	Else
		FileWriteLine($hFile, "<html><head><meta charset=""utf-8""><title>Workdays Report - SIMPLE - " & $Year & "</title>")
	EndIf

	FileWriteLine($hFile, "<style>body{font-family:Arial;} table{border-collapse:collapse;width:100%;margin-bottom:20px;} th,td{border:1px solid #ccc;padding:6px;} th{background:#f0f0f0;} .stat,.qstat{margin:10px 0;padding:10px;background:#eef;border-left:4px solid #88f;} ul{margin:0;padding-left:20px;} h2{margin-top:30px;}</style></head><body>")

	If $Full = 1 Then
		FileWriteLine($hFile, "<h1>Workdays Report - DETAILED - " & $Year & "</h1>")
	Else
		FileWriteLine($hFile, "<h1>Workdays Report - SIMPLE - " & $Year & "</h1>")
	EndIf

	FileWriteLine($hFile, "<div class='stat'><b>Total Days Recorded:</b> " & $TotalDays & "<br><b>Work Days:</b> " & $WorkDays & "<br><b>Ratio*:</b> " & $Ratio & "<br><b>Estimated OnSite*:</b> " & $ExpectedTotal & "<br><b>Real On-Site*:</b> " & $RealOnSite & "<br><b>Remaining*:</b> " & ($ExpectedTotal - $RealOnSite) & "<br>*These values are for reference only. For an accurate analysis, consider the quarterly data.</div>")

	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop

		Local $qRatio = 0
		If $QuarterStats[$q][1] > 0 Then $qRatio = Round($QuarterStats[$q][4] / ($QuarterStats[$q][1] / 5), 2)

		FileWriteLine($hFile, "<h2>Quarter " & ($q + 1) & "</h2>")
		FileWriteLine($hFile, "<div class='qstat'><b>Total Days:</b> " & $QuarterStats[$q][0] & "<br><b>Work Days:</b> " & $QuarterStats[$q][1] & "<br><b>Ratio:</b> " & $qRatio & "<br><b>Estimated OnSite:</b> " & $QuarterStats[$q][3] & "<br><b>Real On-Site:</b> " & $QuarterStats[$q][4] & "<br><b>Remaining:</b> " & $QuarterStats[$q][5] & "</div>")

		If $Full = 1 Then
			FileWriteLine($hFile, "<table><tr><th>Category</th><th>Count</th><th>Dates &amp; Notes</th></tr>")
		Else
			FileWriteLine($hFile, "<table><tr><th>Category</th><th>Count</th></tr>")
		EndIf

		For $c = 0 To 8
			If $CategoryCount[$q][$c] = 0 Then ContinueLoop

			FileWriteLine($hFile, "<tr style='background-color:" & $Colors[$c] & ";color:" & $FontColors[$c] & ";'><td><b>" & _HtmlAsciiEntityEncode($CatNames[$c]) & "</b></td><td>" & $CategoryCount[$q][$c] & "</td>")
			If $Full = 1 Then
				If $CategoryNotes[$q][$c] <> "" Then
					FileWriteLine($hFile, "<td><ul>" & $CategoryNotes[$q][$c] & "</ul></td></tr>")
				Else
					FileWriteLine($hFile, "<td>No details listed</td></tr>")
				EndIf
			EndIf
		Next
		FileWriteLine($hFile, "</table>")
	Next

	Local $YearlyTotals[9] = [0, 0, 0, 0, 0, 0, 0, 0, 0]
	For $q = 0 To 3
		For $c = 0 To 8
			$YearlyTotals[$c] += $CategoryCount[$q][$c]
		Next
	Next

	FileWriteLine($hFile, "<h2>Yearly Summary</h2>")
	FileWriteLine($hFile, "<table><tr><th>Category</th><th>Total Count</th></tr>")
	For $c = 0 To 8
		If $YearlyTotals[$c] > 0 Then
			FileWriteLine($hFile, "<tr style='background-color:" & $Colors[$c] & ";color:" & $FontColors[$c] & ";'><td><b>" & _HtmlAsciiEntityEncode($CatNames[$c]) & "</b></td><td>" & $YearlyTotals[$c] & "</td></tr>")
		EndIf
	Next
	FileWriteLine($hFile, "</table>")

	FileWriteLine($hFile, "<p style='color:gray;font-size:small;'>Generated on " & @YEAR & "/" & @MON & "/" & @MDAY & " at " & @HOUR & ":" & @MIN & "</p>")
	FileWriteLine($hFile, "<p style='color:gray;font-size:small;'>Develop by Fabricio Zambroni - Version: " & _HtmlAsciiEntityEncode(FileGetVersion(@ScriptFullPath)) & "</p>")
	FileWriteLine($hFile, "</body></html>")
	FileClose($hFile)

	$oObject.Input = $OutputPathTemp
	$oObject.Output = $OutputPath
	$oObject.Convert()

	If Not FileExists($sPdfPath) Then
		MsgBox(16, "Error", "Failed to convert HTML to PDF.")
		Return SetError(2, 0, 0)
	EndIf

	MsgBox(262208, "Report", "The report file was saved on: " & @CRLF & $sPdfPath)
	FileDelete($sHtmlPath)
	ShellExecute($sPdfPath)

	Return 1
EndFunc   ;==>GenerateWorkdaysReportHTML

Func _GetReportColorHTML($sRegValueName, $iDefaultColor)
	Local $vColor = RegRead("HKEY_CURRENT_USER\Software\WorkDays", $sRegValueName)
	If @error Or $vColor = "" Then $vColor = $iDefaultColor

	Local $iColor = Number($vColor)
	If $iColor < 0 Then $iColor = $iDefaultColor

	Return "#" & Hex($iColor, 6)
EndFunc   ;==>_GetReportColorHTML

Func _GetReportFontColorHTML($sRegValueName)
	Local $vFontIsWhite = RegRead("HKEY_CURRENT_USER\Software\WorkDays", $sRegValueName)
	If @error Then Return "#000000"

	If Number($vFontIsWhite) = 1 Then
		Return "#FFFFFF"
	EndIf

	Return "#000000"
EndFunc   ;==>_GetReportFontColorHTML

Func _HtmlAsciiEntityEncode($sText)
	If $sText = "" Then Return ""

	Local $sOut = ""
	Local $iLen = StringLen($sText)

	For $i = 1 To $iLen
		Local $ch = StringMid($sText, $i, 1)
		Local $cp = AscW($ch)

		Switch $ch
			Case "&"
				$sOut &= "&amp;"
			Case "<"
				$sOut &= "&lt;"
			Case ">"
				$sOut &= "&gt;"
			Case '"'
				$sOut &= "&quot;"
			Case "'"
				$sOut &= "&#39;"
			Case @CR
				; ignora, o relatório já controla as quebras
			Case @LF
				$sOut &= "<br>"
			Case Else
				If $cp < 32 Then
					; ignora controles
				ElseIf $cp > 126 Then
					$sOut &= "&#" & $cp & ";"
				Else
					$sOut &= $ch
				EndIf
		EndSwitch
	Next

	Return $sOut
EndFunc   ;==>_HtmlAsciiEntityEncode
