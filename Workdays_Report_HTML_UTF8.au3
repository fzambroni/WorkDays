#include-once

Func GenerateWorkdaysReportHTML($Year, $Full)

	Local $oObject = WKHtmlToX()

	Local $ReportType = "Simple"
	If $Full = 1 Then $ReportType = "Detailed"

	Local $RegistryBase = "HKEY_CURRENT_USER\Software\WorkDays\" & $Year
	Local $OutputPath = "Workdays_" & $ReportType & "_Report_" & $Year & "_" & @MON & "_" & @MDAY & "_" & @YEAR & ".pdf"
	Local $OutputPathTemp = "Workdays_" & $ReportType & "_Report_" & $Year & "_" & @MON & "_" & @MDAY & "_" & @YEAR & ".html"
	Local $sHtmlPath = @ScriptDir & "\Reports\" & $OutputPathTemp
	Local $sPdfPath = @ScriptDir & "\Reports\" & $OutputPath

	Local $hFile = FileOpen($sHtmlPath, 10)
	If $hFile = -1 Then
		MsgBox(16, "Error", "Failed to create the HTML report.")
		Return SetError(1, 0, 0)
	EndIf

	Local $CatNames[9] = ["On-Site", "Remote", "Holiday", "PTO", "Travel", "Sick", "Other", "Blank", "Weekend"]
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

	Local $CategoryCount[4][9]
	Local $CategoryNotes[4][9]
	Local $QuarterStats[4][7]
	Local $YearlyTotals[9]

	Local $TotalDays = 0
	Local $WorkDays = 0
	Local $RealOnSite = 0
	Local $NotesCount = 0
	Local $UnknownCount = 0

	For $m = 1 To 12
		Local $q = Int(($m - 1) / 3)
		Local $DaysInMonth = _WD_ReportDaysInMonth(Number($Year), $m)

		; Read dates in calendar order. RegEnumVal can return registry values in a
		; non-calendar order, which makes detailed reports harder to review.
		For $d = 1 To $DaysInMonth
			Local $Day = ""
			Local $RawVal = _WD_ReportReadDayValue($RegistryBase, $m, $d, $Day)
			If @error Then ContinueLoop

			Local $CatLetter = StringUpper(StringLeft($RawVal, 1))
			Local $Note = StringTrimLeft($RawVal, 1)
			If $Note = $RawVal Then $Note = ""

			Local $CatIndex = _WD_ReportCategoryIndex($RawVal)
			Local $DateStr = StringFormat("%04d-%02d-%02d", Number($Year), $m, $d)

			; Ignore completely empty custom/unknown entries. Valid blank days are still
			; captured as the Blank category by _WD_ReportCategoryIndex().
			If $CatIndex = 6 And $Note = "" Then ContinueLoop

			$CategoryCount[$q][$CatIndex] += 1
			$YearlyTotals[$CatIndex] += 1
			$QuarterStats[$q][0] += 1
			$TotalDays += 1

			If $CatIndex = 6 Then $UnknownCount += 1

			If _WD_ReportIsWorkDay($CatLetter) Then
				$QuarterStats[$q][1] += 1
				$WorkDays += 1
			EndIf

			If $CatLetter = "O" Or $CatLetter = "T" Then
				$QuarterStats[$q][2] += 1
				$RealOnSite += 1
			EndIf

			If $Note <> "" Then
				$NotesCount += 1
				$QuarterStats[$q][6] += 1
			EndIf

			If $Full = 1 Then
				If $CatIndex <> 7 And $CatIndex <> 8 Then
					Local $sDetailNote = "<span class='muted'>No note</span>"
					If $Note <> "" Then $sDetailNote = _WD_ReportNoteToHtml($Note)
					$CategoryNotes[$q][$CatIndex] &= "<div class='detail-entry'><span class='detail-date'>" & _HtmlAsciiEntityEncode($DateStr) & "</span><span class='detail-note'>" & $sDetailNote & "</span></div>"
				EndIf
			EndIf
		Next
	Next

	For $q = 0 To 3
		$QuarterStats[$q][3] = Ceiling(($QuarterStats[$q][1] / 5) * 3) ; expected on-site
		$QuarterStats[$q][4] = $QuarterStats[$q][2] - $QuarterStats[$q][3] ; signed gap
		If $QuarterStats[$q][3] > 0 Then
			$QuarterStats[$q][5] = Round(($QuarterStats[$q][2] / $QuarterStats[$q][3]) * 100, 0)
		Else
			$QuarterStats[$q][5] = 0
		EndIf
	Next

	Local $ExpectedTotal = Ceiling(($WorkDays / 5) * 3)
	Local $GapTotal = $RealOnSite - $ExpectedTotal
	Local $RemainingTotal = $ExpectedTotal - $RealOnSite
	If $RemainingTotal < 0 Then $RemainingTotal = 0

	Local $CompliancePct = 0
	If $ExpectedTotal > 0 Then $CompliancePct = Round(($RealOnSite / $ExpectedTotal) * 100, 0)

	Local $Ratio = 0
	If $WorkDays > 0 Then $Ratio = Round($RealOnSite / ($WorkDays / 5), 2)

	Local $StatusText = _WD_ReportStatusText($GapTotal)
	Local $StatusClass = _WD_ReportStatusClass($GapTotal)

	FileWriteLine($hFile, "<html><head><meta charset=""utf-8""><title>" & _HtmlAsciiEntityEncode($ReportType) & " Workdays Report - " & _HtmlAsciiEntityEncode($Year) & "</title>")
	FileWriteLine($hFile, "<style>")
	FileWriteLine($hFile, "@page{margin:14mm;} body{font-family:Arial,Helvetica,sans-serif;color:#1f2933;margin:0;background:#f4f6f8;} .page{width:1040px;margin:0 auto;background:#fff;padding:30px 34px;} .hero{background:#102a43;color:#fff;border-radius:12px;padding:22px 24px;margin-bottom:18px;} h1{margin:0;font-size:28px;letter-spacing:.2px;} h2{font-size:18px;margin:26px 0 10px;color:#102a43;border-bottom:2px solid #d9e2ec;padding-bottom:7px;} h3{font-size:15px;margin:20px 0 8px;color:#243b53;} .subtitle{color:#d9e2ec;margin-top:7px;font-size:13px;} .section-note{color:#627d98;font-size:12px;margin:-4px 0 10px;} .cards{width:100%;border-collapse:separate;border-spacing:10px;margin:10px -10px 12px -10px;} .card{border:1px solid #d9e2ec;border-radius:10px;padding:12px;background:#fbfcfd;vertical-align:top;} .label{font-size:10px;text-transform:uppercase;color:#829ab1;letter-spacing:.6px;} .value{font-size:24px;font-weight:bold;margin-top:4px;color:#102a43;} .small{font-size:12px;color:#627d98;line-height:1.35;} table{border-collapse:collapse;width:100%;margin:10px 0 18px;} th{background:#edf2f7;color:#243b53;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.45px;} th,td{border:1px solid #d9e2ec;padding:8px 9px;font-size:12px;vertical-align:top;} tr:nth-child(even){background:#fbfcfd;} .right{text-align:right;} .center{text-align:center;} .pill{display:inline-block;border-radius:999px;padding:4px 9px;font-size:11px;font-weight:bold;min-width:72px;text-align:center;} .status-ok{color:#0b6b3a;font-weight:bold;} .status-watch{color:#9a5b00;font-weight:bold;} .status-bad{color:#b42318;font-weight:bold;} .callout{border-left:5px solid #486581;background:#f0f4f8;padding:12px 14px;margin:14px 0;font-size:13px;line-height:1.45;} .bar-track{height:8px;background:#edf2f7;border-radius:20px;overflow:hidden;} .bar-fill{height:8px;border-radius:20px;} .count-cell{font-size:14px;font-weight:bold;color:#102a43;} .detail-entry{border-bottom:1px solid #edf2f7;padding:5px 0;line-height:1.35;} .detail-entry:last-child{border-bottom:0;} .detail-date{display:inline-block;width:90px;font-weight:bold;color:#334e68;} .detail-note{color:#334e68;} .muted{color:#829ab1;} .footer{margin-top:28px;border-top:1px solid #d9e2ec;padding-top:10px;color:#829ab1;font-size:11px;line-height:1.4;} .avoid-break{page-break-inside:avoid;}")
	FileWriteLine($hFile, "</style></head><body><div class='page'>")

	FileWriteLine($hFile, "<div class='hero'><h1>" & _HtmlAsciiEntityEncode($ReportType) & " Workdays Report</h1><div class='subtitle'>Year " & _HtmlAsciiEntityEncode($Year) & " &bull; generated on " & @YEAR & "/" & @MON & "/" & @MDAY & " at " & @HOUR & ":" & @MIN & "</div></div>")

	FileWriteLine($hFile, "<table class='cards'><tr>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Overall status</div><div class='value " & $StatusClass & "'>" & _HtmlAsciiEntityEncode($StatusText) & "</div><div class='small'>Based on a 3 on-site days per 5 workdays reference target.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Tracked days</div><div class='value'>" & $TotalDays & "</div><div class='small'>" & $WorkDays & " tracked workday(s) included in the target calculation.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Expected on-site</div><div class='value'>" & $ExpectedTotal & "</div><div class='small'>Calculated as ceiling(workdays / 5 * 3).</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Actual on-site</div><div class='value'>" & $RealOnSite & "</div><div class='small'>On-Site + Travel days.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Gap</div><div class='value " & $StatusClass & "'>" & _WD_ReportSignedNumber($GapTotal) & "</div><div class='small'>" & $CompliancePct & "% of expected coverage. " & $RemainingTotal & " day(s) remaining to target.</div></td>")
	FileWriteLine($hFile, "</tr></table>")

	FileWriteLine($hFile, "<div class='callout'><b>Report logic:</b> Travel is counted as on-site coverage. Remote and Blank are counted as workdays. Holiday, PTO, Sick, and Weekend entries are excluded from the on-site expectation. The current ratio is <b>" & $Ratio & "</b> on-site/travel days per 5 workdays.</div>")

	FileWriteLine($hFile, "<h2>Quarterly dashboard</h2>")
	FileWriteLine($hFile, "<div class='section-note'>Use this section to quickly validate whether the year is balanced or concentrated in a specific quarter.</div>")
	FileWriteLine($hFile, "<table><tr><th>Quarter</th><th class='right'>Tracked days</th><th class='right'>Workdays</th><th class='right'>Expected on-site</th><th class='right'>Actual on-site</th><th class='right'>Gap</th><th class='right'>Compliance</th><th class='right'>Notes</th><th>Status</th></tr>")
	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop
		Local $qClass = _WD_ReportStatusClass($QuarterStats[$q][4])
		FileWriteLine($hFile, "<tr><td><b>Q" & ($q + 1) & "</b></td><td class='right'>" & $QuarterStats[$q][0] & "</td><td class='right'>" & $QuarterStats[$q][1] & "</td><td class='right'>" & $QuarterStats[$q][3] & "</td><td class='right'>" & $QuarterStats[$q][2] & "</td><td class='right " & $qClass & "'>" & _WD_ReportSignedNumber($QuarterStats[$q][4]) & "</td><td class='right'>" & $QuarterStats[$q][5] & "%</td><td class='right'>" & $QuarterStats[$q][6] & "</td><td class='" & $qClass & "'>" & _HtmlAsciiEntityEncode(_WD_ReportStatusText($QuarterStats[$q][4])) & "</td></tr>")
	Next
	FileWriteLine($hFile, "</table>")

	FileWriteLine($hFile, "<h2>Category summary</h2>")
	FileWriteLine($hFile, "<div class='section-note'>The colors below follow the same category color settings configured in the application.</div>")
	FileWriteLine($hFile, "<table><tr><th>Category</th><th class='right'>Total days</th><th class='right'>Share</th><th>Distribution</th></tr>")
	For $c = 0 To 8
		If $YearlyTotals[$c] = 0 Then ContinueLoop
		Local $Share = 0
		If $TotalDays > 0 Then $Share = Round(($YearlyTotals[$c] / $TotalDays) * 100, 1)
		FileWriteLine($hFile, "<tr><td><span class='pill' style='background:" & $Colors[$c] & ";color:" & $FontColors[$c] & ";'>" & _HtmlAsciiEntityEncode($CatNames[$c]) & "</span></td><td class='right count-cell'>" & $YearlyTotals[$c] & "</td><td class='right'>" & $Share & "%</td><td><div class='bar-track'><div class='bar-fill' style='width:" & $Share & "%;background:" & $Colors[$c] & ";'></div></div></td></tr>")
	Next
	FileWriteLine($hFile, "</table>")

	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop

		FileWriteLine($hFile, "<div class='avoid-break'>")
		FileWriteLine($hFile, "<h2>Quarter " & ($q + 1) & " category breakdown</h2>")

		If $Full = 1 Then
			FileWriteLine($hFile, "<table><tr><th>Category</th><th class='right'>Count</th><th>Dates &amp; notes</th></tr>")
		Else
			FileWriteLine($hFile, "<table><tr><th>Category</th><th class='right'>Count</th><th class='right'>Share of quarter</th></tr>")
		EndIf

		For $c = 0 To 8
			If $CategoryCount[$q][$c] = 0 Then ContinueLoop

			Local $QuarterShare = 0
			If $QuarterStats[$q][0] > 0 Then $QuarterShare = Round(($CategoryCount[$q][$c] / $QuarterStats[$q][0]) * 100, 1)

			If $Full = 1 Then
				Local $sDetails = $CategoryNotes[$q][$c]
				If $sDetails = "" Then $sDetails = "<span class='muted'>No detail listed for this category.</span>"
				FileWriteLine($hFile, "<tr><td><span class='pill' style='background:" & $Colors[$c] & ";color:" & $FontColors[$c] & ";'>" & _HtmlAsciiEntityEncode($CatNames[$c]) & "</span></td><td class='right count-cell'>" & $CategoryCount[$q][$c] & "</td><td>" & $sDetails & "</td></tr>")
			Else
				FileWriteLine($hFile, "<tr><td><span class='pill' style='background:" & $Colors[$c] & ";color:" & $FontColors[$c] & ";'>" & _HtmlAsciiEntityEncode($CatNames[$c]) & "</span></td><td class='right count-cell'>" & $CategoryCount[$q][$c] & "</td><td class='right'>" & $QuarterShare & "%</td></tr>")
			EndIf
		Next

		FileWriteLine($hFile, "</table>")
		FileWriteLine($hFile, "</div>")
	Next

	FileWriteLine($hFile, "<div class='footer'>Generated by Work Days &bull; Developed by Fabricio Zambroni &bull; Version: " & _HtmlAsciiEntityEncode(FileGetVersion(@ScriptFullPath)) & "<br>Disclaimer: this report uses the data currently stored in the WorkDays registry database and applies the same reference rule used by the app: expected on-site days = ceiling(workdays / 5 * 3).</div>")
	FileWriteLine($hFile, "</div></body></html>")
	FileClose($hFile)

;~ 	$oObject.Input = $OutputPathTemp
;~ 	$oObject.Output = $OutputPath

	$oObject.Input = $sHtmlPath
	$oObject.Output = $sPdfPath

	$oObject.Convert()

	If Not FileExists($sPdfPath) Then
		MsgBox(16, "Error", "Failed to convert the HTML report to PDF.")
		Return SetError(2, 0, 0)
	EndIf

	MsgBox(262208, $ReportType & " Report", "The " & StringLower($ReportType) & " report file was saved on: " & @CRLF & $sPdfPath)
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


Func GenerateWorkdaysProfessionalReportHTML($Year)

	Local $oObject = WKHtmlToX()

	Local $RegistryBase = "HKEY_CURRENT_USER\Software\WorkDays\" & $Year
	Local $OutputPath = "Workdays_Analytical_Report_" & $Year & "_" & @MON & "_" & @MDAY & "_" & @YEAR & ".pdf"
	Local $OutputPathTemp = "Workdays_Analytical_Report_" & $Year & "_" & @MON & "_" & @MDAY & "_" & @YEAR & ".html"
	Local $sHtmlPath = @ScriptDir & "\Reports\" & $OutputPathTemp
	Local $sPdfPath = @ScriptDir & "\Reports\" & $OutputPath

	Local $hFile = FileOpen($sHtmlPath, 10)
	If $hFile = -1 Then
		MsgBox(16, "Error", "Failed to create Analytical HTML report.")
		Return SetError(1, 0, 0)
	EndIf

	Local $CatNames[9] = ["On-Site", "Remote", "Holiday", "PTO", "Travel", "Sick", "Other", "Blank", "Weekends"]
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

	Local $CategoryTotal[9] = [0, 0, 0, 0, 0, 0, 0, 0, 0]
	Local $QuarterStats[4][8]
	Local $MonthCategoryCount[13][9]
	Local $MonthWorkDays[13]
	Local $MonthOnSite[13]
	Local $MonthTotalDays[13]
	Local $MonthNotesCount[13]

	Local $TotalDays = 0
	Local $WorkDays = 0
	Local $RealOnSite = 0
	Local $NotesCount = 0
	Local $UnknownCount = 0
	Local $sNotesRows = ""

	For $m = 1 To 12
		Local $q = Int(($m - 1) / 3)
		Local $DaysInMonth = _WD_ReportDaysInMonth(Number($Year), $m)

		; Read the registry in calendar order. RegEnumVal can return the days in reverse
		; order on some Windows installations, which made the Notes and exceptions log
		; show each month from the last day back to the first.
		For $d = 1 To $DaysInMonth
			Local $Day = ""
			Local $RawVal = _WD_ReportReadDayValue($RegistryBase, $m, $d, $Day)
			If @error Then ContinueLoop

			Local $CatLetter = StringUpper(StringLeft($RawVal, 1))
			Local $Note = StringTrimLeft($RawVal, 1)
			If $Note = $RawVal Then $Note = ""

			Local $CatIndex = _WD_ReportCategoryIndex($RawVal)
			Local $DateStr = StringFormat("%04d-%02d-%02d", Number($Year), $m, $d)

			$CategoryTotal[$CatIndex] += 1
			$MonthCategoryCount[$m][$CatIndex] += 1
			$QuarterStats[$q][0] += 1 ; total tracked days
			$QuarterStats[$q][7] += 0
			$MonthTotalDays[$m] += 1
			$TotalDays += 1

			If $CatIndex = 6 Then $UnknownCount += 1

			If _WD_ReportIsWorkDay($CatLetter) Then
				$QuarterStats[$q][1] += 1
				$MonthWorkDays[$m] += 1
				$WorkDays += 1
			Else
				$QuarterStats[$q][6] += 1
			EndIf

			If $CatLetter = "O" Or $CatLetter = "T" Then
				$QuarterStats[$q][2] += 1
				$MonthOnSite[$m] += 1
				$RealOnSite += 1
			EndIf

			If $Note <> "" Then
				$NotesCount += 1
				$MonthNotesCount[$m] += 1
				$QuarterStats[$q][7] += 1
				$sNotesRows &= "<tr><td>" & _HtmlAsciiEntityEncode($DateStr) & "</td><td><span class='pill' style='background:" & $Colors[$CatIndex] & ";color:" & $FontColors[$CatIndex] & ";'>" & _HtmlAsciiEntityEncode($CatNames[$CatIndex]) & "</span></td><td>" & _WD_ReportNoteToHtml($Note) & "</td></tr>"
			EndIf
		Next
	Next

	For $q = 0 To 3
		$QuarterStats[$q][3] = Ceiling(($QuarterStats[$q][1] / 5) * 3) ; expected on-site
		$QuarterStats[$q][4] = $QuarterStats[$q][2] - $QuarterStats[$q][3] ; gap
		If $QuarterStats[$q][3] > 0 Then
			$QuarterStats[$q][5] = Round(($QuarterStats[$q][2] / $QuarterStats[$q][3]) * 100, 0)
		Else
			$QuarterStats[$q][5] = 0
		EndIf
	Next

	Local $ExpectedTotal = Ceiling(($WorkDays / 5) * 3)
	Local $GapTotal = $RealOnSite - $ExpectedTotal
	Local $CompliancePct = 0
	If $ExpectedTotal > 0 Then $CompliancePct = Round(($RealOnSite / $ExpectedTotal) * 100, 0)

	Local $RemoteDays = $CategoryTotal[1]
	Local $TravelDays = $CategoryTotal[4]
	Local $NonWorkingDays = $CategoryTotal[2] + $CategoryTotal[3] + $CategoryTotal[5] + $CategoryTotal[8]
	Local $StatusText = _WD_ReportStatusText($GapTotal)
	Local $StatusClass = _WD_ReportStatusClass($GapTotal)

	Local $WorstQuarter = 0
	Local $WorstGap = 99999
	For $q = 0 To 3
		If $QuarterStats[$q][0] > 0 And $QuarterStats[$q][4] < $WorstGap Then
			$WorstGap = $QuarterStats[$q][4]
			$WorstQuarter = $q + 1
		EndIf
	Next

	FileWriteLine($hFile, "<html><head><meta charset=""utf-8""><title>Analytical Workdays Report - " & _HtmlAsciiEntityEncode($Year) & "</title>")
	FileWriteLine($hFile, "<style>")
	FileWriteLine($hFile, "body{font-family:Arial,Helvetica,sans-serif;color:#1f2933;margin:0;background:#f4f6f8;} .page{width:1040px;margin:0 auto;background:#fff;padding:30px 34px;} h1{margin:0;font-size:28px;color:#102a43;} h2{font-size:18px;margin:28px 0 10px;color:#102a43;border-bottom:2px solid #d9e2ec;padding-bottom:6px;} h3{font-size:14px;margin:16px 0 8px;color:#334e68;} .subtitle{color:#627d98;margin-top:6px;} .header{border-bottom:4px solid #243b53;padding-bottom:18px;margin-bottom:18px;} .cards{width:100%;border-collapse:separate;border-spacing:10px;margin:8px -10px 12px -10px;} .card{border:1px solid #d9e2ec;border-radius:8px;padding:12px;background:#fbfcfd;vertical-align:top;} .label{font-size:11px;text-transform:uppercase;color:#829ab1;letter-spacing:.5px;} .value{font-size:24px;font-weight:bold;margin-top:4px;color:#102a43;} .small{font-size:12px;color:#627d98;} table{border-collapse:collapse;width:100%;margin:10px 0 18px;} th{background:#edf2f7;color:#243b53;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.4px;} th,td{border:1px solid #d9e2ec;padding:7px 8px;font-size:12px;vertical-align:top;} .right{text-align:right;} .center{text-align:center;} .pill{display:inline-block;border-radius:10px;padding:3px 8px;font-size:11px;font-weight:bold;} .status-ok{color:#0b6b3a;font-weight:bold;} .status-watch{color:#9a5b00;font-weight:bold;} .status-bad{color:#b42318;font-weight:bold;} .callout{border-left:5px solid #486581;background:#f0f4f8;padding:12px 14px;margin:14px 0;font-size:13px;} .recommendation{border-left:5px solid #0f609b;background:#eef8ff;padding:12px 14px;margin:10px 0;font-size:13px;} .muted{color:#829ab1;} .footer{margin-top:28px;border-top:1px solid #d9e2ec;padding-top:10px;color:#829ab1;font-size:11px;} .pagebreak{page-break-before:always;}")
	FileWriteLine($hFile, "</style></head><body><div class='page'>")

	FileWriteLine($hFile, "<div class='header'><h1>Analytical Workdays Report</h1><div class='subtitle'>Year " & _HtmlAsciiEntityEncode($Year) & " &bull; generated on " & @YEAR & "/" & @MON & "/" & @MDAY & " at " & @HOUR & ":" & @MIN & "</div></div>")

	FileWriteLine($hFile, "<table class='cards'><tr>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Overall status</div><div class='value " & $StatusClass & "'>" & _HtmlAsciiEntityEncode($StatusText) & "</div><div class='small'>Based on a 3 on-site days per 5 workdays target.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Tracked workdays</div><div class='value'>" & $WorkDays & "</div><div class='small'>" & $TotalDays & " total tracked calendar entries.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Expected on-site</div><div class='value'>" & $ExpectedTotal & "</div><div class='small'>Calculated from tracked workdays.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Actual on-site</div><div class='value'>" & $RealOnSite & "</div><div class='small'>On-Site + Travel days.</div></td>")
	FileWriteLine($hFile, "<td class='card'><div class='label'>Gap</div><div class='value " & $StatusClass & "'>" & _WD_ReportSignedNumber($GapTotal) & "</div><div class='small'>" & $CompliancePct & "% of expected coverage.</div></td>")
	FileWriteLine($hFile, "</tr></table>")

	FileWriteLine($hFile, "<h2>Executive summary</h2>")
	FileWriteLine($hFile, "<div class='callout'>" & _WD_ReportExecutiveSummary($GapTotal, $ExpectedTotal, $RealOnSite, $CompliancePct, $WorstQuarter, $WorstGap, $NotesCount, $UnknownCount) & "</div>")

	FileWriteLine($hFile, "<h2>Quarterly compliance dashboard</h2>")
	FileWriteLine($hFile, "<table><tr><th>Quarter</th><th class='right'>Tracked days</th><th class='right'>Workdays</th><th class='right'>Expected on-site</th><th class='right'>Actual on-site</th><th class='right'>Gap</th><th class='right'>Compliance</th><th>Status</th></tr>")
	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop
		Local $qClass = _WD_ReportStatusClass($QuarterStats[$q][4])
		FileWriteLine($hFile, "<tr><td><b>Q" & ($q + 1) & "</b></td><td class='right'>" & $QuarterStats[$q][0] & "</td><td class='right'>" & $QuarterStats[$q][1] & "</td><td class='right'>" & $QuarterStats[$q][3] & "</td><td class='right'>" & $QuarterStats[$q][2] & "</td><td class='right " & $qClass & "'>" & _WD_ReportSignedNumber($QuarterStats[$q][4]) & "</td><td class='right'>" & $QuarterStats[$q][5] & "%</td><td class='" & $qClass & "'>" & _HtmlAsciiEntityEncode(_WD_ReportStatusText($QuarterStats[$q][4])) & "</td></tr>")
	Next
	FileWriteLine($hFile, "</table>")

	FileWriteLine($hFile, "<h2>Monthly operating view</h2>")
	FileWriteLine($hFile, "<table><tr><th>Month</th><th class='right'>Workdays</th><th class='right'>Expected on-site</th><th class='right'>Actual on-site</th><th class='right'>Gap</th><th class='right'>Remote</th><th class='right'>Travel</th><th class='right'>PTO</th><th class='right'>Holiday</th><th class='right'>Sick</th><th class='right'>Notes</th></tr>")
	For $m = 1 To 12
		If $MonthTotalDays[$m] = 0 Then ContinueLoop
		Local $mExpected = Ceiling(($MonthWorkDays[$m] / 5) * 3)
		Local $mGap = $MonthOnSite[$m] - $mExpected
		Local $mClass = _WD_ReportStatusClass($mGap)
		FileWriteLine($hFile, "<tr><td><b>" & _HtmlAsciiEntityEncode(_WD_ReportMonthName($m)) & "</b></td><td class='right'>" & $MonthWorkDays[$m] & "</td><td class='right'>" & $mExpected & "</td><td class='right'>" & $MonthOnSite[$m] & "</td><td class='right " & $mClass & "'>" & _WD_ReportSignedNumber($mGap) & "</td><td class='right'>" & $MonthCategoryCount[$m][1] & "</td><td class='right'>" & $MonthCategoryCount[$m][4] & "</td><td class='right'>" & $MonthCategoryCount[$m][3] & "</td><td class='right'>" & $MonthCategoryCount[$m][2] & "</td><td class='right'>" & $MonthCategoryCount[$m][5] & "</td><td class='right'>" & $MonthNotesCount[$m] & "</td></tr>")
	Next
	FileWriteLine($hFile, "</table>")

	FileWriteLine($hFile, "<h2>Category mix</h2>")
	FileWriteLine($hFile, "<table><tr><th>Category</th><th class='right'>Days</th><th class='right'>Share of tracked days</th><th>Interpretation</th></tr>")
	For $c = 0 To 8
		If $CategoryTotal[$c] = 0 Then ContinueLoop
		Local $Share = 0
		If $TotalDays > 0 Then $Share = Round(($CategoryTotal[$c] / $TotalDays) * 100, 1)
		FileWriteLine($hFile, "<tr><td><span class='pill' style='background:" & $Colors[$c] & ";color:" & $FontColors[$c] & ";'>" & _HtmlAsciiEntityEncode($CatNames[$c]) & "</span></td><td class='right'>" & $CategoryTotal[$c] & "</td><td class='right'>" & $Share & "%</td><td>" & _HtmlAsciiEntityEncode(_WD_ReportCategoryInsight($c, $CategoryTotal[$c])) & "</td></tr>")
	Next
	FileWriteLine($hFile, "</table>")

	FileWriteLine($hFile, "<h2>Recommended actions</h2>")
	FileWriteLine($hFile, _WD_ReportRecommendationsHtml($GapTotal, $WorstQuarter, $WorstGap, $NotesCount, $UnknownCount, $RemoteDays, $TravelDays, $NonWorkingDays))

	FileWriteLine($hFile, "<h2>Notes and exceptions log</h2>")
	If $sNotesRows <> "" Then
		FileWriteLine($hFile, "<table><tr><th>Date</th><th>Category</th><th>Note / Marker</th></tr>" & $sNotesRows & "</table>")
	Else
		FileWriteLine($hFile, "<div class='callout'>No notes or markers were found for this year. For auditability, consider adding short notes to PTO, Travel, Sick, Holiday, and exception days.</div>")
	EndIf

	FileWriteLine($hFile, "<div class='footer'>Generated by Work Days &bull; Developed by Fabricio Zambroni &bull; Version: " & _HtmlAsciiEntityEncode(FileGetVersion(@ScriptFullPath)) & "<br>Disclaimer: this report uses the data currently stored in the WorkDays registry database and applies the same reference rule used by the app: expected on-site days = ceiling(workdays / 5 * 3).</div>")
	FileWriteLine($hFile, "</div></body></html>")
	FileClose($hFile)

;~ Local $sHtmlPath = @ScriptDir & "\Reports\" & $OutputPathTemp
;~ 	Local $sPdfPath = @ScriptDir & "\Reports\" & $OutputPath


;~ 	$oObject.Input = $OutputPathTemp
;~ 	$oObject.Output = $OutputPath

	$oObject.Input = $sHtmlPath
	$oObject.Output = $sPdfPath
	$oObject.Convert()

	If Not FileExists($sPdfPath) Then
		MsgBox(16, "Error", "Failed to convert Analytical HTML report to PDF.")
		Return SetError(2, 0, 0)
	EndIf

	MsgBox(262208, "Analytical Report", "The analytical report file was saved on: " & @CRLF & $sPdfPath)
	FileDelete($sHtmlPath)
	ShellExecute($sPdfPath)

	Return 1
EndFunc   ;==>GenerateWorkdaysProfessionalReportHTML

Func _WD_ReportReadDayValue($sRegistryBase, $iMonth, $iDay, ByRef $sDayName)
	Local $aMonthKeys[2] = [StringFormat("%02d", $iMonth), String($iMonth)]
	Local $aDayKeys[2] = [String($iDay), StringFormat("%02d", $iDay)]

	For $iMonthKey = 0 To 1
		For $iDayKey = 0 To 1
			Local $sFullKey = $sRegistryBase & "\" & $aMonthKeys[$iMonthKey]
			Local $sValue = RegRead($sFullKey, $aDayKeys[$iDayKey])
			If Not @error Then
				$sDayName = $aDayKeys[$iDayKey]
				Return $sValue
			EndIf
		Next
	Next

	Return SetError(1, 0, "")
EndFunc   ;==>_WD_ReportReadDayValue

Func _WD_ReportDaysInMonth($iYear, $iMonth)
	Switch Number($iMonth)
		Case 1, 3, 5, 7, 8, 10, 12
			Return 31
		Case 4, 6, 9, 11
			Return 30
		Case 2
			If Mod($iYear, 400) = 0 Or (Mod($iYear, 4) = 0 And Mod($iYear, 100) <> 0) Then Return 29
			Return 28
	EndSwitch

	Return 31
EndFunc   ;==>_WD_ReportDaysInMonth

Func _WD_ReportCategoryIndex($sRawVal)
	If $sRawVal = "" Then Return 7

	Local $sLetter = StringUpper(StringLeft($sRawVal, 1))
	Switch $sLetter
		Case "O"
			Return 0
		Case "R"
			Return 1
		Case "H"
			Return 2
		Case "P"
			Return 3
		Case "T"
			Return 4
		Case "S"
			Return 5
		Case "B"
			Return 7
		Case "W"
			Return 8
	EndSwitch

	Return 6
EndFunc   ;==>_WD_ReportCategoryIndex

Func _WD_ReportIsWorkDay($sCatLetter)
	If $sCatLetter = "O" Or $sCatLetter = "R" Or $sCatLetter = "T" Or $sCatLetter = "B" Or $sCatLetter = "" Then Return True
	Return False
EndFunc   ;==>_WD_ReportIsWorkDay

Func _WD_ReportSignedNumber($iValue)
	If Number($iValue) > 0 Then Return "+" & $iValue
	Return "" & $iValue
EndFunc   ;==>_WD_ReportSignedNumber

Func _WD_ReportStatusText($iGap)
	If Number($iGap) >= 0 Then Return "On track"
	If Number($iGap) >= -2 Then Return "Watch"
	Return "Behind"
EndFunc   ;==>_WD_ReportStatusText

Func _WD_ReportStatusClass($iGap)
	If Number($iGap) >= 0 Then Return "status-ok"
	If Number($iGap) >= -2 Then Return "status-watch"
	Return "status-bad"
EndFunc   ;==>_WD_ReportStatusClass

Func _WD_ReportMonthName($iMonth)
	Local $MonthNames[13] = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
	If Number($iMonth) < 1 Or Number($iMonth) > 12 Then Return ""
	Return $MonthNames[$iMonth]
EndFunc   ;==>_WD_ReportMonthName

Func _WD_ReportNoteToHtml($sNote)
	Local $sClean = StringReplace($sNote, @CRLF, "/n")
	$sClean = StringReplace($sClean, @LF, "/n")
	$sClean = StringReplace($sClean, @CR, "/n")
	Return StringReplace(_HtmlAsciiEntityEncode($sClean), "/n", "<br>")
EndFunc   ;==>_WD_ReportNoteToHtml

Func _WD_ReportCategoryInsight($iCategory, $iCount)
	Switch Number($iCategory)
		Case 0
			Return "Primary in-office presence category."
		Case 1
			Return "Remote workdays; useful to compare against the on-site target."
		Case 2
			Return "Holiday entries; these reduce practical availability."
		Case 3
			Return "PTO entries; keep notes when PTO affects quarterly expectations."
		Case 4
			Return "Travel is counted as on-site presence in the report logic."
		Case 5
			Return "Sick days; useful for exception tracking and context."
		Case 6
			Return "Unrecognized or custom entries. Review these for data quality."
		Case 7
			Return "Blank workdays; counted as available workdays by the report logic."
		Case 8
			Return "Weekend entries; excluded from the on-site expectation."
	EndSwitch
	Return ""
EndFunc   ;==>_WD_ReportCategoryInsight

Func _WD_ReportExecutiveSummary($iGapTotal, $iExpectedTotal, $iRealOnSite, $iCompliancePct, $iWorstQuarter, $iWorstGap, $iNotesCount, $iUnknownCount)
	Local $sText = ""
	If $iExpectedTotal = 0 Then
		$sText = "There are not enough tracked workdays to calculate an on-site expectation for this year."
	ElseIf $iGapTotal >= 0 Then
		$sText = "The year is currently meeting the on-site reference target. Actual on-site coverage is <b>" & $iRealOnSite & "</b> days against an expected <b>" & $iExpectedTotal & "</b> days, with a gap of <b>" & _WD_ReportSignedNumber($iGapTotal) & "</b>."
	Else
		$sText = "The year is currently below the on-site reference target. Actual on-site coverage is <b>" & $iRealOnSite & "</b> days against an expected <b>" & $iExpectedTotal & "</b> days, leaving <b>" & Abs($iGapTotal) & "</b> additional on-site or travel days to recover."
	EndIf

	If $iWorstQuarter > 0 Then
		$sText &= " The quarter requiring the most attention is <b>Q" & $iWorstQuarter & "</b> with a gap of <b>" & _WD_ReportSignedNumber($iWorstGap) & "</b>."
	EndIf

	$sText &= " Compliance coverage is currently <b>" & $iCompliancePct & "%</b>. The report also found <b>" & $iNotesCount & "</b> notes/markers."
	If $iUnknownCount > 0 Then $sText &= " <b>" & $iUnknownCount & "</b> entries are classified as Other and should be reviewed."
	Return $sText
EndFunc   ;==>_WD_ReportExecutiveSummary

Func _WD_ReportRecommendationsHtml($iGapTotal, $iWorstQuarter, $iWorstGap, $iNotesCount, $iUnknownCount, $iRemoteDays, $iTravelDays, $iNonWorkingDays)
	Local $sHtml = ""

	If $iGapTotal < 0 Then
		$sHtml &= "<div class='recommendation'><b>Recover the on-site gap:</b> plan at least " & Abs($iGapTotal) & " additional On-Site or Travel day(s), or validate whether specific PTO/Holiday/Sick exceptions should be excluded from your internal target calculation.</div>"
	Else
		$sHtml &= "<div class='recommendation'><b>Maintain the trend:</b> the yearly target is currently covered. Keep checking the quarterly view so the final months do not quietly erode the surplus.</div>"
	EndIf

	If $iWorstQuarter > 0 And $iWorstGap < 0 Then
		$sHtml &= "<div class='recommendation'><b>Focus Q" & $iWorstQuarter & ":</b> this quarter has the largest shortfall. Review Remote, PTO, Holiday, Sick, and Travel distribution there before using the report for decision-making.</div>"
	EndIf

	If $iUnknownCount > 0 Then
		$sHtml &= "<div class='recommendation'><b>Clean up data quality:</b> review " & $iUnknownCount & " Other entry/entries. An analytical report is only as good as the classification behind it.</div>"
	EndIf

	If $iNotesCount = 0 Then
		$sHtml &= "<div class='recommendation'><b>Add context:</b> no notes were captured. Add short markers to exception days so future reviews explain the reason, not only the category.</div>"
	Else
		$sHtml &= "<div class='recommendation'><b>Use the notes log:</b> " & $iNotesCount & " note(s) were captured. They are included at the end of this report to support review, audit, or personal planning.</div>"
	EndIf

	If $iRemoteDays > 0 And $iTravelDays > 0 Then
		$sHtml &= "<div class='recommendation'><b>Separate planned presence from mobility:</b> Travel is counted as on-site coverage here. If your policy treats travel differently, review those entries separately.</div>"
	EndIf

	If $iNonWorkingDays > 0 Then
		$sHtml &= "<div class='recommendation'><b>Check availability assumptions:</b> the report includes " & $iNonWorkingDays & " non-working entry/entries. Make sure PTO, Holiday, Sick, and Weekend days are aligned with how you want the target calculated.</div>"
	EndIf

	Return $sHtml
EndFunc   ;==>_WD_ReportRecommendationsHtml

