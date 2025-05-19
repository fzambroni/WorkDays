
Func GenerateWorkdaysReportHTML($Year)
	$Full = 0

	If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
	$iMsgBoxAnswer = MsgBox(262180, "Generate Report", "Would you like to generate a full report that includes all recorded dates?" & @CRLF & @CRLF & "Yes = Full Report" & @CRLF & "No = Summary Only")
	Select
		Case $iMsgBoxAnswer = 6 ;Yes
			$Full = 1

		Case $iMsgBoxAnswer = 7 ;No
			$Full = 0

	EndSelect



	Local $RegistryBase = "HKEY_CURRENT_USER\Software\WorkDays\" & $Year
	Local $OutputPath = @ScriptDir & "\Workdays_Report.html"
	Local $hFile = FileOpen($OutputPath, 2)
	If $hFile = -1 Then
		MsgBox(16, "Error", "Failed to create HTML file.")
		Return
	EndIf

	Local $CatNames[8] = ["OnSite", "Remote", "Holiday", "PTO", "Travel", "Sick", "Other", "Blank or Weekends"]
	Local $Colors[8] = ["#b6fcd5", "#b3d9ff", "#fff5cc", "#ccffff", "#ffd9b3", "#ffcccc", "#dddddd", "#f5f5f5"]

	Local $CategoryCount[4][8] = [[0]]
	Local $CategoryNotes[4][8]
	Local $QuarterStats[4][7] ; [q][0=TotalDays, 1=WorkDays, 2=Ratio, 3=EstimatedOnSite, 4=RealOnSite, 5=RemainingOnSite, 6=--not used--]

	Local $TotalDays = 0, $WorkDays = 0, $RealOnSite = 0
	Local $TotalOnSiteTravel = 0

	For $m = 1 To 12
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

			If $RawVal = "" Or $CatLetter = "B" Or $CatLetter = "W" Then
				$CategoryCount[$q][7] += 1
				If $Note <> "" Then $CategoryNotes[$q][7] &= "<li><b>" & $DateStr & ":</b> " & $Note & "</li>"
				$QuarterStats[$q][0] += 1
				If $RawVal = "" Or $CatLetter = "B" Then
					$QuarterStats[$q][1] += 1
					$WorkDays += 1
				EndIf
				$TotalDays += 1
				$i += 1
				ContinueLoop
			EndIf

			Local $CatIndex = 6
			If $CatLetter = "O" Then $CatIndex = 0
			If $CatLetter = "R" Then $CatIndex = 1
			If $CatLetter = "H" Then $CatIndex = 2
			If $CatLetter = "P" Then $CatIndex = 3
			If $CatLetter = "T" Then $CatIndex = 4
			If $CatLetter = "S" Then $CatIndex = 5

			If $CatIndex = 6 And $Note = "" Then
				$i += 1
				ContinueLoop
			EndIf

			$CategoryCount[$q][$CatIndex] += 1
			If $Note <> "" Then
				$CategoryNotes[$q][$CatIndex] &= "<li><b>" & $DateStr & ":</b> " & $Note & "</li>"
			Else
				$CategoryNotes[$q][$CatIndex] &= "<li>" & $DateStr & "</li>"
			EndIf

			$QuarterStats[$q][0] += 1
			If $CatLetter = "O" Or $CatLetter = "R" Or $CatLetter = "T" Then
				$QuarterStats[$q][1] += 1
				$WorkDays += 1
			EndIf
			If $CatLetter = "O" Or $CatLetter = "T" Then
				$QuarterStats[$q][4] += 1
				$RealOnSite += 1
				$TotalOnSiteTravel += 1
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
		$QuarterStats[$q][2] = ($Actual > 0) ? Round(($Expected / $Actual), 2) : 0
	Next

	Local $ExpectedTotal = Ceiling(($WorkDays / 5) * 3)
	Local $Ratio = Round($RealOnSite / Ceiling($WorkDays / 5), 2)

	FileWriteLine($hFile, "<html><head><title>Workdays Report " & $Year & "</title>")
	FileWriteLine($hFile, "<style>body{font-family:Arial;} table{border-collapse:collapse;width:100%;margin-bottom:20px;} th,td{border:1px solid #ccc;padding:6px;} th{background:#f0f0f0;} .stat,.qstat{margin:10px 0;padding:10px;background:#eef;border-left:4px solid #88f;} ul{margin:0;padding-left:20px;} h2{margin-top:30px;}</style></head><body>")
	FileWriteLine($hFile, "<h1>Workdays Report - " & $Year & "</h1>")
	FileWriteLine($hFile, "<div class='stat'><b>Total Days Recorded:</b> " & $TotalDays & "<br><b>Work Days:</b> " & $WorkDays & "<br><b>Ratio*:</b> " & $Ratio & "<br><b>Estimated OnSite*:</b> " & $ExpectedTotal & "<br><b>Real On-Site*:</b> " & $RealOnSite & "<br><b>Remaining*:</b> " & ($ExpectedTotal - $RealOnSite) & "<br>*These values are for reference only. For an accurate analysis, consider the quarterly data. </div>")

	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop
		FileWriteLine($hFile, "<h2>Quarter " & ($q + 1) & "</h2>")
		FileWriteLine($hFile, "<div class='qstat'><b>Total Days:</b> " & $QuarterStats[$q][0] & "<br><b>Work Days:</b> " & $QuarterStats[$q][1] & "<br><b>Ratio:</b> " & Round($QuarterStats[$q][4] / Ceiling($QuarterStats[$q][1] / 5), 2) & "<br><b>Estimated OnSite:</b> " & $QuarterStats[$q][3] & "<br><b>Real On-Site:</b> " & $QuarterStats[$q][4] & "<br><b>Remaining:</b> " & $QuarterStats[$q][5] & "</div>")
		If $Full = 1 Then
		FileWriteLine($hFile, "<table><tr><th>Category</th><th>Count</th><th>Dates & Notes</th></tr>")
		Else
		FileWriteLine($hFile, "<table><tr><th>Category</th><th>Count</th></tr>")
		EndIf
		For $c = 0 To 7
			If $CategoryCount[$q][$c] = 0 Then ContinueLoop
			FileWriteLine($hFile, "<tr style='background-color:" & $Colors[$c] & ";'><td><b>" & $CatNames[$c] & "</b></td><td>" & $CategoryCount[$q][$c] & "</td>")
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

	FileWriteLine($hFile, "<p style='color:gray;font-size:small;'>Generated on " & @YEAR & "/" & @MON & "/" & @MDAY & " at " & @HOUR & ":" & @MIN & "</p>")
	FileWriteLine($hFile, "<p style='color:gray;font-size:small;'>Develop by Fabricio Zambroni - Version: " & FileGetVersion(@ScriptFullPath) & "</p>")
	FileWriteLine($hFile, "</body></html>")
	FileClose($hFile)
	ShellExecute($OutputPath)
EndFunc   ;==>GenerateWorkdaysReportHTML
