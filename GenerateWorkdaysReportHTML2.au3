Func GenerateWorkdaysReportHTML($Year, $Full)
	; Configs e arquivo de saída
	Local $RegistryBase = "HKEY_CURRENT_USER\Software\WorkDays\" & $Year
	Local $OutputPath = @ScriptDir & "\Workdays_Report.html"
	Local $hFile = FileOpen($OutputPath, 2)
	If $hFile = -1 Then
		MsgBox(16, "Error", "Failed to create HTML file.")
		Return
	EndIf

	; --------- ALTERAÇÕES AQUI ---------
	; Antes: 8 categorias (última era "Blank or Weekends")
	; Agora: 9 categorias, separando "Blank" (idx 7) e "Weekends" (idx 8)
	Local $CatNames[9]  = ["OnSite", "Remote", "Holiday", "PTO", "Travel", "Sick", "Other", "Blank", "Weekends"]
	Local $Colors[9]    = ["#b6fcd5", "#b3d9ff", "#fff5cc", "#ccffff", "#ffd9b3", "#ffcccc", "#dddddd", "#f5f5f5", "#eeeeee"]

	; Matrizes dimensionadas para 9 categorias
	Local $CategoryCount[4][9] = [[0]]
	Local $CategoryNotes[4][9]
	; QuarterStats: [totalDias, workDays, ratioInv, expected, realOnSite, remaining, reservado]
	Local $QuarterStats[4][7]
	; -----------------------------------

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

			Local $DateStr  = $Year & "/" & $MonthKey & "/" & $Day
			Local $CatLetter = StringUpper(StringLeft($RawVal, 1))
			Local $Note = StringTrimLeft($RawVal, 1)
			If $Note = $RawVal Then $Note = ""

			; Map de categoria
			Local $CatIndex = 6 ; Other por padrão
			If $CatLetter = "O" Then $CatIndex = 0
			If $CatLetter = "R" Then $CatIndex = 1
			If $CatLetter = "H" Then $CatIndex = 2
			If $CatLetter = "P" Then $CatIndex = 3
			If $CatLetter = "T" Then $CatIndex = 4
			If $CatLetter = "S" Then $CatIndex = 5

			; >>> ALTERAÇÃO: separar Blank e Weekends <<<
			; "Blank": letra 'B' ou valor vazio
			If $CatLetter = "B" Or $RawVal = "" Then $CatIndex = 7
			; "Weekends": letra 'W'
			If $CatLetter = "W" Then $CatIndex = 8
			; --------------------------------------------

			; Ignora "Other" sem nota (mesmo comportamento original)
			If $CatIndex = 6 And $Note = "" Then
				$i += 1
				ContinueLoop
			EndIf

			$CategoryCount[$q][$CatIndex] += 1

			; Notas / lista de datas
			If $Note <> "" Then
				If StringInStr($Note, "/n", 0, 1) Then
					$Note_Splited = StringSplit($Note, "/n", 1)
					For $Count_Note = 1 To $Note_Splited[0]
						If $Count_Note = 1 Then
							If $Note_Splited[$Count_Note] <> "" Then
								$CategoryNotes[$q][$CatIndex] &= "<li><b>" & $DateStr & ":</b> " & $Note_Splited[$Count_Note] & "</li>"
							EndIf
						Else
							If $Note_Splited[$Count_Note] <> "" Then
								$CategoryNotes[$q][$CatIndex] &= "<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> " & $Note_Splited[$Count_Note] & "<br>"
							EndIf
						EndIf
					Next
				Else
					$CategoryNotes[$q][$CatIndex] &= "<li><b>" & $DateStr & ":</b> " & $Note & "</li>"
				EndIf
			Else
				; Antes: não listava datas para "Blank or Weekends" (idx 7)
				; Agora: não lista para "Blank" (7) nem "Weekends" (8)
				If $CatIndex <> 7 And $CatIndex <> 8 Then
					$CategoryNotes[$q][$CatIndex] &= "<li>" & $DateStr & "</li>"
				EndIf
			EndIf

			$QuarterStats[$q][0] += 1  ; total de dias

			; Cálculo de WorkDays (não contar Weekends)
			; Original contava O, R, T, B e vazio. Mantido. 'W' NÃO entra.
			If $CatLetter = "O" Or $CatLetter = "R" Or $CatLetter = "T" Or $CatLetter = "B" Or $CatLetter = "" Then
				$QuarterStats[$q][1] += 1
				$WorkDays += 1
			EndIf

			; On-site real (O + T)
			If $CatLetter = "O" Or $CatLetter = "T" Then
				$QuarterStats[$q][4] += 1
				$RealOnSite += 1
				$TotalOnSiteTravel += 1
			EndIf

			$TotalDays += 1
			$i += 1
		WEnd
	Next

	; Estatísticas por trimestre
	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop
		Local $Expected = Ceiling(($QuarterStats[$q][1] / 5) * 3)
		Local $Actual = $QuarterStats[$q][4]
		$QuarterStats[$q][3] = $Expected
		$QuarterStats[$q][5] = $Expected - $Actual
		$QuarterStats[$q][2] = ($Actual > 0) ? Round(($Expected / $Actual), 2) : 0
	Next

	Local $ExpectedTotal = Ceiling(($WorkDays / 5) * 3)
	Local $Ratio = Round($RealOnSite / ($WorkDays / 5), 2)

	; Cabeçalho HTML
	If $Full = 1 Then
		FileWriteLine($hFile, "<html><head><title>Workdays Report - DETAILED - " & $Year & "</title>")
	Else
		FileWriteLine($hFile, "<html><head><title>Workdays Report - SIMPLE - " & $Year & "</title>")
	EndIf
	FileWriteLine($hFile, "<style>body{font-family:Arial;} table{border-collapse:collapse;width:100%;margin-bottom:20px;} th,td{border:1px solid #ccc;padding:6px;} th{background:#f0f0f0;} .stat,.qstat{margin:10px 0;padding:10px;background:#eef;border-left:4px solid #88f;} ul{margin:0;padding-left:20px;} h2{margin-top:30px;}</style></head><body>")

	If $Full = 1 Then
		FileWriteLine($hFile, "<h1>Workdays Report - DETAILED - " & $Year & "</h1>")
	Else
		FileWriteLine($hFile, "<h1>Workdays Report - SIMPLE - " & $Year & "</h1>")
	EndIf

	FileWriteLine($hFile, "<div class='stat'><b>Total Days Recorded:</b> " & $TotalDays & "<br><b>Work Days:</b> " & $WorkDays & "<br><b>Ratio*:</b> " & $Ratio & "<br><b>Estimated OnSite*:</b> " & $ExpectedTotal & "<br><b>Real On-Site*:</b> " & $RealOnSite & "<br><b>Remaining*:</b> " & ($ExpectedTotal - $RealOnSite) & "<br>*These values are for reference only. For an accurate analysis, consider the quarterly data. </div>")

	; Tabelas por trimestre
	For $q = 0 To 3
		If $QuarterStats[$q][0] = 0 Then ContinueLoop
		FileWriteLine($hFile, "<h2>Quarter " & ($q + 1) & "</h2>")
		FileWriteLine($hFile, "<div class='qstat'><b>Total Days:</b> " & $QuarterStats[$q][0] & "<br><b>Work Days:</b> " & $QuarterStats[$q][1] & "<br><b>Ratio:</b> " & Round($QuarterStats[$q][4] / ($QuarterStats[$q][1] / 5), 2) & "<br><b>Estimated OnSite:</b> " & $QuarterStats[$q][3] & "<br><b>Real On-Site:</b> " & $QuarterStats[$q][4] & "<br><b>Remaining:</b> " & $QuarterStats[$q][5] & "</div>")
		If $Full = 1 Then
			FileWriteLine($hFile, "<table><tr><th>Category</th><th>Count</th><th>Dates & Notes</th></tr>")
		Else
			FileWriteLine($hFile, "<table><tr><th>Category</th><th>Count</th></tr>")
		EndIf

		; ---- Loop agora vai até 8 (9 categorias) ----
		For $c = 0 To 8
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

	; --------- RESUMO ANUAL (atualizado p/ 9 categorias) ---------
	Local $YearlyTotals[9] = [0,0,0,0,0,0,0,0,0]
	For $q = 0 To 3
		For $c = 0 To 8
			$YearlyTotals[$c] += $CategoryCount[$q][$c]
		Next
	Next

	FileWriteLine($hFile, "<h2>Yearly Summary</h2>")
	FileWriteLine($hFile, "<table><tr><th>Category</th><th>Total Count</th></tr>")
	For $c = 0 To 8
		If $YearlyTotals[$c] > 0 Then
			FileWriteLine($hFile, "<tr style='background-color:" & $Colors[$c] & ";'><td><b>" & $CatNames[$c] & "</b></td><td>" & $YearlyTotals[$c] & "</td></tr>")
		EndIf
	Next
	FileWriteLine($hFile, "</table>")
	; --------------------------------------------------------------

	FileWriteLine($hFile, "<p style='color:gray;font-size:small;'>Generated on " & @YEAR & "/" & @MON & "/" & @MDAY & " at " & @HOUR & ":" & @MIN & "</p>")
	FileWriteLine($hFile, "<p style='color:gray;font-size:small;'>Develop by Fabricio Zambroni - Version: " & FileGetVersion(@ScriptFullPath) & "</p>")
	FileWriteLine($hFile, "</body></html>")
	FileClose($hFile)
	ShellExecute($OutputPath)
EndFunc   ;==>GenerateWorkdaysReportHTML
