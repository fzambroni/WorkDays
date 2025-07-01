Func GenerateWorkdaysReportHTML($aData, $OutputPath)
    Local $CategoryTotals[0][2], $TotalDays = 0, $WorkDays = 0
    Local $html = "<html><head><title>Workdays Report</title></head><body style='font-family:Arial;'>"
    $html &= "<h2>Consolidated Summary for the Year</h2><table border='1' cellspacing='0' cellpadding='5'><tr><th>Category</th><th>Total Days</th></tr>"

    ; Calculate total days per category
    For $i = 0 To UBound($aData) - 1
        Local $date = $aData[$i][0]
        Local $category = $aData[$i][1]
        Local $note = $aData[$i][2]

        ; Normalize category
        Switch $category
            Case "O", "T"
                $WorkDays += 1
                _UpdateCategory("Real On-Site", 1, $CategoryTotals)
            Case "R"
                $WorkDays += 1
                _UpdateCategory("Remote", 1, $CategoryTotals)
            Case "V"
                _UpdateCategory("Vacation", 1, $CategoryTotals)
            Case "S"
                _UpdateCategory("Sick Leave", 1, $CategoryTotals)
            Case "H"
                _UpdateCategory("Holiday", 1, $CategoryTotals)
            Case "B", "", "W"
                If $note <> "" Then
                    $WorkDays += 1
                    _UpdateCategory("Blank or Weekends", 1, $CategoryTotals)
                EndIf
        EndSwitch
        $TotalDays += 1
    Next

    ; Append totals to summary table
    For $i = 0 To UBound($CategoryTotals) - 1
        $html &= "<tr><td>" & $CategoryTotals[$i][0] & "</td><td>" & $CategoryTotals[$i][1] & "</td></tr>"
    Next
    $html &= "<tr><td><b>Total Days</b></td><td><b>" & Ceiling($TotalDays) & "</b></td></tr>"
    $html &= "<tr><td><b>Work Days</b></td><td><b>" & Ceiling($WorkDays) & "</b></td></tr>"
    $html &= "<tr><td><b>Ratio</b></td><td><b>" & StringFormat("%.2f", $WorkDays / $TotalDays) & "</b></td></tr>"
    $html &= "</table><hr>"

    ; Detailed report by date
    $html &= "<h2>Detailed Report</h2>"
    $html &= "<table border='1' cellspacing='0' cellpadding='5'><tr><th>Date</th><th>Category</th><th>Note</th></tr>"
    For $i = 0 To UBound($aData) - 1
        Local $category = $aData[$i][1]
        If $category = "B" Or $category = "" Or $category = "W" Then
            If $aData[$i][2] <> "" Then
                $html &= "<tr><td>" & $aData[$i][0] & "</td><td>Blank or Weekends</td><td>" & $aData[$i][2] & "</td></tr>"
            EndIf
        Else
            $html &= "<tr><td>" & $aData[$i][0] & "</td><td>" & _CategoryName($category) & "</td><td>" & $aData[$i][2] & "</td></tr>"
        EndIf
    Next
    $html &= "</table></body></html>"

    ; Write to file
    Local $file = FileOpen($OutputPath, 2)
    FileWrite($file, $html)
    FileClose($file)
    ShellExecute($OutputPath)
EndFunc

Func _UpdateCategory($name, $value, ByRef $array)
    For $i = 0 To UBound($array) - 1
        If $array[$i][0] = $name Then
            $array[$i][1] += $value
            Return
        EndIf
    Next
    ReDim $array[UBound($array) + 1][2]
    $array[UBound($array) - 1][0] = $name
    $array[UBound($array) - 1][1] = $value
EndFunc

Func _CategoryName($c)
    Switch $c
        Case "O", "T"
            Return "Real On-Site"
        Case "R"
            Return "Remote"
        Case "V"
            Return "Vacation"
        Case "S"
            Return "Sick Leave"
        Case "H"
            Return "Holiday"
        Case "B", "", "W"
            Return "Blank or Weekends"
        Case Else
            Return "Unknown"
    EndSwitch
EndFunc
;~ Ceiling
;~ Func Ceiling($num)
;~     Return Int($num) + (Mod($num, 1) > 0)
;~ EndFunc