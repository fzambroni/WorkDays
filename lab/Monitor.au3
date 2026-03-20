; ==================================================================================================
; MonitorUDF_Examples.au3
; Interactive example tester for MonitorUDF.au3 UDF
; Fixed: Array index bug in FuncTest_10, improved error handling
; ==================================================================================================

#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <WindowsConstants.au3>
#include <GuiListBox.au3>
#include <GuiEdit.au3>
#include <Array.au3>
#include "Monitor_UDF.au3"  ; Updated include name

; ==================================================================================================
; Create GUI
; ==================================================================================================
Global $GUI_W = 860, $GUI_H = 600
Global $hGUI = GUICreate("MonitorUDF - Examples (by TRONG.PRO)", $GUI_W, $GUI_H, -1, -1)
GUISetBkColor(0xF5F5F5, $hGUI)

; Title
GUICtrlCreateLabel("MonitorUDF Example Launcher", 12, 10, 400, 24)
GUICtrlSetFont(-1, 12, 800, 0, 'Segoe UI', 5)

; Buttons column 1 - 6 buttons (1-6)
Local $x1 = 12, $y1 = 48, $bw = 260, $bh = 36, $gap = 8
Global $iBtn1 = GUICtrlCreateButton("1. Enumerate monitors", $x1, $y1 + ($bh + $gap) * 0, $bw, $bh)
Global $iBtn2 = GUICtrlCreateButton("2. Move Notepad -> Monitor #2 (center)", $x1, $y1 + ($bh + $gap) * 1, $bw, $bh)
Global $iBtn3 = GUICtrlCreateButton("3. Move Notepad -> Monitor #2 @ (100,100)", $x1, $y1 + ($bh + $gap) * 2, $bw, $bh)
Global $iBtn4 = GUICtrlCreateButton("4. Which monitor is mouse on?", $x1, $y1 + ($bh + $gap) * 3, $bw, $bh)
Global $iBtn5 = GUICtrlCreateButton("5. Show virtual desktop bounds", $x1, $y1 + ($bh + $gap) * 4, $bw, $bh)
Global $iBtn6 = GUICtrlCreateButton("6. Convert coords (local <-> virtual)", $x1, $y1 + ($bh + $gap) * 5, $bw, $bh)

; Buttons column 2 - 6 buttons (7-12)
Local $x2 = $x1 + $bw + 12
Global $iBtn7 = GUICtrlCreateButton("7. Show monitor info (MsgBox)", $x2, $y1 + ($bh + $gap) * 0, $bw, $bh)
Global $iBtn8 = GUICtrlCreateButton("8. Check if Notepad is visible", $x2, $y1 + ($bh + $gap) * 1, $bw, $bh)
Global $iBtn9 = GUICtrlCreateButton("9. Create small GUI on each monitor", $x2, $y1 + ($bh + $gap) * 2, $bw, $bh)
Global $iBtn10 = GUICtrlCreateButton("10. Move all visible windows -> primary", $x2, $y1 + ($bh + $gap) * 3, $bw, $bh)
Global $iBtn11 = GUICtrlCreateButton("11. Refresh monitor list", $x2, $y1 + ($bh + $gap) * 4, $bw, $bh)
Global $iBtn12 = GUICtrlCreateButton("12. Check monitor connected", $x2, $y1 + ($bh + $gap) * 5, $bw, $bh)

; Buttons column 3 - 6 buttons (13-17, but layout for 6)
Local $x3 = $x2 + $bw + 12
Global $iBtn13 = GUICtrlCreateButton("13. Get DPI scaling", $x3, $y1 + ($bh + $gap) * 0, $bw, $bh)
Global $iBtn14 = GUICtrlCreateButton("14. Get display orientation", $x3, $y1 + ($bh + $gap) * 1, $bw, $bh)
Global $iBtn15 = GUICtrlCreateButton("15. Enumerate display modes", $x3, $y1 + ($bh + $gap) * 2, $bw, $bh)
Global $iBtn16 = GUICtrlCreateButton("16. Get monitor from rect", $x3, $y1 + ($bh + $gap) * 3, $bw, $bh)
Global $iBtn17 = GUICtrlCreateButton("17. Get monitor from window", $x3, $y1 + ($bh + $gap) * 4, $bw, $bh)

; Controls: log edit, clear, auto-demo, close
Local $logX = 12, $logY = $y1 + ($bh + $gap) * 6
Local $logW = $GUI_W - 24, $logH = 180
Global $idLog = GUICtrlCreateEdit("", $logX, $logY, $logW, $logH, BitOR($ES_READONLY, $WS_HSCROLL, $WS_VSCROLL, $ES_MULTILINE))
GUICtrlSetFont($idLog, 9)
Global $idFuncTest__ClearLog = GUICtrlCreateButton("Clear Log", 12, $logY + $logH + 10, 120, 28)
Global $idFuncTest__RunAllDemo = GUICtrlCreateButton("Auto Demo (Run 1..17)", 150, $logY + $logH + 10, 220, 28)
Global $idFuncTest__Close = GUICtrlCreateButton("Close", $GUI_W - 120, $logY + $logH + 10, 100, 28)
Global $pidNotepad = 0
GUISetState(@SW_SHOW)

; Keep a list of created GUIs for Example 10 so we can close them later
Global $Msg, $g_createdGUIs[0]  ; FIXED: Initialize as empty array
Global $g_bAutoMode = False  ; Track if running in auto demo mode
Global $g_hNotepadWindows[0]  ; Track all notepad windows created

; ==================================================================================================
; Main loop
; ==================================================================================================
While 1
    $Msg = GUIGetMsg()
    Switch $Msg
        Case $GUI_EVENT_CLOSE, $idFuncTest__Close
            ; close any GUIs created in example 10
            For $i = 0 To UBound($g_createdGUIs) - 1
                If IsHWnd($g_createdGUIs[$i]) Then GUIDelete($g_createdGUIs[$i])
            Next
            ExitLoop
        Case $idFuncTest__ClearLog
            GUICtrlSetData($idLog, '', '')
        Case $idFuncTest__RunAllDemo
            $g_bAutoMode = True
            ReDim $g_hNotepadWindows[0]
            FuncTest_1()
            Sleep(1000)
            FuncTest_2()
            Sleep(2000)
            FuncTest_3()
            Sleep(2000)
            FuncTest_4()
            Sleep(1500)
            FuncTest_5()
            Sleep(1000)
            FuncTest_6()
            Sleep(1000)
            FuncTest_7()
            Sleep(1000)
            FuncTest_8()
            Sleep(2000)
            FuncTest_9()
            Sleep(1000)
            FuncTest_10()
            Sleep(1000)
            FuncTest_11()
            Sleep(1000)
            FuncTest_12()
            Sleep(1000)
            FuncTest_13()
            Sleep(1000)
            FuncTest_14()
            Sleep(1000)
            FuncTest_15()
            Sleep(1000)
            FuncTest_16()
            Sleep(1000)
            FuncTest_17()
            Sleep(2000)
            ; Cleanup all created windows
            _CleanupAllWindows()
            $g_bAutoMode = False

        Case $iBtn1
            FuncTest_1()
        Case $iBtn2
            FuncTest_2()
        Case $iBtn3
            FuncTest_3()
        Case $iBtn4
            FuncTest_4()
        Case $iBtn5
            FuncTest_5()
        Case $iBtn6
            FuncTest_6()
        Case $iBtn7
            FuncTest_7()
        Case $iBtn8
            FuncTest_8()
        Case $iBtn9
            FuncTest_9()
        Case $iBtn10
            FuncTest_10()
        Case $iBtn11
            FuncTest_11()
        Case $iBtn12
            FuncTest_12()
        Case $iBtn13
            FuncTest_13()
        Case $iBtn14
            FuncTest_14()
        Case $iBtn15
            FuncTest_15()
        Case $iBtn16
            FuncTest_16()
        Case $iBtn17
            FuncTest_17()
    EndSwitch
    Sleep(5)
WEnd

GUIDelete()
Exit 0

; ==================================================================================================
; Helper: append to log (with timestamp)
; ==================================================================================================
Func _Log($s)
    ConsoleWrite($s & @CRLF)
    ; Format: DD/MM/YYYY HH:MM:SS
    Local $sDate = StringFormat("%02d/%02d/%04d", @MDAY, @MON, @YEAR)
    Local $sTime = StringFormat("%02d:%02d:%02d", @HOUR, @MIN, @SEC)
    Local $t = $sDate & " " & $sTime
    Local $cur = GUICtrlRead($idLog)
    If $cur = "" Then
        GUICtrlSetData($idLog, "[" & $t & "] " & $s)
    Else
        GUICtrlSetData($idLog, $cur & @CRLF & "[" & $t & "] " & $s)
    EndIf
    ; move caret to end
    _GUICtrlEdit_LineScroll($idLog, 0, _GUICtrlEdit_GetLineCount($idLog))
EndFunc   ;==>_Log

; ==================================================================================================
; Helper: Cleanup all created windows (Notepad and GUIs)
; ==================================================================================================
Func _CleanupAllWindows()
    ; Close all Notepad windows
    Local $aList = WinList("[CLASS:Notepad]")
    For $i = 1 To $aList[0][0]
        If $aList[$i][0] <> "" Then
            Local $hWnd = WinGetHandle($aList[$i][0])
            If $hWnd Then WinClose($hWnd)
        EndIf
    Next
    Sleep(300)
    ; Force close any remaining Notepad processes
    While ProcessExists("notepad.exe")
        ProcessClose("notepad.exe")
        Sleep(100)
    WEnd
    ; Close all created GUIs
    For $i = 0 To UBound($g_createdGUIs) - 1
        If IsHWnd($g_createdGUIs[$i]) Then GUIDelete($g_createdGUIs[$i])
    Next
    ReDim $g_createdGUIs[0]
    ReDim $g_hNotepadWindows[0]
EndFunc   ;==>_CleanupAllWindows

; ==================================================================================================
; Helper: Track Notepad window
; ==================================================================================================
Func _TrackNotepadWindow($hWnd)
    If $hWnd Then
        Local $n = UBound($g_hNotepadWindows)
        ReDim $g_hNotepadWindows[$n + 1]
        $g_hNotepadWindows[$n] = $hWnd
    EndIf
EndFunc   ;==>_TrackNotepadWindow

Func FuncTest_1()
    _Log('+ TEST 1: Enumerate monitors -----------------------\')
    ; Enumerate monitors
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If @error Then
        Local $sMsg = "TEST 1: FAILED" & @CRLF & "ERROR - Failed to enumerate monitors" & @CRLF & "@error=" & @error
        _Log("---> Example 1: ERROR - Failed to enumerate monitors")
        If Not $g_bAutoMode Then MsgBox(48, "Example 1", $sMsg, 3)
        Return
    EndIf
    _Log("---> Example 1: Monitors detected: " & $cnt)
    Local $sResults = "Total Monitors: " & $cnt & @CRLF & @CRLF
    For $i = 1 To $cnt
        Local $a = _Monitor_GetInfo($i)
        If @error Then
            _Log("  Monitor " & $i & ": ERROR getting info")
            $sResults &= "Monitor #" & $i & ": ERROR" & @CRLF
        Else
            Local $sPrimary = $a[9] ? " [PRIMARY]" : ""
            _Log("  Monitor " & $i & ": Device=" & $a[10] & " Bounds=(" & $a[1] & "," & $a[2] & ")-(" & $a[3] & "," & $a[4] & ") Work=(" & $a[5] & "," & $a[6] & ")-(" & $a[7] & "," & $a[8] & ") Primary=" & $a[9])
            $sResults &= "Monitor #" & $i & $sPrimary & ": " & $a[10] & @CRLF & _
                    "  Bounds: " & $a[1] & "," & $a[2] & " to " & $a[3] & "," & $a[4] & @CRLF & _
                    "  Work Area: " & $a[5] & "," & $a[6] & " to " & $a[7] & "," & $a[8] & @CRLF
        EndIf
    Next
    Local $sMsg = "TEST 1: SUCCESS" & @CRLF & @CRLF & $sResults
    If Not $g_bAutoMode Then MsgBox(64, "Example 1", $sMsg, 8)
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_1

Func FuncTest_2()
    _Log('+ TEST 2: Move Notepad to monitor #2 centered ------\')
    ; Move Notepad to monitor #2 centered
    $pidNotepad = Run("notepad.exe")
    If Not WinWaitActive("[CLASS:Notepad]", "", 5) Then
        _Log("---> Example 2: Notepad did not start / focus")
        Local $sMsg = "TEST 2: FAILED" & @CRLF & "Notepad did not start / focus"
        If Not $g_bAutoMode Then MsgBox(48, "Example 2", $sMsg, 3)
    Else
        Sleep(1000)
        Local $hWnd = WinGetHandle("[CLASS:Notepad]")
        _TrackNotepadWindow($hWnd)
        _Monitor_GetList()
        Local $cnt = _Monitor_GetCount()
        If $cnt < 2 Then
            Local $sMsg = "TEST 2: SKIPPED" & @CRLF & "Need at least 2 monitors" & @CRLF & "Current: " & $cnt & " monitor(s)"
            _Log("---> Example 2: Need at least 2 monitors")
            If Not $g_bAutoMode Then MsgBox(64, "Example 2", $sMsg, 3)
        Else
            Local $iResult = _Monitor_MoveWindowToScreen("[CLASS:Notepad]", "", 2, -1, -1, True)
            If @error Then
                Local $sMsg = "TEST 2: FAILED" & @CRLF & "ERROR moving window" & @CRLF & "@error=" & @error
                _Log("---> Example 2: ERROR moving window: @error=" & @error)
                If Not $g_bAutoMode Then MsgBox(48, "Example 2", $sMsg, 3)
            Else
                Local $sMsg = "TEST 2: SUCCESS" & @CRLF & "Notepad moved to Monitor #2" & @CRLF & "Position: Centered"
                _Log("---> Example 2: Notepad moved to monitor #2 (centered)")
                If Not $g_bAutoMode Then MsgBox(64, "Example 2", $sMsg, 3)
            EndIf
        EndIf
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_2

Func FuncTest_3()
    _Log('+ TEST 3: Move Notepad to monitor #2 at (100,100 ----\)')
    ; Move Notepad to monitor #2 at (100,100)
    $pidNotepad = Run("notepad.exe")
    If Not WinWaitActive("[CLASS:Notepad]", "", 5) Then
        Local $sMsg = "TEST 3: FAILED" & @CRLF & "Notepad did not start / focus"
        _Log("---> Example 3: Notepad did not start / focus")
        If Not $g_bAutoMode Then MsgBox(48, "Example 3", $sMsg, 3)
    Else
        Sleep(1000)
        Local $hWnd = WinGetHandle("[CLASS:Notepad]")
        _TrackNotepadWindow($hWnd)
        _Monitor_GetList()
        Local $cnt = _Monitor_GetCount()
        If $cnt < 2 Then
            Local $sMsg = "TEST 3: SKIPPED" & @CRLF & "Need at least 2 monitors" & @CRLF & "Current: " & $cnt & " monitor(s)"
            _Log("---> Example 3: Need at least 2 monitors")
            If Not $g_bAutoMode Then MsgBox(64, "Example 3", $sMsg, 3)
        Else
            Local $iResult = _Monitor_MoveWindowToScreen("[CLASS:Notepad]", "", 2, 100, 100, True)
            If @error Then
                Local $sMsg = "TEST 3: FAILED" & @CRLF & "ERROR moving window" & @CRLF & "@error=" & @error
                _Log("---> Example 3: ERROR moving window: @error=" & @error)
                If Not $g_bAutoMode Then MsgBox(48, "Example 3", $sMsg, 3)
            Else
                Local $sMsg = "TEST 3: SUCCESS" & @CRLF & "Notepad moved to Monitor #2" & @CRLF & "Position: (100, 100)"
                _Log("---> Example 3: Notepad moved to monitor #2 at (100,100)")
                If Not $g_bAutoMode Then MsgBox(64, "Example 3", $sMsg, 3)
            EndIf
        EndIf
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_3

Func FuncTest_4()
    _Log('+ TEST 4: Which monitor is mouse on ------------------\')
    ; Automatically move mouse to each monitor and test
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If $cnt < 1 Then
        Local $sMsg = "TEST 4: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 4: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 4", $sMsg, 3)
    Else
        Local $aResults = ""
        Local $sCurrentPos = MouseGetPos()
        Local $iOrigX = $sCurrentPos[0], $iOrigY = $sCurrentPos[1]

        _Log("---> Example 4: Testing " & $cnt & " monitor(s)")
        For $i = 1 To $cnt
            ; Get monitor center
            Local $iLeft, $iTop, $iRight, $iBottom
            _Monitor_GetBounds($i, $iLeft, $iTop, $iRight, $iBottom)
            Local $iCenterX = $iLeft + ($iRight - $iLeft) / 2
            Local $iCenterY = $iTop + ($iBottom - $iTop) / 2

            ; Move mouse to center of monitor
            MouseMove($iCenterX, $iCenterY, 0)
            Sleep(200)

            ; Check which monitor mouse is on
            Local $m = _Monitor_GetFromPoint()
            If @error Then
                _Log("  Monitor " & $i & ": ERROR - @error=" & @error)
                $aResults &= "Monitor " & $i & ": ERROR" & @CRLF
            Else
                Local $aInfo = _Monitor_GetInfo($i)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                Local $sStatus = ($m = $i) ? "CORRECT" : "WRONG (detected #" & $m & ")"
                _Log("  Monitor " & $i & " (" & $sDevice & "): Mouse detected on #" & $m & " - " & $sStatus)
                $aResults &= "Monitor #" & $i & " (" & $sDevice & "): #" & $m & " - " & $sStatus & @CRLF
            EndIf
            Sleep(300)
        Next

        ; Restore original mouse position
        MouseMove($iOrigX, $iOrigY, 0)

        Local $sMsg = "TEST 4: COMPLETE" & @CRLF & @CRLF & "Mouse Position Test Results:" & @CRLF & $aResults
        _Log("---> Example 4: Test complete, mouse restored to original position")
        If Not $g_bAutoMode Then MsgBox(64, "Example 4", $sMsg, 5)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_4

Func FuncTest_5()
    _Log('+ TEST 5: Virtual desktop bounds -----------------------\')
    ; Virtual desktop bounds
    Local $aV = _Monitor_GetVirtualBounds()
    If @error Then
        Local $sMsg = "TEST 5: FAILED" & @CRLF & "ERROR getting virtual bounds" & @CRLF & "@error=" & @error
        _Log("---> Example 5: ERROR getting virtual bounds: @error=" & @error)
        If Not $g_bAutoMode Then MsgBox(48, "Example 5", $sMsg, 3)
    Else
        Local $sMsg = "TEST 5: SUCCESS" & @CRLF & @CRLF & "Virtual Desktop Bounds:" & @CRLF & _
                "Left: " & $aV[0] & @CRLF & _
                "Top: " & $aV[1] & @CRLF & _
                "Width: " & $aV[2] & @CRLF & _
                "Height: " & $aV[3] & @CRLF & @CRLF & _
                "Right: " & ($aV[0] + $aV[2]) & @CRLF & _
                "Bottom: " & ($aV[1] + $aV[3])
        _Log("---> Example 5: Virtual bounds L=" & $aV[0] & " T=" & $aV[1] & " W=" & $aV[2] & " H=" & $aV[3])
        If Not $g_bAutoMode Then MsgBox(64, "Example 5", $sMsg, 5)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_5

Func FuncTest_6()
    _Log('+ TEST 6: Convert coords example (local -> virtual -> back) --\')
    ; Convert coords example (local -> virtual -> back)
    _Monitor_GetList()
    Local $mon = _Monitor_GetCount()
    If $mon < 1 Then
        Local $sMsg = "TEST 6: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 6: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 6", $sMsg, 3)
    Else
        Local $sResults = ""
        For $i = 1 To $mon
            Local $xLocal = 50, $yLocal = 100
            Local $aV = _Monitor_ToVirtual($i, $xLocal, $yLocal)
            If @error Then
                _Log("---> Example 6: Monitor " & $i & " ERROR converting to virtual: @error=" & @error)
                $sResults &= "Monitor #" & $i & ": ERROR (to virtual)" & @CRLF
            Else
                Local $aBack = _Monitor_FromVirtual($i, $aV[0], $aV[1])
                If @error Then
                    _Log("---> Example 6: Monitor " & $i & " ERROR converting from virtual: @error=" & @error)
                    $sResults &= "Monitor #" & $i & ": ERROR (from virtual)" & @CRLF
                Else
                    Local $bMatch = (Abs($aBack[0] - $xLocal) < 1) And (Abs($aBack[1] - $yLocal) < 1)
                    _Log("---> Example 6: Mon " & $i & " local(" & $xLocal & "," & $yLocal & ") -> virtual(" & $aV[0] & "," & $aV[1] & ") -> back(" & $aBack[0] & "," & $aBack[1] & ")")
                    $sResults &= "Monitor #" & $i & ": (" & $xLocal & "," & $yLocal & ") -> (" & $aV[0] & "," & $aV[1] & ") -> (" & $aBack[0] & "," & $aBack[1] & ")" & _
                            ($bMatch ? " ✓" : " ✗") & @CRLF
                EndIf
            EndIf
        Next
        Local $sMsg = "TEST 6: COMPLETE" & @CRLF & @CRLF & "Coordinate Conversion Test:" & @CRLF & $sResults
        If Not $g_bAutoMode Then MsgBox(64, "Example 6", $sMsg, 5)
    EndIf
    _Log('- End --------------------------------------------------------/')
EndFunc   ;==>FuncTest_6

Func FuncTest_7()
    _Log('+ TEST 7: Show detailed info via MsgBox (calls UDF) ------\')
    ; Show detailed info via MsgBox (calls UDF)
    _Monitor_GetList()
    Local $sResult = _Monitor_ShowInfo(1, 8)
    If @error Then
        Local $sMsg = "TEST 7: FAILED" & @CRLF & "ERROR showing info" & @CRLF & "@error=" & @error
        _Log("---> Example 7: ERROR showing info: @error=" & @error)
        If Not $g_bAutoMode Then MsgBox(48, "Example 7", $sMsg, 3)
    Else
        Local $sMsg = "TEST 7: SUCCESS" & @CRLF & @CRLF & "Detailed monitor information displayed above."
        _Log("---> Example 7: _Monitor_ShowInfo() called successfully")
        If Not $g_bAutoMode Then MsgBox(64, "Example 7", $sMsg, 3)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_7

Func FuncTest_8()
    _Log('+ TEST 8: Start notepad, check visible ---------------\')
    ; Start notepad, check visible
    $pidNotepad = Run("notepad.exe")
    If Not WinWaitActive("[CLASS:Notepad]", "", 5) Then
        Local $sMsg = "TEST 8: FAILED" & @CRLF & "Notepad did not start/focus"
        _Log("---> Example 8: Notepad did not start/focus")
        If Not $g_bAutoMode Then MsgBox(48, "Example 8", $sMsg, 3)
    Else
        Sleep(1000)
        Local $h = WinGetHandle("[CLASS:Notepad]")
        _TrackNotepadWindow($h)
        Local $b = _Monitor_IsVisibleWindow($h)
        If @error Then
            Local $sMsg = "TEST 8: FAILED" & @CRLF & "ERROR checking visibility" & @CRLF & "@error=" & @error
            _Log("---> Example 8: ERROR checking visibility: @error=" & @error)
            If Not $g_bAutoMode Then MsgBox(48, "Example 8", $sMsg, 3)
        Else
            Local $sTitle = WinGetTitle($h)
            Local $sMsg = "TEST 8: SUCCESS" & @CRLF & @CRLF & "Window: " & ($sTitle = "" ? "[No Title]" : $sTitle) & @CRLF & _
                    "Handle: " & $h & @CRLF & _
                    "Visible: " & ($b ? "YES ✓" : "NO ✗")
            _Log("---> Example 8: Notepad handle " & $h & " visible? " & ($b ? "Yes" : "No"))
            If Not $g_bAutoMode Then MsgBox(64, "Example 8", $sMsg, 3)
        EndIf
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_8

Func FuncTest_9()
    _Log('+ TEST 9: Create small GUI on each monitor -------------\')
    ; Create small GUI on each monitor
    _Monitor_GetList()
    ; close previously created
    For $i = 0 To UBound($g_createdGUIs) - 1
        If IsHWnd($g_createdGUIs[$i]) Then GUIDelete($g_createdGUIs[$i])
    Next
    ReDim $g_createdGUIs[0]  ; reset
    Local $created = 0
    For $i = 1 To _Monitor_GetCount()
        Local $a = _Monitor_GetInfo($i)
        If @error Then
            _Log("  Monitor " & $i & ": ERROR getting info")
            ContinueLoop
        EndIf
        Local $h = GUICreate("Monitor #" & $i & " - " & $a[10], 260, 120, $a[1] + 40, $a[2] + 40)
        GUICtrlCreateLabel("Monitor " & $i & ($a[9] ? " (Primary)" : ""), 10, 12, 240, 20)
        GUISetState(@SW_SHOW, $h)
        ; store to close later
        __ArrayAdd($g_createdGUIs, $h)
        $created += 1
    Next
    Local $sMsg = "TEST 9: SUCCESS" & @CRLF & @CRLF & "Created " & $created & " GUI window(s)" & @CRLF & @CRLF
    If $created > 0 Then
        $sMsg &= "One GUI created on each monitor:" & @CRLF
        For $i = 1 To _Monitor_GetCount()
            Local $a = _Monitor_GetInfo($i)
            If Not @error Then
                Local $sPrimary = $a[9] ? " [PRIMARY]" : ""
                $sMsg &= "  Monitor #" & $i & $sPrimary & ": " & $a[10] & @CRLF
            EndIf
        Next
        $sMsg &= @CRLF & "Windows will be closed when you close the launcher."
    Else
        $sMsg &= "No GUIs were created."
    EndIf
    _Log("---> Example 9: Created " & $created & " GUI(s) on monitors. Use Close to exit (they will be closed).")
    If Not $g_bAutoMode Then MsgBox(64, "Example 9", $sMsg, 5)
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_9

Func FuncTest_10()
    _Log('+ TEST 10: Move all visible windows to primary ----------\')
    ; Move all visible windows to primary
    _Monitor_GetList()
    Local $prim = _Monitor_GetPrimary()
    If $prim = 0 Or @error Then
        Local $sMsg = "TEST 10: FAILED" & @CRLF & "Primary monitor not found or error" & @CRLF & "@error=" & @error
        _Log("---> Example 10: Primary monitor not found or error")
        If Not $g_bAutoMode Then MsgBox(48, "Example 10", $sMsg, 3)
    Else
        Local $aList = WinList()
        Local $moved = 0
        Local $aInfo = _Monitor_GetInfo($prim)
        Local $sDevice = @error ? "N/A" : $aInfo[10]
        For $i = 1 To $aList[0][0]
            If $aList[$i][0] <> "" Then
                ; FIXED: Use correct array index - WinList()[i][0] is title, need to get handle from title
                Local $h = WinGetHandle($aList[$i][0])  ; FIXED: Changed from [1] to [0]
                If Not @error And $h Then
                    If _Monitor_IsVisibleWindow($h) Then
                        Local $iResult = _Monitor_MoveWindowToScreen($h, "", $prim)
                        If Not @error Then $moved += 1
                    EndIf
                EndIf
            EndIf
        Next
        Local $sMsg = "TEST 10: SUCCESS" & @CRLF & @CRLF & _
                "Moved " & $moved & " visible window(s)" & @CRLF & _
                "to Primary Monitor #" & $prim & @CRLF & _
                "Device: " & $sDevice
        _Log("---> Example 10: Moved " & $moved & " visible windows to primary monitor #" & $prim)
        If Not $g_bAutoMode Then MsgBox(64, "Example 10", $sMsg, 4)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_10

Func FuncTest_11()
    _Log('+ TEST 11: Refresh monitor list ----------------------\')
    ; Refresh monitor list
    Local $cntBefore = _Monitor_GetCount()
    _Log("---> Example 11: Monitors before refresh: " & $cntBefore)

    Local $cntAfter = _Monitor_Refresh()
    If @error Then
        Local $sMsg = "TEST 11: FAILED" & @CRLF & "ERROR refreshing monitor list" & @CRLF & "@error=" & @error
        _Log("---> Example 11: ERROR refreshing: @error=" & @error)
        If Not $g_bAutoMode Then MsgBox(48, "Example 11", $sMsg, 3)
    Else
        Local $sChangeInfo = ""
        If $cntBefore <> $cntAfter Then
            _Log("  --> Monitor count changed! (was " & $cntBefore & ", now " & $cntAfter & ")")
            $sChangeInfo = @CRLF & "⚠ CHANGE DETECTED! ⚠" & @CRLF & "Before: " & $cntBefore & " monitor(s)" & @CRLF & "After: " & $cntAfter & " monitor(s)"
        Else
            _Log("  --> Monitor count unchanged")
            $sChangeInfo = @CRLF & "Count unchanged: " & $cntAfter & " monitor(s)"
        EndIf
        Local $sMsg = "TEST 11: SUCCESS" & @CRLF & @CRLF & "Monitor list refreshed" & $sChangeInfo
        _Log("---> Example 11: Monitors after refresh: " & $cntAfter)
        If Not $g_bAutoMode Then MsgBox(64, "Example 11", $sMsg, 4)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_11

Func FuncTest_12()
    _Log('+ TEST 12: Check if monitors are connected -----------\')
    ; Check if monitors are still connected
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If $cnt < 1 Then
        Local $sMsg = "TEST 12: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 12: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 12", $sMsg, 3)
    Else
        _Log("---> Example 12: Checking connection status for " & $cnt & " monitor(s):")
        Local $sResults = ""
        Local $iConnected = 0, $iDisconnected = 0
        For $i = 1 To $cnt
            Local $bConnected = _Monitor_IsConnected($i)
            If @error Then
                _Log("  Monitor " & $i & ": ERROR checking connection: @error=" & @error)
                $sResults &= "Monitor #" & $i & ": ERROR" & @CRLF
            Else
                Local $sStatus = $bConnected ? "CONNECTED ✓" : "DISCONNECTED ✗"
                Local $aInfo = _Monitor_GetInfo($i)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                _Log("  Monitor " & $i & " (" & $sDevice & "): " & ($bConnected ? "CONNECTED" : "DISCONNECTED"))
                $sResults &= "Monitor #" & $i & " (" & $sDevice & "): " & $sStatus & @CRLF
                If $bConnected Then
                    $iConnected += 1
                Else
                    $iDisconnected += 1
                EndIf
            EndIf
        Next
        Local $sMsg = "TEST 12: COMPLETE" & @CRLF & @CRLF & "Connection Status:" & @CRLF & _
                "Connected: " & $iConnected & @CRLF & _
                "Disconnected: " & $iDisconnected & @CRLF & @CRLF & _
                $sResults
        If Not $g_bAutoMode Then MsgBox(64, "Example 12", $sMsg, 5)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_12

Func FuncTest_13()
    _Log('+ TEST 13: Get DPI scaling for monitors -------------\')
    ; Get DPI scaling for each monitor
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If $cnt < 1 Then
        Local $sMsg = "TEST 13: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 13: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 13", $sMsg, 3)
    Else
        _Log("---> Example 13: DPI information for " & $cnt & " monitor(s):")
        Local $sResults = ""
        For $i = 1 To $cnt
            Local $aDPI = _Monitor_GetDPI($i)
            If @error Then
                _Log("  Monitor " & $i & ": ERROR getting DPI: @error=" & @error)
                $sResults &= "Monitor #" & $i & ": ERROR" & @CRLF
            Else
                Local $aInfo = _Monitor_GetInfo($i)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                _Log("  Monitor " & $i & " (" & $sDevice & "):")
                _Log("    DPI X: " & $aDPI[0] & ", DPI Y: " & $aDPI[1])
                _Log("    Scaling: " & $aDPI[2] & "%")
                $sResults &= "Monitor #" & $i & " (" & $sDevice & "):" & @CRLF & _
                        "  DPI X: " & $aDPI[0] & ", DPI Y: " & $aDPI[1] & @CRLF & _
                        "  Scaling: " & $aDPI[2] & "%" & @CRLF
            EndIf
        Next
        Local $sMsg = "TEST 13: COMPLETE" & @CRLF & @CRLF & "DPI Information:" & @CRLF & $sResults
        If Not $g_bAutoMode Then MsgBox(64, "Example 13", $sMsg, 6)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_13

Func FuncTest_14()
    _Log('+ TEST 14: Get display orientation -------------------\')
    ; Get display orientation for each monitor
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If $cnt < 1 Then
        Local $sMsg = "TEST 14: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 14: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 14", $sMsg, 3)
    Else
        _Log("---> Example 14: Display orientation for " & $cnt & " monitor(s):")
        Local $sOrientationNames[4] = ["Landscape (0°)", "Portrait (90°)", "Landscape Flipped (180°)", "Portrait Flipped (270°)"]
        Local $sResults = ""
        For $i = 1 To $cnt
            Local $iOrientation = _Monitor_GetOrientation($i)
            If @error Then
                _Log("  Monitor " & $i & ": ERROR getting orientation: @error=" & @error)
                $sResults &= "Monitor #" & $i & ": ERROR" & @CRLF
            Else
                Local $aInfo = _Monitor_GetInfo($i)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                Local $sOrientationName = "Unknown"
                If $iOrientation >= 0 And $iOrientation <= 270 Then
                    Local $iIndex = Int($iOrientation / 90)
                    If $iIndex >= 0 And $iIndex < 4 Then $sOrientationName = $sOrientationNames[$iIndex]
                EndIf
                _Log("  Monitor " & $i & " (" & $sDevice & "): " & $iOrientation & "° (" & $sOrientationName & ")")
                $sResults &= "Monitor #" & $i & " (" & $sDevice & "): " & $iOrientation & "°" & @CRLF & _
                        "  " & $sOrientationName & @CRLF
            EndIf
        Next
        Local $sMsg = "TEST 14: COMPLETE" & @CRLF & @CRLF & "Display Orientation:" & @CRLF & $sResults
        If Not $g_bAutoMode Then MsgBox(64, "Example 14", $sMsg, 5)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_14

Func FuncTest_15()
    _Log('+ TEST 15: Enumerate all display modes ---------------\')
    ; Enumerate all display modes for all monitors (auto mode) or selected monitor
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If $cnt < 1 Then
        Local $sMsg = "TEST 15: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 15: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 15", $sMsg, 3)
    Else
        Local $sAllResults = ""
        Local $iTotalModes = 0

        ; In auto mode, test all monitors. Otherwise, ask user
        Local $bTestAll = $g_bAutoMode
        Local $aTestMonitors[1] = [1]  ; Default to monitor 1

        If Not $bTestAll And $cnt > 1 Then
            Local $sInput = InputBox("Example 15", "Select monitor to test (1-" & $cnt & ") or 0 for all:", "0", "", 250, 150)
            If Not @error And StringIsDigit($sInput) Then
                Local $iInput = Int($sInput)
                If $iInput = 0 Then
                    $bTestAll = True
                ElseIf $iInput >= 1 And $iInput <= $cnt Then
                    ; Test single monitor
                    $bTestAll = False
                    $aTestMonitors[0] = $iInput
                EndIf
            EndIf
        EndIf

        ; If testing all, create array of all monitor indices
        If $bTestAll Then
            ReDim $aTestMonitors[$cnt]
            For $i = 0 To $cnt - 1
                $aTestMonitors[$i] = $i + 1
            Next
        EndIf

        ; Test each monitor
        For $iMonitorIndex = 0 To UBound($aTestMonitors) - 1
            Local $iTestMonitor = $aTestMonitors[$iMonitorIndex]

            Local $aModes = _Monitor_EnumAllDisplayModes($iTestMonitor)
            If @error Then
                _Log("---> Example 15: Monitor " & $iTestMonitor & " ERROR enumerating modes: @error=" & @error)
                $sAllResults &= "Monitor #" & $iTestMonitor & ": ERROR" & @CRLF
            Else
                Local $aInfo = _Monitor_GetInfo($iTestMonitor)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                $iTotalModes += $aModes[0][0]
                _Log("---> Example 15: Monitor " & $iTestMonitor & " (" & $sDevice & ") has " & $aModes[0][0] & " display mode(s):")

                Local $sModesList = ""
                Local $iShowCount = ($aModes[0][0] > 5) ? 5 : $aModes[0][0]
                For $i = 1 To $iShowCount
                    Local $sModeInfo = $aModes[$i][0] & "x" & $aModes[$i][1] & " @ " & $aModes[$i][3] & "Hz, " & $aModes[$i][2] & " bpp"
                    _Log("  Mode " & $i & ": " & $sModeInfo)
                    $sModesList &= "  " & $sModeInfo & @CRLF
                Next
                If $aModes[0][0] > 5 Then
                    _Log("  ... and " & ($aModes[0][0] - 5) & " more mode(s)")
                    $sModesList &= "  ... and " & ($aModes[0][0] - 5) & " more mode(s)" & @CRLF
                EndIf

                If $bTestAll Then
                    $sAllResults &= "Monitor #" & $iTestMonitor & " (" & $sDevice & "): " & $aModes[0][0] & " modes" & @CRLF & $sModesList
                Else
                    $sAllResults = "Monitor #" & $iTestMonitor & " (" & $sDevice & "): " & $aModes[0][0] & " modes" & @CRLF & $sModesList
                EndIf
            EndIf
        Next

        Local $sMsg = "TEST 15: COMPLETE" & @CRLF & @CRLF
        If $bTestAll Then
            $sMsg &= "Total Modes Found: " & $iTotalModes & @CRLF & @CRLF
        EndIf
        $sMsg &= "Display Modes:" & @CRLF & $sAllResults

        If Not $g_bAutoMode Then MsgBox(64, "Example 15", $sMsg, 10)
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_15

Func FuncTest_16()
    _Log('+ TEST 16: Get monitor from rectangle ----------------\')
    ; Get monitor that overlaps with a rectangle
    _Monitor_GetList()
    Local $cnt = _Monitor_GetCount()
    If $cnt < 1 Then
        Local $sMsg = "TEST 16: SKIPPED" & @CRLF & "No monitors detected"
        _Log("---> Example 16: No monitors detected")
        If Not $g_bAutoMode Then MsgBox(64, "Example 16", $sMsg, 3)
    Else
        ; Create a test rectangle (center of primary monitor)
        Local $prim = _Monitor_GetPrimary()
        If $prim = 0 Then $prim = 1

        Local $iLeft, $iTop, $iRight, $iBottom
        _Monitor_GetBounds($prim, $iLeft, $iTop, $iRight, $iBottom)

        Local $iCenterX = $iLeft + ($iRight - $iLeft) / 2
        Local $iCenterY = $iTop + ($iBottom - $iTop) / 2
        Local $iRectW = 200, $iRectH = 150

        Local $iRectLeft = $iCenterX - $iRectW / 2
        Local $iRectTop = $iCenterY - $iRectH / 2
        Local $iRectRight = $iCenterX + $iRectW / 2
        Local $iRectBottom = $iCenterY + $iRectH / 2

        _Log("---> Example 16: Testing rectangle: L=" & $iRectLeft & ", T=" & $iRectTop & ", R=" & $iRectRight & ", B=" & $iRectBottom)

        Local $iMonitor = _Monitor_GetFromRect($iRectLeft, $iRectTop, $iRectRight, $iRectBottom)
        If @error Then
            Local $sMsg = "TEST 16: FAILED" & @CRLF & "ERROR getting monitor from rect" & @CRLF & "@error=" & @error
            _Log("---> Example 16: ERROR getting monitor from rect: @error=" & @error)
            If Not $g_bAutoMode Then MsgBox(48, "Example 16", $sMsg, 3)
        Else
            If $iMonitor > 0 Then
                Local $aInfo = _Monitor_GetInfo($iMonitor)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                Local $sPrimary = @error ? "" : ($aInfo[9] ? " [PRIMARY]" : "")
                Local $sMsg = "TEST 16: SUCCESS" & @CRLF & @CRLF & _
                        "Test Rectangle:" & @CRLF & _
                        "  Left: " & $iRectLeft & ", Top: " & $iRectTop & @CRLF & _
                        "  Right: " & $iRectRight & ", Bottom: " & $iRectBottom & @CRLF & @CRLF & _
                        "Found Monitor: #" & $iMonitor & $sPrimary & @CRLF & _
                        "Device: " & $sDevice
                _Log("---> Example 16: Rectangle overlaps with Monitor #" & $iMonitor & " (" & $sDevice & ")")
                If Not $g_bAutoMode Then MsgBox(64, "Example 16", $sMsg, 4)
            Else
                Local $sMsg = "TEST 16: NO MATCH" & @CRLF & @CRLF & "Rectangle does not overlap with any monitor"
                _Log("---> Example 16: Rectangle does not overlap with any monitor")
                If Not $g_bAutoMode Then MsgBox(48, "Example 16", $sMsg, 3)
            EndIf
        EndIf
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_16

Func FuncTest_17()
    _Log('+ TEST 17: Get monitor from window -------------------\')
    ; Get monitor containing a specific window
    _Monitor_GetList()

    ; Try to find Notepad or use current active window
    Local $hWnd = WinGetHandle("[CLASS:Notepad]")
    If @error Or Not $hWnd Then
        $hWnd = WinGetHandle("[ACTIVE]")
        If @error Or Not $hWnd Then
            _Log("---> Example 17: No suitable window found. Opening Notepad...")
            $pidNotepad = Run("notepad.exe")
            If WinWaitActive("[CLASS:Notepad]", "", 3) Then
                Sleep(500)
                $hWnd = WinGetHandle("[CLASS:Notepad]")
            Else
                _Log("---> Example 17: ERROR - Could not find or create test window")
                Return
            EndIf
        EndIf
    EndIf

    If $hWnd Then
        Local $sTitle = WinGetTitle($hWnd)
        _Log("---> Example 17: Testing window: " & ($sTitle = "" ? "[No Title]" : $sTitle) & " (Handle: " & $hWnd & ")")

        Local $iMonitor = _Monitor_GetFromWindow($hWnd)
        If @error Then
            _Log("---> Example 17: ERROR getting monitor from window: @error=" & @error)
        Else
            If $iMonitor > 0 Then
                Local $aInfo = _Monitor_GetInfo($iMonitor)
                Local $sDevice = @error ? "N/A" : $aInfo[10]
                Local $sPrimary = @error ? "" : ($aInfo[9] ? " [PRIMARY]" : "")
                _Log("---> Example 17: Window is on Monitor #" & $iMonitor & " (" & $sDevice & ")" & $sPrimary)

                ; Get window position for verification
                Local $aWinPos = WinGetPos($hWnd)
                If Not @error Then
                    _Log("  Window position: X=" & $aWinPos[0] & ", Y=" & $aWinPos[1])
                    _Log("  Monitor bounds: L=" & $aInfo[1] & ", T=" & $aInfo[2] & ", R=" & $aInfo[3] & ", B=" & $aInfo[4])
                EndIf

                Local $sMsg = "TEST 17: SUCCESS" & @CRLF & @CRLF & _
                        "Window: " & ($sTitle = "" ? "[No Title]" : $sTitle) & @CRLF & _
                        "Handle: " & $hWnd & @CRLF & @CRLF & _
                        "Monitor: #" & $iMonitor & $sPrimary & @CRLF & _
                        "Device: " & $sDevice & @CRLF & @CRLF & _
                        "Window Position:" & @CRLF & _
                        "  X: " & $aWinPos[0] & ", Y: " & $aWinPos[1]
                If Not $g_bAutoMode Then MsgBox(64, "Example 17", $sMsg, 5)
            Else
                _Log("---> Example 17: Window is not on any monitor")
            EndIf
        EndIf
    EndIf
    _Log('- End ------------------------------------------------/')
EndFunc   ;==>FuncTest_17

; ==================================================================================================
; Small helper to push item into dynamic array (simple)
; ==================================================================================================
Func __ArrayAdd(ByRef $a, $v)
    Local $n = 0
    If IsArray($a) Then $n = UBound($a)
    ReDim $a[$n + 1]
    $a[$n] = $v
EndFunc   ;==>__ArrayAdd
