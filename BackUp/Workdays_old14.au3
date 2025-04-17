#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=calendar.ico
#AutoIt3Wrapper_Res_Description=Work Day management
#AutoIt3Wrapper_Res_Fileversion=1.0.1.4
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=Work Days
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#pragma compile(inputboxres, true)
Opt("TrayIconHide", 1)
Opt("TrayAutoPause", 0)
;~ Opt("TrayMenuMode", 3)

; Script Start - Add your code below here
#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <ListViewConstants.au3>
#include <Array.au3>
#include <WinAPIRes.au3>
#include <WinAPIInternals.au3>
;~ #include <ColorChooser.au3>
#include <ColorPicker.au3>
#include <WinAPI.au3>

Global $IniSection[999][999]
Global $LabelMonth[99999]
Global $LabelMonthX[99999]
Global $Inputs[32][32]

Global $Year = @YEAR
Global $Ratio_Q1 = 0
Global $Ratio_Q2 = 0
Global $Ratio_Q3 = 0
Global $Ratio_Q4 = 0

Global $White = 0xFFFFFF
Global $Black = 0x000000

$DB = "HKEY_CURRENT_USER\Software\WorkDays"
_CriaINI(@YEAR)
;~ _ReadColors()

	Global $CalendarTag = RegRead($DB, "caltag")
	If @error Then $CalendarTag = "1"

	Global $Color_bk_OnSite = RegRead($DB, "Color_OnSite")
	If @error Then $Color_bk_OnSite = 0x00CC66


	Global $Color_bk_Remote = RegRead($DB, "Color_Remote")
	If @error Then $Color_bk_Remote = 0x0080FF

	Global $Color_bk_holiday = RegRead($DB, "Color_holiday")
	If @error Then $Color_bk_holiday = 0xFFFFCC

	Global $Color_bk_PTO = RegRead($DB, "Color_PTO")
	If @error Then $Color_bk_PTO = 0x66FFFF

	Global $Color_bk_Travel = RegRead($DB, "Color_Travel")
	If @error Then $Color_bk_Travel = 0xFF8000

	Global $Color_bk_Sick = RegRead($DB, "Color_Sick")
	If @error Then $Color_bk_Sick = 0xFF6666

	Global $Color_bk_Blank = RegRead($DB, "Color_Blank")
	If @error Then $Color_bk_Blank = 0xFFFFFF

	Global $Color_bk_Weekend = RegRead($DB, "Color_Weekend")
	If @error Then $Color_bk_Weekend = 0xA0A0A0


	Global $Picker_Font_OnSite_Read = RegRead($DB, "Font_OnSite")
	Global $Font_OnSite = $Black
	If $Picker_Font_OnSite_Read = 1 Then
		$Font_OnSite = $White
	EndIf

	Global $Picker_Font_Remote_Read = RegRead($DB, "Font_Remote")
	Global $Font_Remote = $Black
	If $Picker_Font_Remote_Read = 1 Then
		$Font_Remote = $White
	EndIf

	Global $Picker_Font_Holiday_Read = RegRead($DB, "Font_holiday")
	Global $Font_Holiday = $Black
	If $Picker_Font_Holiday_Read = 1 Then
		$Font_Holiday = $White
	EndIf

	Global $Picker_Font_PTO_Read = RegRead($DB, "Font_PTO")
	Global $Font_PTO = $Black
	If $Picker_Font_PTO_Read = 1 Then
		$Font_PTO = $White
	EndIf

	Global $Picker_Font_Travel_Read = RegRead($DB, "Font_Travel")
	Global $Font_Travel = $Black
	If $Picker_Font_Travel_Read = 1 Then
		$Font_Travel = $White
	EndIf

	Global $Picker_Font_Sick_Read = RegRead($DB, "Font_Sick")
	Global $Font_Sick = $Black
	If $Picker_Font_Sick_Read = 1 Then
		$Font_Sick = $White
	EndIf


	Global $Picker_Font_Blank_Read = RegRead($DB, "Font_Blank")
	Global $Font_Blank = $Black
	If $Picker_Font_Blank_Read = 1 Then
		$Font_Blank = $White
	EndIf

	Global $Picker_Font_Weekend_Read = RegRead($DB, "Font_Weekend")
	Global $Font_Weekend = $Black
	If $Picker_Font_Weekend_Read = 1 Then
		$Font_Weekend = $White
	EndIf

$Form_WorkDays = GUICreate("Work Days", 1140, 620, -1, -1)

Global $DBpMenu_db = GUICtrlCreateMenu("File")
Global $DBpMenu_backup_Data = GUICtrlCreateMenu("Data", $DBpMenu_db)
Global $DBpMenu_backup = GUICtrlCreateMenuItem("Create Backup", $DBpMenu_backup_Data)
Global $BkpMenu_Batch = GUICtrlCreateMenuItem("Restore Backup", $DBpMenu_backup_Data)
Global $DBpMenu_backup_2 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
Global $BkpMenu_reset_all = GUICtrlCreateMenuItem("Reset Entire Database", $DBpMenu_backup_Data)
Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
Global $DBpMenu_backup_Data_Holidays = GUICtrlCreateMenuItem("Import Holidays File", $DBpMenu_backup_Data)
;~ Global $BkpMenu_reset = GUICtrlCreateMenu("Reset Data", $DBpMenu_db)
Global $BkpMenu_reset_1 = GUICtrlCreateMenuItem("", $DBpMenu_db)
Global $BkpMenu_Exit = GUICtrlCreateMenuItem("&Exit", $DBpMenu_db)

Global $DBpMenu_settings = GUICtrlCreateMenu("Settings")
Global $BkpMenu_settings_BKcolors = GUICtrlCreateMenuItem("Background Colors", $DBpMenu_settings)
Global $BkpMenu_help = GUICtrlCreateMenu("?")
Global $BkpMenu_help_help = GUICtrlCreateMenuItem("Help", $BkpMenu_help)
Global $BkpMenu_help_space = GUICtrlCreateMenuItem("", $BkpMenu_help)
Global $BkpMenu_help_About = GUICtrlCreateMenuItem("About", $BkpMenu_help)
;~ Global $BkpMenu_reset_year = GUICtrlCreateMenuItem("Reset Specific Year", $BkpMenu_reset)

$Calendar = GUICtrlCreateMonthCal(@YEAR & "/" & @MON & "/" & @MDAY, 8, 8, 273, 201)
$Group1 = GUICtrlCreateGroup("", 288, 8, 270, 200)
$Input_SelDate = GUICtrlCreateInput("", 376, 24, 70, 21, $ES_READONLY)
GUICtrlSetData($Input_SelDate, GUICtrlRead($Calendar))
GUICtrlSetColor($Input_SelDate, 0x990000)
$Label1 = GUICtrlCreateLabel("Selected Date:", 296, 28, 75, 17)
$Input_Quarter = GUICtrlCreateInput("", 450, 24, 20, 21, $ES_READONLY)
GUICtrlSetColor($Input_Quarter, 0x00994C)

$Input_Tip = GUICtrlCreateInput("", 296, 54, 175, 21, $ES_READONLY)

$Checkbox_calendtarTag = GUICtrlCreateCheckbox("Calendar Tag", 472, 54) ;## Calendar TAG
If $CalendarTag = "1" Then
	GUICtrlSetState($Checkbox_calendtarTag, $gui_checked)
Else
	GUICtrlSetState($Checkbox_calendtarTag, $gui_unchecked)
EndIf

$Button_OnSite = GUICtrlCreateButton("&On Site", 296, 84, 75, 25)
GUICtrlSetBkColor($Button_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Button_OnSite, $Font_OnSite)


$Button_Remote = GUICtrlCreateButton("&Remote", 384, 84, 75, 25)
GUICtrlSetBkColor($Button_Remote, $Color_bk_Remote)
GUICtrlSetColor($Button_Remote, $Font_Remote)

$Button_holiday = GUICtrlCreateButton("&Holiday", 296, 114, 75, 25)
GUICtrlSetBkColor($Button_holiday, $Color_bk_holiday)
GUICtrlSetColor($Button_holiday, $Font_Holiday)

$Button_PTO = GUICtrlCreateButton("&PTO", 384, 114, 75, 25)
GUICtrlSetBkColor($Button_PTO, $Color_bk_PTO)
GUICtrlSetColor($Button_PTO, $Font_PTO)

$Button_Travel = GUICtrlCreateButton("&Travel", 296, 144, 75, 25)
GUICtrlSetBkColor($Button_Travel, $Color_bk_Travel)
GUICtrlSetColor($Button_Travel, $Font_Travel)

$Button_Sick = GUICtrlCreateButton("&Sick", 384, 144, 75, 25)
GUICtrlSetBkColor($Button_Sick, $Color_bk_Sick)
GUICtrlSetColor($Button_Sick, $Font_Sick)

$Button_Blank = GUICtrlCreateButton("&Blank", 296, 174, 75, 25)
GUICtrlSetBkColor($Button_Blank, $Color_bk_Blank)
GUICtrlSetColor($Button_Blank, $Font_Blank)

$Button_Weekend = GUICtrlCreateButton("&Weekend", 384, 174, 75, 25)
GUICtrlSetBkColor($Button_Weekend, $Color_bk_Weekend)
GUICtrlSetColor($Button_Weekend, $Font_Weekend)


GUICtrlCreateGroup("", -99, -99, 1, 1)


$Button_Reload = GUICtrlCreateButton("Reload Data", 472, 22, 75, 25)

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q1 = GUICtrlCreateGroup(" Q1 - " & @YEAR, 570, 8, 270, 100)

$Label_1_q1 = GUICtrlCreateLabel("Total Days:", 571, 30, 79, 21, $SS_RIGHT)
$Label_2_q1 = GUICtrlCreateLabel("Work Days:", 571, 50, 79, 21, $SS_RIGHT)
$Label_3_q1 = GUICtrlCreateLabel("Ratio:", 571, 70, 79, 21, $SS_RIGHT)
$Label_ratio_q1 = GUICtrlCreateLabel("Ratio to Date:", 571, 89, 79, 16, $SS_RIGHT)
GUICtrlSetColor($Label_ratio_q1, 0x0066CC)

$Label_4_q1 = GUICtrlCreateLabel("Estim.On-Site:", 700, 30, 65, 21, $SS_RIGHT)
$Label_5_q1 = GUICtrlCreateLabel("Real On-Site:", 700, 50, 65, 21, $SS_RIGHT)
$Label_6_q1 = GUICtrlCreateLabel("Remaining:", 700, 70, 65, 21, $SS_RIGHT)

$Input_TD_q1 = GUICtrlCreateInput("", 651, 25, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q1 = GUICtrlCreateInput("", 651, 45, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q1 = GUICtrlCreateInput("", 651, 65, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q1 = GUICtrlCreateInput("", 651, 85, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q1 = GUICtrlCreateInput("", 770, 25, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q1 = GUICtrlCreateInput("", 770, 45, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q1 = GUICtrlCreateInput("", 770, 65, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q2x = GUICtrlCreateGroup(" Q2 - " & @YEAR, 855, 8, 270, 100)

$Label_1_q2 = GUICtrlCreateLabel("Total Days:", 856, 30, 79, 21, $SS_RIGHT)
$Label_2_q2 = GUICtrlCreateLabel("Work Days:", 856, 50, 79, 21, $SS_RIGHT)
$Label_3_q2 = GUICtrlCreateLabel("Ratio:", 856, 70, 79, 21, $SS_RIGHT)
$Label_Ratio_q2 = GUICtrlCreateLabel("Ratio to Date:", 856, 90, 79, 16, $SS_RIGHT)
GUICtrlSetColor($Label_Ratio_q2, 0x0066CC)

$Label_4_q2 = GUICtrlCreateLabel("Estim.On-Site:", 985, 30, 65, 21, $SS_RIGHT)
$Label_5_q2 = GUICtrlCreateLabel("Real On-Site:", 985, 50, 65, 21, $SS_RIGHT)
$Label_6_q2 = GUICtrlCreateLabel("Remaining:", 985, 70, 65, 21, $SS_RIGHT)

$Input_TD_q2 = GUICtrlCreateInput("", 936, 25, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q2 = GUICtrlCreateInput("", 936, 45, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q2 = GUICtrlCreateInput("", 936, 65, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q2 = GUICtrlCreateInput("", 936, 85, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q2 = GUICtrlCreateInput("", 1055, 25, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q2 = GUICtrlCreateInput("", 1055, 45, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q2 = GUICtrlCreateInput("", 1055, 65, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q3 = GUICtrlCreateGroup(" Q3 - " & @YEAR, 570, 108, 270, 100)

$Label_1_q3 = GUICtrlCreateLabel("Total Days:", 571, 130, 79, 21, $SS_RIGHT)
$Label_2_q3 = GUICtrlCreateLabel("Work Days:", 571, 150, 79, 21, $SS_RIGHT)
$Label_3_q3 = GUICtrlCreateLabel("Ratio:", 571, 170, 79, 21, $SS_RIGHT)
$Label_Ratio_q3 = GUICtrlCreateLabel("Ratio to Date:", 571, 190, 79, 16, $SS_RIGHT)
GUICtrlSetColor($Label_Ratio_q3, 0x0066CC)

$Label_4_q3 = GUICtrlCreateLabel("Estim.On-Site:", 700, 130, 65, 21, $SS_RIGHT)
$Label_5_q3 = GUICtrlCreateLabel("Real On-Site:", 700, 150, 65, 21, $SS_RIGHT)
$Label_6_q3 = GUICtrlCreateLabel("Remaining:", 700, 170, 65, 21, $SS_RIGHT)

$Input_TD_q3 = GUICtrlCreateInput("", 651, 125, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q3 = GUICtrlCreateInput("", 651, 145, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q3 = GUICtrlCreateInput("", 651, 165, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q3 = GUICtrlCreateInput("", 651, 185, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q3 = GUICtrlCreateInput("", 770, 125, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q3 = GUICtrlCreateInput("", 770, 145, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q3 = GUICtrlCreateInput("", 770, 165, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q4 = GUICtrlCreateGroup(" Q4 - " & @YEAR, 855, 108, 270, 100)

$Label_1_q4 = GUICtrlCreateLabel("Total Days:", 856, 130, 79, 21, $SS_RIGHT)
$Label_2_q4 = GUICtrlCreateLabel("Work Days:", 856, 150, 79, 21, $SS_RIGHT)
$Label_3_q4 = GUICtrlCreateLabel("Ratio:", 856, 170, 79, 21, $SS_RIGHT)
$Label_Ratio_q4 = GUICtrlCreateLabel("Ratio to Date:", 856, 190, 79, 16, $SS_RIGHT)
GUICtrlSetColor($Label_Ratio_q4, 0x0066CC)

$Label_4_q4 = GUICtrlCreateLabel("Estim.On-Site:", 985, 130, 65, 21, $SS_RIGHT)
$Label_5_q4 = GUICtrlCreateLabel("Real On-Site:", 985, 150, 65, 21, $SS_RIGHT)
$Label_6_q4 = GUICtrlCreateLabel("Remaining:", 985, 170, 65, 21, $SS_RIGHT)

$Input_TD_q4 = GUICtrlCreateInput("", 936, 125, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q4 = GUICtrlCreateInput("", 936, 145, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q4 = GUICtrlCreateInput("", 936, 165, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q4 = GUICtrlCreateInput("", 936, 185, 40, 20, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q4 = GUICtrlCreateInput("", 1055, 125, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q4 = GUICtrlCreateInput("", 1055, 145, 40, 20, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q4 = GUICtrlCreateInput("", 1055, 165, 40, 20, BitOR($ES_CENTER, $ES_READONLY))


GUICtrlSetState($Input_RaTio_q1, $gui_hide)
GUICtrlSetState($Input_RaTio_q2, $gui_hide)
GUICtrlSetState($Input_RaTio_q3, $gui_hide)
GUICtrlSetState($Input_RaTio_q4, $gui_hide)

GUICtrlCreateGroup("", -99, -99, 1, 1)


GUICtrlCreateGroup("", 10, 315, 1120, 9)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlCreateGroup("", 10, 400, 1120, 9)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlCreateGroup("", 10, 485, 1120, 9)
GUICtrlSetColor(-1, 0x0000FF)
$StatusBar1 = _GUICtrlStatusBar_Create($Form_WorkDays)

$sMessage = "Developed by Fabricio Zambroni - VERSION: " & FileGetVersion(@ScriptFullPath) & " - Today: " & @YEAR & "/" & @MON & "/" & @MDAY
_GUICtrlStatusBar_SetText($StatusBar1, $sMessage)


_ReadINI(@YEAR)

_CheckQuarter()

_AutoBKP()

GUISetState(@SW_SHOW)


While 1
	$nMsg = GUIGetMsg()

	For $J = 1 To 12
		For $i = 1 To 31
			If $Inputs[$i][$J] <> 0 And $nMsg = $Inputs[$i][$J] Then
				If $i < 10 Then
					$n = "0" & $i
				Else
					$n = $i
				EndIf

				If $J < 10 Then
					$s = "0" & $J
				Else
					$s = $J
				EndIf
				$FullDate = GUICtrlRead($Input_SelDate)
				$FullDate_Split = StringSplit($FullDate, "/")
				$ClickedDate = $FullDate_Split[1] & "/" & $s & "/" & $n

				GUICtrlSetData($Calendar, $ClickedDate)
				_CalendarRead()

			EndIf
		Next
	Next

	Switch $nMsg

		Case $BkpMenu_Exit
			Exit

		Case $BkpMenu_settings_BKcolors
			$Return_Color = _BKColorPallet()
			If $Return_Color = 1 Then

				$Color_bk_OnSite = RegRead($DB, "Color_OnSite")
				If @error Then $Color_bk_OnSite = 0x00CC66

				$Color_bk_Remote = RegRead($DB, "Color_Remote")
				If @error Then $Color_bk_Remote = 0x0080FF

				$Color_bk_holiday = RegRead($DB, "Color_holiday")
				If @error Then $Color_bk_holiday = 0xFFFFCC

				$Color_bk_PTO = RegRead($DB, "Color_PTO")
				If @error Then $Color_bk_PTO = 0x66FFFF

				$Color_bk_Travel = RegRead($DB, "Color_Travel")
				If @error Then $Color_bk_Travel = 0xFF8000

				$Color_bk_Sick = RegRead($DB, "Color_Sick")
				If @error Then $Color_bk_Sick = 0xFF6666

				$Color_bk_Blank = RegRead($DB, "Color_Blank")
				If @error Then $Color_bk_Blank = 0xFFFFFF

				$Color_bk_Weekend = RegRead($DB, "Color_Weekend")
				If @error Then $Color_bk_Weekend = 0xA0A0A0

				GUICtrlSetBkColor($Button_OnSite, $Color_bk_OnSite)
				GUICtrlSetBkColor($Button_Remote, $Color_bk_Remote)
				GUICtrlSetBkColor($Button_holiday, $Color_bk_holiday)
				GUICtrlSetBkColor($Button_PTO, $Color_bk_PTO)
				GUICtrlSetBkColor($Button_Travel, $Color_bk_Travel)
				GUICtrlSetBkColor($Button_Sick, $Color_bk_Sick)
				GUICtrlSetBkColor($Button_Blank, $Color_bk_Blank)
				GUICtrlSetBkColor($Button_Weekend, $Color_bk_Weekend)

				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				_ReadINI($SelDate_slipt[1])
			EndIf

		Case $BkpMenu_help_help
			$HelpFile = @ScriptDir & "\help.pdf"
			If Not FileExists($HelpFile) Then
				MsgBox(262160, "Work Days", "Help file not found in the application folder.")
			Else
				ShellExecute($HelpFile)
			EndIf

		Case $BkpMenu_help_About
			MsgBox(262144 + 64, "Work Days", "Work Days is a user-friendly calendar-based application for managing and categorizing your workdays�On Site, Remote, and Holiday�throughout the year." & @CRLF & @CRLF & "Version: 1.0.1.3 - Custom colors and bug fixes" & @CRLF & @CRLF & "Developed by Fabricio Zambroni - CURRENT VERSION: " & FileGetVersion(@ScriptFullPath))

		Case $GUI_EVENT_CLOSE
			Exit

		Case $Checkbox_calendtarTag
;~ 			MsgBox(262144,"",GUICtrlRead($Checkbox_calendtarTag))
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				RegWrite($DB, "caltag", "REG_SZ", "1")
			Else
				RegWrite($DB, "caltag", "REG_SZ", "0")
			EndIf

		Case $BkpMenu_Batch
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Batch Import", "**WARNING** Importing data will overwrite any existing records. Do you want to proceed?" & @CRLF & @CRLF & "Check the help file for more details.")
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					$ResetReturn = _ResetDatabase("1")
					If $ResetReturn = "1" Then
						_CriaINI(@YEAR)
						_RestoreBackup()
					Else
						MsgBox(262160, "Reset Database", "Oops! Something went wrong. Please try again.")
					EndIf
					_ClearScreen()
					_ReadColors()
;~ 					_CriaINI(@YEAR)
					_ReadINI(@YEAR)

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		Case $DBpMenu_backup_Data_Holidays
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Holidays Import", "**WARNING** Importing data will overwrite any existing records for the selected dates. Do you want to proceed?")
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					_ImportHolidays()
					_ClearScreen()
					_ReadColors()
					_CriaINI(@YEAR)
					_ReadINI(@YEAR)

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect


		Case $Calendar
			_CalendarRead()

		Case $DBpMenu_backup
			_CreateBackup()

		Case $BkpMenu_reset_all
			_ResetDatabase()
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			_ClearScreen()
			_ReadColors()
			_CriaINI(@YEAR)
			_ReadINI(@YEAR)

		Case $Button_Reload
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
;~ 			_ClearScreen()
			_ReadINI($SelDate_slipt[1])


		Case $Button_OnSite
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				$holidayName = InputBox("Calendar Tag", "Give a tag name for this ON SITE event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
				If Not @error Then
					$holidayName = StringReplace($holidayName, "-", "=")
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "O" & $holidayName)
;~ 					_ReadINI($SelDate_slipt[1])
					_Update($SelDate)
				EndIf
			Else
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "O" & $holidayName)
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
			EndIf

		Case $Button_Blank
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")

			$WeekDayNum = _DateToDayOfWeek($SelDate_slipt[1], $SelDate_slipt[2], $SelDate_slipt[3])

			$WeekEnd = 0
			If $WeekDayNum = "1" Then
				$WeekEnd = 1
			EndIf
			If $WeekDayNum = "7" Then
				$WeekEnd = 1
			EndIf
			If $WeekEnd = 1 Then
				MsgBox(262160, "Weekend", "Weekends cannot be blank.")
			Else
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "")
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
			EndIf

		Case $Button_Remote
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				$holidayName = InputBox("Calendar Tag", "Give a tag name for this REMOTE WORK event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
				If Not @error Then
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "R" & $holidayName)
;~ 					_ReadINI($SelDate_slipt[1])
					_Update($SelDate)
				EndIf
			Else
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "R" & $holidayName)
				_Update($SelDate)
;~ 				_ReadINI($SelDate_slipt[1])
			EndIf
		Case $Button_Travel
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				$holidayName = InputBox("Calendar Tag", "Give a tag name for this TRAVEL event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
				If Not @error Then
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "T" & $holidayName)
;~ 					_ReadINI($SelDate_slipt[1])
					_Update($SelDate)
				EndIf
			Else
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "T" & $holidayName)
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
			EndIf

		Case $Button_PTO
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				$holidayName = InputBox("Calendar Tag", "Give a tag name for this PTO event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
				If Not @error Then
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "P" & $holidayName)
;~ 					_ReadINI($SelDate_slipt[1]).
					_Update($SelDate)
				EndIf
			Else
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "P" & $holidayName)
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
			EndIf

		Case $Button_holiday
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				$holidayName = InputBox("Calendar Tag", "Give a tag name for this HOLIDAY event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
				If Not @error Then
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "H" & $holidayName)
;~ 					_ReadINI($SelDate_slipt[1])
					_Update($SelDate)
				EndIf
			Else
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "H" & $holidayName)
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
			EndIf

		Case $Button_Sick
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			If GUICtrlRead($Checkbox_calendtarTag) = "1" Then
				$holidayName = InputBox("Calendar Tag", "Give a tag name for this SICK DAY event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
				If Not @error Then
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "S" & $holidayName)
;~ 					_ReadINI($SelDate_slipt[1])
					_Update($SelDate)
				EndIf
			Else
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "S" & $holidayName)
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
			EndIf


		Case $Button_Weekend
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")

			$WeekDayNum = _DateToDayOfWeek($SelDate_slipt[1], $SelDate_slipt[2], $SelDate_slipt[3])

			$WeekEnd = 0
			If $WeekDayNum <> "1" And $WeekDayNum <> "7" Then
				$WeekEnd = 1
			EndIf

			If $WeekEnd = 1 Then
				MsgBox(262160, "Weekend", "This date is not a weekend.")
			Else
;~ 				$holidayName = InputBox("Calendar Tag", "Give a tag name for this WEEKEND event on " & $SelDate & ":", "", "", -1, -1, Default, Default, 0, $Form_WorkDays)
;~ 				If Not @error Then
				$holidayName = ""
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "W" & $holidayName)
;~ 				_ReadINI($SelDate_slipt[1])
				_Update($SelDate)
;~ 				EndIf
			EndIf

	EndSwitch


WEnd

Func _RestoreBackup()

	$HolidaysError = ""
	$HolidaysSucess = ""
	$ImportCount = 0

	$HolidaysFile = FileOpenDialog("File to import", @ScriptDir, "All (*.*)", 3, "", $Form_WorkDays)
	If @error Then
		MsgBox(262160, "Import", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error)
	Else
		$FileHolidays_hwd = FileOpen($HolidaysFile, 0)
		If $FileHolidays_hwd = -1 Then

		Else

			While 1
				$HolidaysLine = FileReadLine($FileHolidays_hwd)
				If @error = -1 Then ExitLoop
				If @error = 1 Then
					MsgBox(262160, "Import", "Oops! Something went wrong when read the file. Please try again." & @CRLF & "Error code: " & @error)
					Return
				EndIf
				If Not StringInStr($HolidaysLine, "\") Then
					If Not StringInStr($HolidaysLine, "=") Then
						$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
					Else
						$HolidaysLine_Setting = StringSplit($HolidaysLine,"=")
						$RegError = RegWrite($DB, $HolidaysLine_Setting[1], "REG_SZ",$HolidaysLine_Setting[2])
						If @error Then
							$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
						Else
							$ImportCount += 1
						EndIf
					EndIf
				Else



				EndIf

				#cs
				If StringInStr($HolidaysLine, "-") Then
					$HolidaysLineSplited = StringSplit($HolidaysLine, "-")
					If _DateIsValid($HolidaysLineSplited[1]) Then
						$HolidaysDateSplited = StringSplit($HolidaysLineSplited[1], "/")
						If @error Then
							$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
						Else

							If $HolidaysLineSplited[2] = "O" Or $HolidaysLineSplited[2] = "R" Or $HolidaysLineSplited[2] = "B" Or $HolidaysLineSplited[2] = "T" Or $HolidaysLineSplited[2] = "P" Or $HolidaysLineSplited[2] = "H" Or $HolidaysLineSplited[2] = "S" Then

								$RegError = RegWrite($DB & "\" & $HolidaysDateSplited[1] & "\" & $HolidaysDateSplited[2], $HolidaysDateSplited[3], "REG_SZ", $HolidaysLineSplited[2] & $HolidaysLineSplited[3])
								$ImportCount += 1
							Else
								If $HolidaysLineSplited[2] = "W" Then
									$DayofWeek = _DateToDayOfWeek($HolidaysDateSplited[1], $HolidaysDateSplited[2], $HolidaysDateSplited[3])
									If $DayofWeek = "1" Or $DayofWeek = "7" Then
										$RegError = RegWrite($DB & "\" & $HolidaysDateSplited[1] & "\" & $HolidaysDateSplited[2], $HolidaysDateSplited[3], "REG_SZ", $HolidaysLineSplited[2] & $HolidaysLineSplited[3])
										$ImportCount += 1
									Else
										$HolidaysError = $HolidaysError & $HolidaysLine & " - DATE IS NOT A WEEKEND" & @CRLF
									EndIf
								Else
									If $HolidaysLineSplited[2] <> "" Then
										$HolidaysError = $HolidaysError & $HolidaysLine & " - INVALID TYPE OF OPERATION" & @CRLF
									EndIf
								EndIf
							EndIf

							If @error Then
								$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
							Else
								$HolidaysSucess = $HolidaysSucess & $HolidaysLine & @CRLF
							EndIf

						EndIf

					Else
						$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
					EndIf
				Else
					$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
				EndIf
				#ce

			WEnd

			If $HolidaysError <> "" Then
				MsgBox(262160, "Import", "Oops! Something went wrong when read the file." & @CRLF & "The following lines was not imported:" & @CRLF & @CRLF & $HolidaysError & @CRLF & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess)
			Else
				If $ImportCount > 10 Then
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & @CRLF & $ImportCount & " lines imported.")
				Else
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess)
				EndIf
			EndIf
		EndIf

	EndIf

EndFunc   ;==>_ImportBatch


Func _Update($SelDate)

	$SelDate_splited = StringSplit($SelDate, "/")
	$Data_year = Number($SelDate_splited[1])
	$Data_month = Number($SelDate_splited[2])
	$Data_day = Number($SelDate_splited[3])

	If $Data_month < 10 Then $Data_month = "0" & $Data_month
	If $Data_day < 10 Then $Data_day = "0" & $Data_day

	$Data_Register1 = RegRead($DB & "\" & $Data_year & "\" & $Data_month, $Data_day)
	$Data_Register = StringLeft($Data_Register1, 1)

	If StringLen($Data_Register1) > 1 Then
		$tip = " - " & StringTrimLeft($Data_Register1, 1)
		GUICtrlSetData($Input_Tip, StringTrimLeft($Data_Register1, 1))
	Else
		$tip = ""
		GUICtrlSetData($Input_Tip, $tip)
	EndIf
	$WeekDayNum = _DateToDayOfWeek($Data_year, $Data_month, $Data_day)
	$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
	GUICtrlSetData($Inputs[$Data_day][$Data_month], $Data_Register)
	GUICtrlSetTip($Inputs[$Data_day][$Data_month], $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & $tip)

	If $Data_Register = "T" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Travel) ; Travel

	If $Data_Register = "W" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Weekend)             ; Weekend
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Color_bk_Blank)
	EndIf

	If $Data_Register = "O" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_OnSite) ; On-site

	If $Data_Register = "R" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Remote) ; Remote

	If $Data_Register = "P" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_PTO) ; PTO

	If $Data_Register = "H" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_holiday) ; holiday

	If $Data_Register = "S" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Sick) ; Sick

	If $Data_Register = "" Then GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Blank) ; Blank

	If $Data_year & "/" & $Data_month & "/" & $Data_day = @YEAR & "/" & @MON & "/" & @MDAY Then

		GUICtrlSetColor($Inputs[$Data_day][$Data_month], 0xFFCCCC)
		GUICtrlSetFont($Inputs[$Data_day][$Data_month], 10, 1200, "", "", 5)

		GUICtrlSetTip($Inputs[$Data_day][$Data_month], $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & " - TODAY" & $tip)
		If $Data_Register = "" Then
			$Data_Register = "X"
			GUICtrlSetColor($Inputs[$Data_day][$Data_month], 0xFF0000)             ; today
		EndIf
		GUICtrlSetData($Inputs[$Data_day][$Data_month], "[" & $Data_Register & "]")
	EndIf

	_ReadStatistics($Data_year)

EndFunc   ;==>_Update

Func _AutoBKP()
	$BKPDB = @ScriptDir & "\autosave.db"
	If Not FileExists($BKPDB) Then
		_CreateBackup($BKPDB)
	Else
		$AutoSaveDate = FileGetTime($BKPDB)
		If _DateDiff('D', $AutoSaveDate[0] & "/" & $AutoSaveDate[1] & "/" & $AutoSaveDate[2], @YEAR & "/" & @MON & "/" & @MDAY) > 7 Then
			_CreateBackup($BKPDB)
		EndIf
	EndIf

EndFunc   ;==>_AutoBKP

Func _ImportHolidays()

	$HolidaysError = ""
	$HolidaysSucess = ""
	$ImportCount = 0

	$HolidaysFile = FileOpenDialog("File to import", @ScriptDir, "All (*.*)", 3, "", $Form_WorkDays)
	If @error Then
		MsgBox(262160, "Import", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error)
	Else
		$FileHolidays_hwd = FileOpen($HolidaysFile, 0)
		If $FileHolidays_hwd = -1 Then

		Else

			While 1
				$HolidaysLine = FileReadLine($FileHolidays_hwd)
				If @error = -1 Then ExitLoop
				If @error = 1 Then
					MsgBox(262160, "Import", "Oops! Something went wrong when read the file. Please try again." & @CRLF & "Error code: " & @error)
					Return
				EndIf
				If StringInStr($HolidaysLine, "-") Then
					$HolidaysLineSplited = StringSplit($HolidaysLine, "-")
					If _DateIsValid($HolidaysLineSplited[1]) Then
						$HolidaysDateSplited = StringSplit($HolidaysLineSplited[1], "/")
						If @error Then
							$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
						Else

							If $HolidaysLineSplited[2] = "O" Or $HolidaysLineSplited[2] = "R" Or $HolidaysLineSplited[2] = "B" Or $HolidaysLineSplited[2] = "T" Or $HolidaysLineSplited[2] = "P" Or $HolidaysLineSplited[2] = "H" Or $HolidaysLineSplited[2] = "S" Then

								$RegError = RegWrite($DB & "\" & $HolidaysDateSplited[1] & "\" & $HolidaysDateSplited[2], $HolidaysDateSplited[3], "REG_SZ", $HolidaysLineSplited[2] & $HolidaysLineSplited[3])
								$ImportCount += 1
							Else
								If $HolidaysLineSplited[2] = "W" Then
									$DayofWeek = _DateToDayOfWeek($HolidaysDateSplited[1], $HolidaysDateSplited[2], $HolidaysDateSplited[3])
									If $DayofWeek = "1" Or $DayofWeek = "7" Then
										$RegError = RegWrite($DB & "\" & $HolidaysDateSplited[1] & "\" & $HolidaysDateSplited[2], $HolidaysDateSplited[3], "REG_SZ", $HolidaysLineSplited[2] & $HolidaysLineSplited[3])
										$ImportCount += 1
									Else
										$HolidaysError = $HolidaysError & $HolidaysLine & " - DATE IS NOT A WEEKEND" & @CRLF
									EndIf
								Else
									If $HolidaysLineSplited[2] <> "" Then
										$HolidaysError = $HolidaysError & $HolidaysLine & " - INVALID TYPE OF OPERATION" & @CRLF
									EndIf
								EndIf
							EndIf

							If @error Then
								$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
							Else
								$HolidaysSucess = $HolidaysSucess & $HolidaysLine & @CRLF
							EndIf

						EndIf

					Else
						$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
					EndIf
				Else
					$HolidaysError = $HolidaysError & $HolidaysLine & @CRLF
				EndIf

			WEnd

			If $HolidaysError <> "" Then
				MsgBox(262160, "Import", "Oops! Something went wrong when read the file." & @CRLF & "The following lines was not imported:" & @CRLF & @CRLF & $HolidaysError & @CRLF & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess)
			Else
				If $ImportCount > 10 Then
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & @CRLF & $ImportCount & " lines imported.")
				Else
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess)
				EndIf
			EndIf
		EndIf

	EndIf

EndFunc   ;==>_ImportBatch

Func _ResetDatabase($step = "0")

	$sKey = $DB & "\"
	If $step = "0" Then
		If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(262452, "Reset Database", "**Warning!** " & @CRLF & "Are you sure you want to permanently delete all data from the database? This action cannot be undone.")
		Select
			Case $iMsgBoxAnswer = 6 ;Yes

				RegDelete($sKey)
				If @error Then
					MsgBox(262160, "Reset Database", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error)
					Return
				Else
					MsgBox(262208, "Reset Database", "**Success!** The command was executed successfully. All data has been removed.")
					Return
				EndIf

			Case $iMsgBoxAnswer = 7 ;No
				Return

		EndSelect
	Else

		RegDelete($sKey)
		If @error Then
			MsgBox(262160, "Reset Database", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error)
			Return 0
		Else
;~ 				MsgBox(262208, "Reset Database", "**Success!** The command was executed successfully. All data has been removed.")
			Return 1
		EndIf

	EndIf


	Return


EndFunc   ;==>_ResetDatabase

Func _CalendarRead()

	$SelDate = GUICtrlRead($Calendar)
	$SelDateYear = GUICtrlRead($Input_SelDate)
	$SelDate_slipt = StringSplit($SelDate, "/")
	$Input_SelDate_slipt = StringSplit($SelDateYear, "/")

	GUICtrlSetData($Group_Q1, " Q1 - " & $SelDate_slipt[1])
	GUICtrlSetData($Group_Q2x, " Q2 - " & $SelDate_slipt[1])
	GUICtrlSetData($Group_Q3, " Q3 - " & $SelDate_slipt[1])
	GUICtrlSetData($Group_Q4, " Q4 - " & $SelDate_slipt[1])

	GUICtrlSetState($Label_ratio_q1, $gui_hide)
	GUICtrlSetState($Label_Ratio_q2, $gui_hide)
	GUICtrlSetState($Label_Ratio_q3, $gui_hide)
	GUICtrlSetState($Label_Ratio_q4, $gui_hide)



	GUICtrlSetState($Input_RaTio_q1, $gui_hide)
	GUICtrlSetState($Input_RaTio_q2, $gui_hide)
	GUICtrlSetState($Input_RaTio_q3, $gui_hide)
	GUICtrlSetState($Input_RaTio_q4, $gui_hide)

	If $SelDate_slipt[1] <> $Input_SelDate_slipt[1] Then
		_CriaINI($SelDate_slipt[1])
		_ClearScreen()
		$SelDate_slipt = StringSplit($SelDate, "/")
		_ReadINI($SelDate_slipt[1])
	EndIf
	GUICtrlSetData($Input_SelDate, $SelDate)
	_CheckQuarter()

	$Status_Tip = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	ConsoleWrite($Status_Tip & " - " & StringTrimLeft($Status_Tip, 1) & " - " & $DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2] & " - " & $SelDate_slipt[3] & @CRLF)
	GUICtrlSetData($Input_Tip, StringTrimLeft($Status_Tip, 1))

	Return

EndFunc   ;==>_CalendarRead

Func _ClearScreen()

	For $J = 1 To 12
		For $i = 1 To 31
			GUICtrlDelete($Inputs[$i][$J])
		Next
	Next

EndFunc   ;==>_ClearScreen

Func _ReadStatistics($Year)

	$Counta_TD_Quarter_Q1 = 0
	$Counta_TD_Quarter_Q2 = 0
	$Counta_TD_Quarter_Q3 = 0
	$Counta_TD_Quarter_Q4 = 0

	$Counta_WD_Quarter_Q1 = 0
	$Counta_WD_Quarter_Q2 = 0
	$Counta_WD_Quarter_Q3 = 0
	$Counta_WD_Quarter_Q4 = 0

	$Counta_R_Onsite_Quarter_Q1 = 0
	$Counta_R_Onsite_Quarter_Q2 = 0
	$Counta_R_Onsite_Quarter_Q3 = 0
	$Counta_R_Onsite_Quarter_Q4 = 0

	$Counta_TD_q1 = 0
	$Counta_TD_q2 = 0
	$Counta_TD_q3 = 0
	$Counta_TD_q4 = 0

	$Counta_WD_q1 = 0
	$Counta_WD_q2 = 0
	$Counta_WD_q3 = 0
	$Counta_WD_q4 = 0

	$Counta_R_Onsite_q1 = 0
	$Counta_R_Onsite_q2 = 0
	$Counta_R_Onsite_q3 = 0
	$Counta_R_Onsite_q4 = 0

	$Ratio_R_Q1 = 0
	$Ratio_R_Q2 = 0
	$Ratio_R_Q3 = 0
	$Ratio_R_Q4 = 0

	$Ratio_Q1 = 0
	$Ratio_Q2 = 0
	$Ratio_Q3 = 0
	$Ratio_Q4 = 0

	GUICtrlSetData($Input_TD_q1, "") ;## Total Days ##
	GUICtrlSetData($Input_TD_q2, "")
	GUICtrlSetData($Input_TD_q3, "")
	GUICtrlSetData($Input_TD_q4, "")

	GUICtrlSetData($Input_WD_q1, "") ;## Work Days ##
	GUICtrlSetData($Input_WD_q2, "")
	GUICtrlSetData($Input_WD_q3, "")
	GUICtrlSetData($Input_WD_q4, "")

	GUICtrlSetData($Input_E_Onsite_q1, "") ;## Estm.On-Site ##
	GUICtrlSetData($Input_E_Onsite_q2, "")
	GUICtrlSetData($Input_E_Onsite_q3, "")
	GUICtrlSetData($Input_E_Onsite_q4, "")

	GUICtrlSetData($Input_R_Onsite_q1, "") ;## Real On-Site ##
	GUICtrlSetData($Input_R_Onsite_q2, "")
	GUICtrlSetData($Input_R_Onsite_q3, "")
	GUICtrlSetData($Input_R_Onsite_q4, "")

	GUICtrlSetData($Input_Remaining_q1, "") ;## Remaining ##
	GUICtrlSetData($Input_Remaining_q2, "")
	GUICtrlSetData($Input_Remaining_q3, "")
	GUICtrlSetData($Input_Remaining_q4, "")

	GUICtrlSetData($Input_RT_q1, "") ; ## Ration ##
	GUICtrlSetData($Input_RT_q2, "")
	GUICtrlSetData($Input_RT_q3, "")
	GUICtrlSetData($Input_RT_q4, "")

	GUICtrlSetBkColor($Input_RT_q1, 0xFFFFFF)
	GUICtrlSetBkColor($Input_RT_q2, 0xFFFFFF)
	GUICtrlSetBkColor($Input_RT_q3, 0xFFFFFF)
	GUICtrlSetBkColor($Input_RT_q4, 0xFFFFFF)

	; Criar ListView com colunas para os dias do mês
	$Headers = ""
	For $i = 1 To 31
		$Headers &= "|" & $i
	Next

	; Criar Inputs para cabeçalhos (dias do mês)
;~ 	$LabelMonth[0] = GUICtrlCreateLabel("", 8, 216, 50, 20)
	For $i = 1 To 31
		If $i < 10 Then
			$n = "0" & $i
		Else
			$n = $i
		EndIf
;~ 		$LabelMonth[$i] = GUICtrlCreateLabel($n, 8 + ($i * 35), 216, 30, 25, $SS_CENTER)
	Next
	$C = 0
	$Skip = 0
	For $J = 1 To 12
		If $J < 10 Then
			$X = "0" & $J
		Else
			$X = $J
		EndIf

		For $i = 1 To 31
;~
			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf
			$IniSection[$J][$i] = RegEnumVal($DB & "\" & $Year & "\" & $X, $n)
			If @error Then ExitLoop
		Next

		$Return = _DateToMonth($X, 1)

		If @error Then ContinueLoop ; Se a seção não existir, pula para o próximo mês

		;Days
		For $i = 1 To 31

			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf

			If _DateIsValid($Year & "/" & $X & "/" & $i) = 1 Then

				$WeekDayNum = _DateToDayOfWeek($Year, $X, $i)
				$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
				$Status1 = RegRead($DB & "\" & $Year & "\" & $X, $n)
				$Status = StringLeft($Status1, 1)
				If StringLen($Status1) > 1 Then
					$tip = " - " & StringTrimLeft($Status1, 1)
				Else
					$tip = ""
				EndIf

				If $J = "01" Or $J = "02" Or $J = "03" Then
					$Counta_TD_q1 += 1

					If $Year = @YEAR Then
						If $X = @MON Then
							If $i < @MDAY Or $i = @MDAY Then
								$Counta_TD_Quarter_Q1 += 1
							EndIf
						Else
							If $X < @MON Then
								$Counta_TD_Quarter_Q1 += 1
							EndIf
						EndIf
					EndIf

				EndIf

				If $J = "04" Or $J = "05" Or $J = "06" Then
					$Counta_TD_q2 += 1

					If $Year = @YEAR Then
						If $X = @MON Then
							If $i < @MDAY Or $i = @MDAY Then
								$Counta_TD_Quarter_Q2 += 1
							EndIf
						Else
							If $X < @MON Then
								$Counta_TD_Quarter_Q2 += 1
							EndIf
						EndIf
					EndIf

				EndIf

				If $J = "07" Or $J = "08" Or $J = "09" Then
					$Counta_TD_q3 += 1
					If $Year = @YEAR Then
						If $X = @MON Then
							If $i < @MDAY Or $i = @MDAY Then
								$Counta_TD_Quarter_Q3 += 1
							EndIf
						Else
							If $X < @MON Then
								$Counta_TD_Quarter_Q3 += 1
							EndIf
						EndIf
					EndIf
				EndIf

				If $J = "10" Or $J = "11" Or $J = "12" Then
					$Counta_TD_q4 += 1

					If $Year = @YEAR Then
						If $X = @MON Then
							If $i < @MDAY Or $i = @MDAY Then
								$Counta_TD_Quarter_Q4 += 1
							EndIf
						Else
							If $X < @MON Then
								$Counta_TD_Quarter_Q4 += 1
							EndIf
						EndIf
					EndIf

				EndIf

				If $Status = "O" Then
					If $J = "01" Or $J = "02" Or $J = "03" Then
						$Counta_WD_q1 += 1
						$Counta_R_Onsite_q1 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q1 += 1
									$Counta_R_Onsite_Quarter_Q1 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q1 += 1
									$Counta_R_Onsite_Quarter_Q1 += 1
								EndIf
							EndIf
						EndIf

					EndIf

					If $J = "04" Or $J = "05" Or $J = "06" Then
						$Counta_WD_q2 += 1
						$Counta_R_Onsite_q2 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q2 += 1
									$Counta_R_Onsite_Quarter_Q2 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q2 += 1
									$Counta_R_Onsite_Quarter_Q2 += 1
								EndIf
							EndIf
						EndIf

					EndIf

					If $J = "07" Or $J = "08" Or $J = "09" Then
						$Counta_WD_q3 += 1
						$Counta_R_Onsite_q3 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q3 += 1
									$Counta_R_Onsite_Quarter_Q3 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q3 += 1
									$Counta_R_Onsite_Quarter_Q3 += 1
								EndIf
							EndIf
						EndIf
					EndIf

					If $J = "10" Or $J = "11" Or $J = "12" Then
						$Counta_WD_q4 += 1
						$Counta_R_Onsite_q4 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q4 += 1
									$Counta_R_Onsite_Quarter_Q4 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q4 += 1
									$Counta_R_Onsite_Quarter_Q4 += 1
								EndIf
							EndIf
						EndIf
					EndIf
				EndIf
				If $Status = "R" Then
					If $J = "01" Or $J = "02" Or $J = "03" Then
						$Counta_WD_q1 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q1 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q1 += 1
								EndIf
							EndIf
						EndIf
					EndIf
					If $J = "04" Or $J = "05" Or $J = "06" Then
						$Counta_WD_q2 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q2 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q2 += 1
								EndIf
							EndIf
						EndIf
					EndIf
					If $J = "07" Or $J = "08" Or $J = "09" Then
						$Counta_WD_q3 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q3 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q3 += 1
								EndIf
							EndIf
						EndIf
					EndIf
					If $J = "10" Or $J = "11" Or $J = "12" Then
						$Counta_WD_q4 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q4 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q4 += 1
								EndIf
							EndIf
						EndIf
					EndIf
				EndIf
				If $Status = "T" Then
					If $J = "01" Or $J = "02" Or $J = "03" Then
						$Counta_WD_q1 += 1
						$Counta_R_Onsite_q1 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q1 += 1
									$Counta_R_Onsite_Quarter_Q1 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q1 += 1
									$Counta_R_Onsite_Quarter_Q1 += 1
								EndIf
							EndIf
						EndIf
					EndIf
					If $J = "04" Or $J = "05" Or $J = "06" Then
						$Counta_WD_q2 += 1
						$Counta_R_Onsite_q2 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q2 += 1
									$Counta_R_Onsite_Quarter_Q2 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q2 += 1
									$Counta_R_Onsite_Quarter_Q2 += 1
								EndIf
							EndIf
						EndIf
					EndIf

					If $J = "07" Or $J = "08" Or $J = "09" Then
						$Counta_WD_q3 += 1
						$Counta_R_Onsite_q3 += 1
						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q3 += 1
									$Counta_R_Onsite_Quarter_Q3 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q3 += 1
									$Counta_R_Onsite_Quarter_Q3 += 1
								EndIf
							EndIf
						EndIf

					EndIf

					If $J = "10" Or $J = "11" Or $J = "12" Then
						$Counta_WD_q4 += 1
						$Counta_R_Onsite_q4 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q4 += 1
									$Counta_R_Onsite_Quarter_Q4 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q4 += 1
									$Counta_R_Onsite_Quarter_Q4 += 1
								EndIf
							EndIf
						EndIf
					EndIf
				EndIf
				If $Status = "" Then
					If $J = "01" Or $J = "02" Or $J = "03" Then
						$Counta_WD_q1 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q1 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q1 += 1
								EndIf
							EndIf
						EndIf
					EndIf

					If $J = "04" Or $J = "05" Or $J = "06" Then
						$Counta_WD_q2 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q2 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q2 += 1
								EndIf
							EndIf
						EndIf

					EndIf

					If $J = "07" Or $J = "08" Or $J = "09" Then
						$Counta_WD_q3 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q3 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q3 += 1
								EndIf
							EndIf
						EndIf

					EndIf

					If $J = "10" Or $J = "11" Or $J = "12" Then
						$Counta_WD_q4 += 1

						If $Year = @YEAR Then
							If $X = @MON Then
								If $i < @MDAY Or $i = @MDAY Then
									$Counta_WD_Quarter_Q4 += 1
								EndIf
							Else
								If $X < @MON Then
									$Counta_WD_Quarter_Q4 += 1
								EndIf
							EndIf
						EndIf

					EndIf
				EndIf

			EndIf
		Next

		$C += 1
		If $C > 2 Then
			$C = 0
			$Skip = $Skip + 10
		EndIf

	Next

	GUICtrlSetData($Input_TD_q1, $Counta_TD_q1) ;## Total Days ##
	GUICtrlSetData($Input_TD_q2, $Counta_TD_q2)
	GUICtrlSetData($Input_TD_q3, $Counta_TD_q3)
	GUICtrlSetData($Input_TD_q4, $Counta_TD_q4)

	GUICtrlSetData($Input_WD_q1, $Counta_WD_q1) ;## Work Days ##
	GUICtrlSetData($Input_WD_q2, $Counta_WD_q2)
	GUICtrlSetData($Input_WD_q3, $Counta_WD_q3)
	GUICtrlSetData($Input_WD_q4, $Counta_WD_q4)

	GUICtrlSetData($Input_E_Onsite_q1, ($Counta_WD_q1 / 5) * 3) ;## Estm.On-Site ##
	GUICtrlSetData($Input_E_Onsite_q2, ($Counta_WD_q2 / 5) * 3)
	GUICtrlSetData($Input_E_Onsite_q3, ($Counta_WD_q3 / 5) * 3)
	GUICtrlSetData($Input_E_Onsite_q4, ($Counta_WD_q4 / 5) * 3)

	GUICtrlSetData($Input_R_Onsite_q1, $Counta_R_Onsite_q1) ;## Real On-Site ##
	GUICtrlSetData($Input_R_Onsite_q2, $Counta_R_Onsite_q2)
	GUICtrlSetData($Input_R_Onsite_q3, $Counta_R_Onsite_q3)
	GUICtrlSetData($Input_R_Onsite_q4, $Counta_R_Onsite_q4)

	GUICtrlSetData($Input_Remaining_q1, Round((($Counta_WD_q1 / 5) * 3) - $Counta_R_Onsite_q1, 1)) ;## Remaining ##
	GUICtrlSetData($Input_Remaining_q2, Round((($Counta_WD_q2 / 5) * 3) - $Counta_R_Onsite_q2, 1))
	GUICtrlSetData($Input_Remaining_q3, Round((($Counta_WD_q3 / 5) * 3) - $Counta_R_Onsite_q3, 1))
	GUICtrlSetData($Input_Remaining_q4, Round((($Counta_WD_q4 / 5) * 3) - $Counta_R_Onsite_q4, 1))

	$Ratio_R_Q1 = Round(($Counta_R_Onsite_q1 / ($Counta_WD_q1 / 5)), 2)
	$Ratio_R_Q2 = Round(($Counta_R_Onsite_q2 / ($Counta_WD_q2 / 5)), 2)
	$Ratio_R_Q3 = Round(($Counta_R_Onsite_q3 / ($Counta_WD_q3 / 5)), 2)
	$Ratio_R_Q4 = Round(($Counta_R_Onsite_q4 / ($Counta_WD_q4 / 5)), 2)

	GUICtrlSetData($Input_RT_q1, $Ratio_R_Q1) ; ## Ration ##
	If $Ratio_R_Q1 > 3 Or $Ratio_R_Q1 = 3 Then
		GUICtrlSetBkColor($Input_RT_q1, 0x00CC66)
	Else
		If $Ratio_R_Q1 = 0 Then
			GUICtrlSetBkColor($Input_RT_q1, 0xFFFFFF)
		Else
			GUICtrlSetBkColor($Input_RT_q1, 0xFF9933)
		EndIf
	EndIf

	GUICtrlSetData($Input_RT_q2, $Ratio_R_Q2)
	If $Ratio_R_Q2 > 3 Or $Ratio_R_Q2 = 3 Then
		GUICtrlSetBkColor($Input_RT_q2, 0x00CC66)
	Else
		If $Ratio_R_Q2 = 0 Then
			GUICtrlSetBkColor($Input_RT_q2, 0xFFFFFF)
		Else
			GUICtrlSetBkColor($Input_RT_q2, 0xFF9933)
		EndIf
	EndIf

	GUICtrlSetData($Input_RT_q3, $Ratio_R_Q3)
	If $Ratio_R_Q3 > 3 Or $Ratio_R_Q3 = 3 Then
		GUICtrlSetBkColor($Input_RT_q3, 0x00CC66)
	Else
		If $Ratio_R_Q3 = 0 Then
			GUICtrlSetBkColor($Input_RT_q3, 0xFFFFFF)
		Else
			GUICtrlSetBkColor($Input_RT_q3, 0xFF9933)
		EndIf
	EndIf

	GUICtrlSetData($Input_RT_q4, $Ratio_R_Q4)
	If $Ratio_R_Q4 > 3 Or $Ratio_R_Q4 = 3 Then
		GUICtrlSetBkColor($Input_RT_q4, 0x00CC66)
	Else
		If $Ratio_R_Q4 = 0 Then
			GUICtrlSetBkColor($Input_RT_q4, 0xFFFFFF)
		Else
			GUICtrlSetBkColor($Input_RT_q4, 0xFF9933)
		EndIf
	EndIf

	$Ratio_Q1 = Round(($Counta_R_Onsite_Quarter_Q1 / ($Counta_WD_Quarter_Q1 / 5)), 2)
	$Ratio_Q2 = Round(($Counta_R_Onsite_Quarter_Q2 / ($Counta_WD_Quarter_Q2 / 5)), 2)
	$Ratio_Q3 = Round(($Counta_R_Onsite_Quarter_Q3 / ($Counta_WD_Quarter_Q3 / 5)), 2)
	$Ratio_Q4 = Round(($Counta_R_Onsite_Quarter_Q4 / ($Counta_WD_Quarter_Q4 / 5)), 2)

	GUICtrlSetData($Input_RaTio_q1, "")
	GUICtrlSetData($Input_RaTio_q2, "")
	GUICtrlSetData($Input_RaTio_q3, "")
	GUICtrlSetData($Input_RaTio_q4, "")


	If $Year = @YEAR Then
		If @MON = "01" Or @MON = "02" Or @MON = "03" Then
			GUICtrlSetData($Input_RaTio_q1, $Ratio_Q1)
			If $Ratio_Q1 > 3 Or $Ratio_Q1 = 3 Then
				GUICtrlSetBkColor($Input_RaTio_q1, 0x00CC66)
			Else

				GUICtrlSetBkColor($Input_RaTio_q1, 0xFF9933)
			EndIf
		EndIf

		If @MON = "04" Or @MON = "05" Or @MON = "06" Then
			GUICtrlSetData($Input_RaTio_q2, $Ratio_Q2)
			If $Ratio_Q2 > 3 Or $Ratio_Q2 = 3 Then
				GUICtrlSetBkColor($Input_RaTio_q2, 0x00CC66)
			Else
				GUICtrlSetBkColor($Input_RaTio_q2, 0xFF9933)
			EndIf
		EndIf

		If @MON = "07" Or @MON = "08" Or @MON = "09" Then
			GUICtrlSetData($Input_RaTio_q3, $Ratio_Q3)
			If $Ratio_Q3 > 3 Or $Ratio_Q3 = 3 Then
				GUICtrlSetBkColor($Input_RaTio_q3, 0x00CC66)
			Else
				GUICtrlSetBkColor($Input_RaTio_q3, 0xFF9933)
			EndIf
		EndIf

		If @MON = "10" Or @MON = "11" Or @MON = "12" Then
			GUICtrlSetData($Input_RaTio_q4, $Ratio_Q4)
			If $Ratio_Q4 > 3 Or $Ratio_Q4 = 3 Then
				GUICtrlSetBkColor($Input_RaTio_q4, 0x00CC66)
			Else
				GUICtrlSetBkColor($Input_RaTio_q4, 0xFF9933)
			EndIf

		EndIf
	EndIf

;~ 	MsgBox(262144, "", "$Counta_TD_Quarter_Q1: " & $Counta_TD_Quarter_Q1 & @CRLF & "$Counta_WD_Quarter_Q1: " & $Counta_WD_Quarter_Q1 & @CRLF & "$Counta_R_Onsite_Quarter_Q1: " & $Counta_R_Onsite_Quarter_Q1 & @CRLF & "Ratio_Q1: " & Round(($Counta_R_Onsite_Quarter_Q1 / ($Counta_WD_Quarter_Q1 / 5)), 2))
;~ 	MsgBox(262144, "", "$Counta_TD_Quarter_Q2: " & $Counta_TD_Quarter_Q2 & @CRLF & "$Counta_WD_Quarter_Q2: " & $Counta_WD_Quarter_Q2 & @CRLF & "$Counta_R_Onsite_Quarter_Q2: " & $Counta_R_Onsite_Quarter_Q2 & @CRLF & "Ratio_Q2: " & Round(($Counta_R_Onsite_Quarter_Q2 / ($Counta_WD_Quarter_Q2 / 5)), 2))
;~ 	MsgBox(262144, "", "$Counta_TD_Quarter_Q3: " & $Counta_TD_Quarter_Q3 & @CRLF & "$Counta_WD_Quarter_Q3: " & $Counta_WD_Quarter_Q3 & @CRLF & "$Counta_R_Onsite_Quarter_Q3: " & $Counta_R_Onsite_Quarter_Q3 & @CRLF & "Ratio_Q3: " & Round(($Counta_R_Onsite_Quarter_Q3 / ($Counta_WD_Quarter_Q3 / 5)), 2))
;~ 	MsgBox(262144, "", "$Counta_TD_Quarter_Q4: " & $Counta_TD_Quarter_Q4 & @CRLF & "$Counta_WD_Quarter_Q4: " & $Counta_WD_Quarter_Q4 & @CRLF & "$Counta_R_Onsite_Quarter_Q4: " & $Counta_R_Onsite_Quarter_Q4 & @CRLF & "Ratio_Q4: " & Round(($Counta_R_Onsite_Quarter_Q4 / ($Counta_WD_Quarter_Q4 / 5)), 2))
	Return

EndFunc   ;==>_ReadStatistics

Func _ReadINI($Year)

	_ClearScreen()

	_ReadStatistics($Year)

	; Criar ListView com colunas para os dias do mês
	$Headers = ""
	For $i = 1 To 31
		$Headers &= "|" & $i
	Next

	; Criar Inputs para cabeçalhos (dias do mês)
	$LabelMonth[0] = GUICtrlCreateLabel("", 8, 216, 50, 20)
	For $i = 1 To 31
		If $i < 10 Then
			$n = "0" & $i
		Else
			$n = $i
		EndIf
		$LabelMonth[$i] = GUICtrlCreateLabel($n, 8 + ($i * 35), 216, 30, 25, $SS_CENTER)
	Next
	$C = 0
	$Skip = 0
	For $J = 1 To 12
		If $J < 10 Then
			$X = "0" & $J
		Else
			$X = $J
		EndIf

		For $i = 1 To 31
;~
			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf
			$IniSection[$J][$i] = RegEnumVal($DB & "\" & $Year & "\" & $X, $n)
			If @error Then ExitLoop
		Next


		$Return = _DateToMonth($X, 1)

		If @error Then ContinueLoop ; Se a seção não existir, pula para o próximo mês

		; Month
		$LabelMonthX[$J] = GUICtrlCreateLabel($Return, 8, 216 + $Skip + ($J * 25), 50, 20)

		;Days
		For $i = 1 To 31

			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf

			If _DateIsValid($Year & "/" & $X & "/" & $i) = 1 Then

				$Inputs[$i][$J] = GUICtrlCreateButton("", 8 + ($i * 35), 216 + $Skip + ($J * 25), 30, 25, BitOR($ES_READONLY, $ES_CENTER, $BS_FLAT))
				$WeekDayNum = _DateToDayOfWeek($Year, $X, $i)
				$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
				$Status1 = RegRead($DB & "\" & $Year & "\" & $X, $n)
				$Status = StringLeft($Status1, 1)
				If StringLen($Status1) > 1 Then
					$tip = " - " & StringTrimLeft($Status1, 1)
					GUICtrlSetData($Input_Tip, StringTrimLeft($Status1, 1))
				Else
					$tip = ""
					GUICtrlSetData($Input_Tip, $tip)
				EndIf


				GUICtrlSetTip($Inputs[$i][$J], $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & $tip)

				GUICtrlSetData($Inputs[$i][$J], $Status)
				If $Status = "W" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_Weekend) ; Weekend

					$Font_Weekend = $Black
					If $Picker_Font_Weekend_Read = 1 Then
						$Font_Weekend = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_Weekend)

				EndIf

				If $Status = "O" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_OnSite) ; On-site

					$Font_OnSite = $Black
					If $Picker_Font_OnSite_Read = 1 Then
						$Font_OnSite = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_OnSite)

				EndIf

				If $Status = "R" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_Remote) ; Remote

					$Font_Remote = $Black
					If $Picker_Font_Remote_Read = 1 Then
						$Font_Remote = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_Remote)


				EndIf

				If $Status = "T" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_Travel) ; Travel

					$Font_Travel = $Black
					If $Picker_Font_Travel_Read = 1 Then
						$Font_Travel = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_Travel)

				EndIf

				If $Status = "P" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_PTO) ; PTO

					$Font_PTO = $Black
					If $Picker_Font_PTO_Read = 1 Then
						$Font_PTO = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_PTO)

				EndIf

				If $Status = "H" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_holiday) ; holiday

					$Font_Holiday = $Black
					If $Picker_Font_Holiday_Read = 1 Then
						$Font_Holiday = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_Holiday)

				EndIf

				If $Status = "" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_Blank) ; Weekend

					$Font_Blank = $Black
					If $Picker_Font_Blank_Read = 1 Then
						$Font_Blank = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_Blank)

				EndIf

				If $Status = "S" Then
					GUICtrlSetBkColor($Inputs[$i][$J], $Color_bk_Sick) ; Sick

					$Font_Sick = $Black
					If $Picker_Font_Sick_Read = 1 Then
						$Font_Sick = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$J], $Font_Sick)

				EndIf

				If $Year & "/" & $X & "/" & $n = @YEAR & "/" & @MON & "/" & @MDAY Then

					GUICtrlSetColor($Inputs[$i][$J], 0xFFCCCC) ; Sick
					GUICtrlSetFont($Inputs[$i][$J], 10, 1200, "", "", 5)

					GUICtrlSetTip($Inputs[$i][$J], $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & " - TODAY" & $tip)
					If $Status = "" Then
						$Status = "X"
						GUICtrlSetColor($Inputs[$i][$J], 0xFF0000) ; today
					EndIf

					GUICtrlSetData($Inputs[$i][$J], "<" & $Status & ">")

				EndIf

			EndIf
		Next

		$C += 1
		If $C > 2 Then
			$C = 0
			$Skip = $Skip + 10
		EndIf

	Next


	Return

EndFunc   ;==>_ReadINI

Func _CheckQuarter()

	$SelDate = GUICtrlRead($Calendar)
	$Color_bk_Black = 0x000000

	$SelDate_slipt = StringSplit($SelDate, "/")
	If $SelDate_slipt[1] = @YEAR Then

		If @MON = "01" Or @MON = "02" Or @MON = "03" Then
			GUICtrlSetData($Label_3_q1, "Ratio:")
			GUICtrlSetColor($Label_3_q1, $Color_bk_Black)
			GUICtrlSetData($Label_3_q2, "Ratio:")
			GUICtrlSetColor($Label_3_q2, $Color_bk_Black)
			GUICtrlSetData($Label_3_q3, "Ratio:")
			GUICtrlSetColor($Label_3_q3, $Color_bk_Black)
			GUICtrlSetData($Label_3_q4, "Ratio:")
			GUICtrlSetColor($Label_3_q4, $Color_bk_Black)

			GUICtrlSetState($Label_ratio_q1, $gui_show)
			GUICtrlSetState($Label_Ratio_q2, $gui_hide)
			GUICtrlSetState($Label_Ratio_q3, $gui_hide)
			GUICtrlSetState($Label_Ratio_q4, $gui_hide)

			GUICtrlSetState($Input_RaTio_q1, $gui_show)
			GUICtrlSetState($Input_RaTio_q2, $gui_hide)
			GUICtrlSetState($Input_RaTio_q3, $gui_hide)
			GUICtrlSetState($Input_RaTio_q4, $gui_hide)
		EndIf

		If @MON = "04" Or @MON = "05" Or @MON = "06" Then
			GUICtrlSetData($Label_3_q1, "Ratio:")
			GUICtrlSetColor($Label_3_q1, $Color_bk_Black)
			GUICtrlSetData($Label_3_q2, "Ratio:")
			GUICtrlSetColor($Label_3_q2, $Color_bk_Black)
			GUICtrlSetData($Label_3_q3, "Ratio:")
			GUICtrlSetColor($Label_3_q3, $Color_bk_Black)
			GUICtrlSetData($Label_3_q4, "Ratio:")
			GUICtrlSetColor($Label_3_q4, $Color_bk_Black)

			GUICtrlSetState($Label_ratio_q1, $gui_hide)
			GUICtrlSetState($Label_Ratio_q2, $gui_show)
			GUICtrlSetState($Label_Ratio_q3, $gui_hide)
			GUICtrlSetState($Label_Ratio_q4, $gui_hide)

			GUICtrlSetState($Input_RaTio_q1, $gui_hide)
			GUICtrlSetState($Input_RaTio_q2, $gui_show)
			GUICtrlSetState($Input_RaTio_q3, $gui_hide)
			GUICtrlSetState($Input_RaTio_q4, $gui_hide)

		EndIf

		If @MON = "07" Or @MON = "08" Or @MON = "09" Then
			GUICtrlSetData($Label_3_q1, "Ratio:")
			GUICtrlSetBkColor($Label_3_q1, $Color_bk_Black)
			GUICtrlSetData($Label_3_q2, "Ratio:")
			GUICtrlSetBkColor($Label_3_q2, $Color_bk_Black)
			GUICtrlSetData($Label_3_q3, "Ratio:")
			GUICtrlSetBkColor($Label_3_q3, $Color_bk_Black)
			GUICtrlSetData($Label_3_q4, "Ratio:")
			GUICtrlSetBkColor($Label_3_q4, $Color_bk_Black)

			GUICtrlSetState($Label_ratio_q1, $gui_hide)
			GUICtrlSetState($Label_Ratio_q2, $gui_hide)
			GUICtrlSetState($Label_Ratio_q3, $gui_show)
			GUICtrlSetState($Label_Ratio_q4, $gui_hide)

			GUICtrlSetState($Input_RaTio_q1, $gui_hide)
			GUICtrlSetState($Input_RaTio_q2, $gui_hide)
			GUICtrlSetState($Input_RaTio_q3, $gui_show)
			GUICtrlSetState($Input_RaTio_q4, $gui_hide)

		EndIf

		If @MON = "10" Or @MON = "11" Or @MON = "12" Then
			GUICtrlSetData($Label_3_q1, "Ratio:")
			GUICtrlSetBkColor($Label_3_q1, $Color_bk_Black)
			GUICtrlSetData($Label_3_q2, "Ratio:")
			GUICtrlSetBkColor($Label_3_q2, $Color_bk_Black)
			GUICtrlSetData($Label_3_q3, "Ratio:")
			GUICtrlSetBkColor($Label_3_q3, $Color_bk_Black)
			GUICtrlSetData($Label_3_q4, "Ratio:")
			GUICtrlSetBkColor($Label_3_q4, $Color_bk_Black)

			GUICtrlSetState($Label_ratio_q1, $gui_hide)
			GUICtrlSetState($Label_Ratio_q2, $gui_hide)
			GUICtrlSetState($Label_Ratio_q3, $gui_hide)
			GUICtrlSetState($Label_Ratio_q4, $gui_show)

			GUICtrlSetState($Input_RaTio_q1, $gui_hide)
			GUICtrlSetState($Input_RaTio_q2, $gui_hide)
			GUICtrlSetState($Input_RaTio_q3, $gui_hide)
			GUICtrlSetState($Input_RaTio_q4, $gui_show)

		EndIf

	Else
		GUICtrlSetData($Label_3_q1, "Ratio:")
		GUICtrlSetColor($Label_3_q1, $Color_bk_Black)
		GUICtrlSetData($Label_3_q2, "Ratio:")
		GUICtrlSetColor($Label_3_q2, $Color_bk_Black)
		GUICtrlSetData($Label_3_q3, "Ratio:")
		GUICtrlSetColor($Label_3_q3, $Color_bk_Black)
		GUICtrlSetData($Label_3_q4, "Ratio:")
		GUICtrlSetColor($Label_3_q4, $Color_bk_Black)

		GUICtrlSetState($Label_ratio_q1, $gui_hide)
		GUICtrlSetState($Label_Ratio_q2, $gui_hide)
		GUICtrlSetState($Label_Ratio_q3, $gui_hide)
		GUICtrlSetState($Label_Ratio_q4, $gui_hide)

		GUICtrlSetState($Input_RaTio_q1, $gui_hide)
		GUICtrlSetState($Input_RaTio_q2, $gui_hide)
		GUICtrlSetState($Input_RaTio_q3, $gui_hide)
		GUICtrlSetState($Input_RaTio_q4, $gui_hide)

	EndIf

	If $SelDate_slipt[2] = "01" Or $SelDate_slipt[2] = "02" Or $SelDate_slipt[2] = "03" Then
		GUICtrlSetData($Input_Quarter, "Q1")
		GUICtrlSetBkColor($Group_Q1, 0x00FF00)

	Else
		GUICtrlSetBkColor($Group_Q1, 0xC0C0C0)
	EndIf

	If $SelDate_slipt[2] = "04" Or $SelDate_slipt[2] = "05" Or $SelDate_slipt[2] = "06" Then
		GUICtrlSetData($Input_Quarter, "Q2")
		GUICtrlSetBkColor($Group_Q2x, 0x00FF00)
	Else
		GUICtrlSetBkColor($Group_Q2x, 0xC0C0C0)
	EndIf

	If $SelDate_slipt[2] = "07" Or $SelDate_slipt[2] = "08" Or $SelDate_slipt[2] = "09" Then
		GUICtrlSetData($Input_Quarter, "Q3")
		GUICtrlSetBkColor($Group_Q3, 0x00FF00)
	Else
		GUICtrlSetBkColor($Group_Q3, 0xC0C0C0)
	EndIf


	If $SelDate_slipt[2] = "10" Or $SelDate_slipt[2] = "11" Or $SelDate_slipt[2] = "12" Then
		GUICtrlSetData($Input_Quarter, "Q4")
		GUICtrlSetBkColor($Group_Q4, 0x00FF00)
	Else
		GUICtrlSetBkColor($Group_Q4, 0xC0C0C0)
	EndIf

	Return


EndFunc   ;==>_CheckQuarter

Func _CriaINI($Year)

	Local $sJulDate1 = _DateToDayValue($Year, "12", "31")
	For $i = 0 To 365 Step 1

		Local $Y, $M, $D
		$sJulDate = _DayValueToDate($sJulDate1 - $i, $Y, $M, $D)
		If $Y = $Year Then


			$Wday = _DateToDayOfWeek($Year, $M, $D)
			If $Wday = 1 Or $Wday = 7 Then
				If RegRead($DB & "\" & $Year & "\" & $M, $D) = "" Then
					RegWrite($DB & "\" & $Year & "\" & $M, $D, "REG_SZ", "W")
				EndIf
			Else
				RegWrite($DB & "\" & $Year & "\" & $M, $D, "REG_SZ", RegRead($DB & "\" & $Year & "\" & $M, $D))
			EndIf

		EndIf
	Next
	Return

EndFunc   ;==>_CriaINI

Func old_CreateBackup($DBBKP = "")

	Local $sRegPath = $DB & "\"
	If $DBBKP = "" Then
		Local $sFilePath = FileSaveDialog("Save backup file", @ScriptDir, "All (*.*)", 18)
		If @error Then
;~ 			MsgBox(262144 + 16, "Error", "Operation aborted.")

			Return

		EndIf
	Else
		$sFilePath = $DBBKP

	EndIf

	$sFilePath_hwd = FileOpen($sFilePath, 10)

	Local $sSubKey = ""

	; Loop from 1 to 10 times, displaying registry keys at the particular instance value.
	For $i = 1 To 10000
		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

		ConsoleWrite($DB & "\" & $sSubKey & @CRLF)


		For $r = 1 To 10000
			$sSubKey_month = RegEnumKey($DB & "\" & $sSubKey, $r)
			If @error Then ExitLoop

			ConsoleWrite($DB & "\" & $sSubKey & "\" & $sSubKey_month & @CRLF)

			For $D = 1 To 10000

				If $D < 10 Then
					$D1 = "0" & $D
				Else
					$D1 = $D
				EndIf

				$sSubKey_day = RegEnumVal($DB & "\" & $sSubKey & "\" & $sSubKey_month, $D1)
				If @error Then ExitLoop
				$RegRead = RegRead($DB & "\" & $sSubKey & "\" & $sSubKey_month, $sSubKey_day)
				FileWriteLine($sFilePath_hwd, $sSubKey & "/" & $sSubKey_month & "/" & $sSubKey_day & "-" & StringLeft($RegRead, 1) & "-" & StringTrimLeft($RegRead, 1))
			Next
		Next
	Next

	FileClose($sFilePath_hwd)

	If $DBBKP = "" Then
		MsgBox(64, "Sucess", "Backup saved: " & $sFilePath)
	EndIf

	Return
EndFunc   ;==>_CreateBackup

Func _BKColorPallet()

	$Form_Colors = GUICreate('Colors', 210, 330, -1, -1, $DS_MODALFRAME, $WS_EX_TOPMOST)
	GUICtrlSetBkColor(-1, 0x50CA1B)

	GUICtrlCreateLabel("On Site:", 10, 15)
	GUICtrlCreateLabel("Remote:", 10, 45)
	GUICtrlCreateLabel("Holiday:", 10, 75)
	GUICtrlCreateLabel("PTO:", 10, 105)
	GUICtrlCreateLabel("Travel:", 10, 135)
	GUICtrlCreateLabel("Sick:", 10, 165)
	GUICtrlCreateLabel("Blank:", 10, 195)
	GUICtrlCreateLabel("Weekend:", 10, 225)

	$Picker_OnSite = _GUIColorPicker_Create('', 60, 10, 60, 23, $Color_bk_OnSite) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_Remote = _GUIColorPicker_Create('', 60, 40, 60, 23, $Color_bk_Remote) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_Holiday = _GUIColorPicker_Create('', 60, 70, 60, 23, $Color_bk_holiday) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_PTO = _GUIColorPicker_Create('', 60, 100, 60, 23, $Color_bk_PTO) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_Travel = _GUIColorPicker_Create('', 60, 130, 60, 23, $Color_bk_Travel) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_Sick = _GUIColorPicker_Create('', 60, 160, 60, 23, $Color_bk_Sick) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_Blank = _GUIColorPicker_Create('', 60, 190, 60, 23, $Color_bk_Blank) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')
	$Picker_Weekend = _GUIColorPicker_Create('', 60, 220, 60, 23, $Color_bk_Weekend) ;, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_MAGNIFICATION, $CP_FLAG_ARROWSTYLE), 0, -1, -1, 0, 'Simple Text', 'Custom...', '_ColorChooserDialog')

	$Picker_Font_OnSite = GUICtrlCreateCheckbox("White Font", 120, 10)
	$Picker_Font_Remote = GUICtrlCreateCheckbox("White Font", 120, 40)
	$Picker_Font_Holiday = GUICtrlCreateCheckbox("White Font", 120, 70)
	$Picker_Font_PTO = GUICtrlCreateCheckbox("White Font", 120, 100)
	$Picker_Font_Travel = GUICtrlCreateCheckbox("White Font", 120, 130)
	$Picker_Font_Sick = GUICtrlCreateCheckbox("White Font", 120, 160)
	$Picker_Font_Blank = GUICtrlCreateCheckbox("White Font", 120, 190)
	$Picker_Font_Weekend = GUICtrlCreateCheckbox("White Font", 120, 220)

	GUICtrlSetState($Picker_Font_OnSite, $Picker_Font_OnSite_Read)
	GUICtrlSetState($Picker_Font_Remote, $Picker_Font_Remote_Read)
	GUICtrlSetState($Picker_Font_Holiday, $Picker_Font_Holiday_Read)
	GUICtrlSetState($Picker_Font_PTO, $Picker_Font_PTO_Read)
	GUICtrlSetState($Picker_Font_Travel, $Picker_Font_Travel_Read)
	GUICtrlSetState($Picker_Font_Sick, $Picker_Font_Sick_Read)
	GUICtrlSetState($Picker_Font_Blank, $Picker_Font_Blank_Read)
	GUICtrlSetState($Picker_Font_Weekend, $Picker_Font_Weekend_Read)

	$Original_Color_1 = $Color_bk_OnSite & $Color_bk_Remote & $Color_bk_holiday & $Color_bk_PTO & $Color_bk_Travel & $Color_bk_Sick & $Color_bk_Blank & $Color_bk_Weekend & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & $Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read
	ConsoleWrite($Original_Color_1 & @CRLF)

	$Colors_Close = GUICtrlCreateButton("Close", 80, 260, 70, 30)

	GUISetState()

	While 1
		$Msg = GUIGetMsg()
		Switch $Msg
			Case $Colors_Close
				$Picker_Color_OnSite = _GUIColorPicker_GetColor($Picker_OnSite)
				$Picker_Color_Remote = _GUIColorPicker_GetColor($Picker_Remote)
				$Picker_Color_Holiday = _GUIColorPicker_GetColor($Picker_Holiday)
				$Picker_Color_PTO = _GUIColorPicker_GetColor($Picker_PTO)
				$Picker_Color_Travel = _GUIColorPicker_GetColor($Picker_Travel)
				$Picker_Color_Sick = _GUIColorPicker_GetColor($Picker_Sick)
				$Picker_Color_Blank = _GUIColorPicker_GetColor($Picker_Blank)
				$Picker_Color_Weekend = _GUIColorPicker_GetColor($Picker_Weekend)

				RegWrite($DB, "Color_OnSite", "REG_SZ", $Picker_Color_OnSite)
				RegWrite($DB, "Color_Remote", "REG_SZ", $Picker_Color_Remote)
				RegWrite($DB, "Color_holiday", "REG_SZ", $Picker_Color_Holiday)
				RegWrite($DB, "Color_PTO", "REG_SZ", $Picker_Color_PTO)
				RegWrite($DB, "Color_Travel", "REG_SZ", $Picker_Color_Travel)
				RegWrite($DB, "Color_Sick", "REG_SZ", $Picker_Color_Sick)
				RegWrite($DB, "Color_Blank", "REG_SZ", $Picker_Color_Blank)
				RegWrite($DB, "Color_Weekend", "REG_SZ", $Picker_Color_Weekend)

				$Picker_Font_OnSite_Read = GUICtrlRead($Picker_Font_OnSite)
				$Picker_Font_Remote_Read = GUICtrlRead($Picker_Font_Remote)
				$Picker_Font_Holiday_Read = GUICtrlRead($Picker_Font_Holiday)
				$Picker_Font_PTO_Read = GUICtrlRead($Picker_Font_PTO)
				$Picker_Font_Travel_Read = GUICtrlRead($Picker_Font_Travel)
				$Picker_Font_Sick_Read = GUICtrlRead($Picker_Font_Sick)
				$Picker_Font_Blank_Read = GUICtrlRead($Picker_Font_Blank)
				$Picker_Font_Weekend_Read = GUICtrlRead($Picker_Font_Weekend)

				RegWrite($DB, "Font_OnSite", "REG_SZ", $Picker_Font_OnSite_Read)
				RegWrite($DB, "Font_Remote", "REG_SZ", $Picker_Font_Remote_Read)
				RegWrite($DB, "Font_holiday", "REG_SZ", $Picker_Font_Holiday_Read)
				RegWrite($DB, "Font_PTO", "REG_SZ", $Picker_Font_PTO_Read)
				RegWrite($DB, "Font_Travel", "REG_SZ", $Picker_Font_Travel_Read)
				RegWrite($DB, "Font_Sick", "REG_SZ", $Picker_Font_Sick_Read)
				RegWrite($DB, "Font_Blank", "REG_SZ", $Picker_Font_Blank_Read)
				RegWrite($DB, "Font_Weekend", "REG_SZ", $Picker_Font_Weekend_Read)




				$Font_OnSite = $Black
				If $Picker_Font_OnSite_Read = 1 Then
					$Font_OnSite = $White
				EndIf


				$Font_Remote = $Black
				If $Picker_Font_Remote_Read = 1 Then
					$Font_Remote = $White
				EndIf


				$Font_Holiday = $Black
				If $Picker_Font_Holiday_Read = 1 Then
					$Font_Holiday = $White
				EndIf


				$Font_PTO = $Black
				If $Picker_Font_PTO_Read = 1 Then
					$Font_PTO = $White
				EndIf


				$Font_Travel = $Black
				If $Picker_Font_Travel_Read = 1 Then
					$Font_Travel = $White
				EndIf


				$Font_Sick = $Black
				If $Picker_Font_Sick_Read = 1 Then
					$Font_Sick = $White
				EndIf


				$Font_Blank = $Black
				If $Picker_Font_Blank_Read = 1 Then
					$Font_Blank = $White
				EndIf


				$Font_Weekend = $Black
				If $Picker_Font_Weekend_Read = 1 Then
					$Font_Weekend = $White
				EndIf


				GUICtrlSetColor($Button_OnSite, $Font_OnSite)
				GUICtrlSetColor($Button_Remote, $Font_Remote)
				GUICtrlSetColor($Button_holiday, $Font_Holiday)
				GUICtrlSetColor($Button_PTO, $Font_PTO)
				GUICtrlSetColor($Button_Travel, $Font_Travel)
				GUICtrlSetColor($Button_Sick, $Font_Sick)
				GUICtrlSetColor($Button_Blank, $Font_Blank)
				GUICtrlSetColor($Button_Weekend, $Font_Weekend)

				$Original_Color_2 = $Picker_Color_OnSite & $Picker_Color_Remote & $Picker_Color_Holiday & $Picker_Color_PTO & $Picker_Color_Travel & $Picker_Color_Sick & $Picker_Color_Blank & $Picker_Color_Weekend & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & $Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read

				ConsoleWrite($Original_Color_2 & @CRLF)

				GUIDelete($Form_Colors)
				If $Original_Color_1 = $Original_Color_2 Then
					Return 0
				Else
					Return 1
				EndIf



		EndSwitch
	WEnd


EndFunc   ;==>_BKColorPallet

Func _ReadColors()


	Global $CalendarTag = RegRead($DB, "caltag")
	If @error Then $CalendarTag = "1"

	Global $Color_bk_OnSite = RegRead($DB, "Color_OnSite")
	If @error Then $Color_bk_OnSite = 0x00CC66


	Global $Color_bk_Remote = RegRead($DB, "Color_Remote")
	If @error Then $Color_bk_Remote = 0x0080FF

	Global $Color_bk_holiday = RegRead($DB, "Color_holiday")
	If @error Then $Color_bk_holiday = 0xFFFFCC

	Global $Color_bk_PTO = RegRead($DB, "Color_PTO")
	If @error Then $Color_bk_PTO = 0x66FFFF

	Global $Color_bk_Travel = RegRead($DB, "Color_Travel")
	If @error Then $Color_bk_Travel = 0xFF8000

	Global $Color_bk_Sick = RegRead($DB, "Color_Sick")
	If @error Then $Color_bk_Sick = 0xFF6666

	Global $Color_bk_Blank = RegRead($DB, "Color_Blank")
	If @error Then $Color_bk_Blank = 0xFFFFFF

	Global $Color_bk_Weekend = RegRead($DB, "Color_Weekend")
	If @error Then $Color_bk_Weekend = 0xA0A0A0


	Global $Picker_Font_OnSite_Read = RegRead($DB, "Font_OnSite")
	Global $Font_OnSite = $Black
	If $Picker_Font_OnSite_Read = 1 Then
		$Font_OnSite = $White
	EndIf

	Global $Picker_Font_Remote_Read = RegRead($DB, "Font_Remote")
	Global $Font_Remote = $Black
	If $Picker_Font_Remote_Read = 1 Then
		$Font_Remote = $White
	EndIf

	Global $Picker_Font_Holiday_Read = RegRead($DB, "Font_holiday")
	Global $Font_Holiday = $Black
	If $Picker_Font_Holiday_Read = 1 Then
		$Font_Holiday = $White
	EndIf

	Global $Picker_Font_PTO_Read = RegRead($DB, "Font_PTO")
	Global $Font_PTO = $Black
	If $Picker_Font_PTO_Read = 1 Then
		$Font_PTO = $White
	EndIf

	Global $Picker_Font_Travel_Read = RegRead($DB, "Font_Travel")
	Global $Font_Travel = $Black
	If $Picker_Font_Travel_Read = 1 Then
		$Font_Travel = $White
	EndIf

	Global $Picker_Font_Sick_Read = RegRead($DB, "Font_Sick")
	Global $Font_Sick = $Black
	If $Picker_Font_Sick_Read = 1 Then
		$Font_Sick = $White
	EndIf


	Global $Picker_Font_Blank_Read = RegRead($DB, "Font_Blank")
	Global $Font_Blank = $Black
	If $Picker_Font_Blank_Read = 1 Then
		$Font_Blank = $White
	EndIf

	Global $Picker_Font_Weekend_Read = RegRead($DB, "Font_Weekend")
	Global $Font_Weekend = $Black
	If $Picker_Font_Weekend_Read = 1 Then
		$Font_Weekend = $White
	EndIf

	GUICtrlSetColor($Button_OnSite, $Font_OnSite)
	GUICtrlSetColor($Button_Remote, $Font_Remote)
	GUICtrlSetColor($Button_holiday, $Font_Holiday)
	GUICtrlSetColor($Button_PTO, $Font_PTO)
	GUICtrlSetColor($Button_Travel, $Font_Travel)
	GUICtrlSetColor($Button_Sick, $Font_Sick)
	GUICtrlSetColor($Button_Blank, $Font_Blank)
	GUICtrlSetColor($Button_Weekend, $Font_Weekend)

	GUICtrlSetBkColor($Button_OnSite, $Color_bk_OnSite)
	GUICtrlSetBkColor($Button_Remote, $Color_bk_Remote)
	GUICtrlSetBkColor($Button_holiday, $Color_bk_holiday)
	GUICtrlSetBkColor($Button_PTO, $Color_bk_PTO)
	GUICtrlSetBkColor($Button_Travel, $Color_bk_Travel)
	GUICtrlSetBkColor($Button_Sick, $Color_bk_Sick)
	GUICtrlSetBkColor($Button_Blank, $Color_bk_Blank)
	GUICtrlSetBkColor($Button_Weekend, $Color_bk_Weekend)


EndFunc   ;==>_ReadColors


Func _CreateBackup($DBBKP = "")

	Local $sRegPath = $DB & "\"

	If $DBBKP = "" Then
		Local $sFilePath = FileSaveDialog("Save backup file", @ScriptDir, "All (*.*)", 18,"Backup_" & @YEAR & "_" & @MON & "_" & @MDAY & ".bkp",$Form_WorkDays)
		If @error Then
;~ 			MsgBox(262144 + 16, "Error", "Operation aborted.")
			Return
		EndIf
	Else
		$sFilePath = $DBBKP

	EndIf

	$sFilePath_hwd = FileOpen($sFilePath, 10)

	Local $sSubKey = ""

	For $i = 1 To 100
		$sSubKey_settings = RegEnumVal($DB, $i)
		 If @error <> 0 Then ExitLoop
		$RegRead = RegRead($DB, $sSubKey_settings)
		FileWriteLine($sFilePath_hwd,$sSubKey_settings & "=" & $RegRead)
		ConsoleWrite($sSubKey_settings & "=" & $RegRead & @CRLF)
	Next

	; Loop from 1 to 10 times, displaying registry keys at the particular instance value.
	For $i = 1 To 10000
		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

		ConsoleWrite($DB & "\" & $sSubKey & @CRLF)

		For $r = 1 To 10000
			$sSubKey_month = RegEnumKey($DB & "\" & $sSubKey, $r)
			If @error Then ExitLoop

			ConsoleWrite($DB & "\" & $sSubKey & "\" & $sSubKey_month & @CRLF)

			For $D = 1 To 10000

				If $D < 10 Then
					$D1 = "0" & $D
				Else
					$D1 = $D
				EndIf

				$sSubKey_day = RegEnumVal($DB & "\" & $sSubKey & "\" & $sSubKey_month, $D1)
				If @error Then ExitLoop
				$RegRead = RegRead($DB & "\" & $sSubKey & "\" & $sSubKey_month, $sSubKey_day)
				FileWriteLine($sFilePath_hwd, $sSubKey & "\" & $sSubKey_month & "\" & $sSubKey_day & "=" & $RegRead)
			Next
		Next
	Next

	FileClose($sFilePath_hwd)

	If $DBBKP = "" Then
		MsgBox(64, "Sucess", "Backup saved: " & $sFilePath)
	EndIf

	Return

EndFunc   ;==>_CreateBackup

