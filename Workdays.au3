#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=calendar.ico
#AutoIt3Wrapper_Res_Description=Work Day management
#AutoIt3Wrapper_Res_Fileversion=1.0.2.0
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
#include <ColorChooser.au3>
#include <ColorPicker.au3>
#include <WinAPI.au3>
#include "GenerateWorkdaysReportHTML.au3"

Global $About = "1.0.1.3 - Custom colors and bug fixes" & @CRLF & "1.0.1.4 - Code polishing and new custom color palette" & @CRLF & "1.0.1.5 - Bug Fixes and improvements" & @CRLF & "1.0.1.6 - Bug Fixes and improvements" & @CRLF & "1.0.1.7 - KPI Bug Fixes" & @CRLF & "1.0.1.8 - Today custom color option" & @CRLF & "1.0.1.9 - Report Functionality" & @CRLF & "1.0.2.0 - Bug Fix"

Global $IniSection[999][999]
Global $LabelMonth[99999]
Global $LabelMonthX[99999]
Global $Inputs[32][32]
Global $TodayLabel[32][32]
Global $SelectLabel[32][32]
Global $DBpMenu_Delete_Year[20]
Global $DBpMenu_Delete_Date[15]

Global $DBpMenu_Report_Year[20]
Global $DBpMenu_Report_Date[15]

Global $Year = @YEAR
Global $Ratio_Q1 = 0
Global $Ratio_Q2 = 0
Global $Ratio_Q3 = 0
Global $Ratio_Q4 = 0

Global $Remaining_q1
Global $Remaining_q2
Global $Remaining_q3
Global $Remaining_q4

Global $Ratio_R_Q1
Global $Ratio_R_Q2
Global $Ratio_R_Q3
Global $Ratio_R_Q4

Global $White = 0xFFFFFF
Global $Black = 0x000000

$DB = "HKEY_CURRENT_USER\Software\WorkDays"


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

Global $Color_bk_Today = RegRead($DB, "Color_Today")
If @error Then $Color_bk_Today = 0xFF0000

Global $Color_bk_Selected = RegRead($DB, "Color_Selected")
If @error Then $Color_bk_Selected = 0x00F0F0

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
Global $BkpMenu_reset_all1 = GUICtrlCreateMenu("Reset Data", $DBpMenu_backup_Data)
Global $BkpMenu_reset_all = GUICtrlCreateMenuItem("Reset Entire Database", $BkpMenu_reset_all1)
Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
Global $DBpMenu_backup_Data_Holidays = GUICtrlCreateMenuItem("Import Holidays File", $DBpMenu_backup_Data)
Global $DBpMenu_Report = GUICtrlCreateMenu("Report", $DBpMenu_db)
;~ Global $DBpMenu_backup_4 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)

;~ Global $BkpMenu_reset = GUICtrlCreateMenu("Reset Data", $DBpMenu_db)
Global $BkpMenu_reset_1 = GUICtrlCreateMenuItem("", $DBpMenu_db)
Global $BkpMenu_Exit = GUICtrlCreateMenuItem("&Exit", $DBpMenu_db)

Global $DBpMenu_settings = GUICtrlCreateMenu("Settings")
Global $BkpMenu_settings_BKcolors = GUICtrlCreateMenuItem("Colors", $DBpMenu_settings)
Global $BkpMenu_help = GUICtrlCreateMenu("?")
Global $BkpMenu_help_help = GUICtrlCreateMenuItem("Help", $BkpMenu_help)
Global $BkpMenu_help_space = GUICtrlCreateMenuItem("", $BkpMenu_help)
Global $BkpMenu_help_About = GUICtrlCreateMenuItem("About", $BkpMenu_help)
;~ Global $BkpMenu_reset_year = GUICtrlCreateMenuItem("Reset Specific Year", $BkpMenu_reset)

$Calendar = GUICtrlCreateMonthCal(@YEAR & "/" & @MON & "/" & @MDAY, 8, 8, 273, 201, $MCS_WEEKNUMBERS)

$Group1 = GUICtrlCreateGroup("", 288, 8, 270, 200)

$Input_SelDate = GUICtrlCreateInput("", 376, 24, 70, 21, $ES_READONLY)
GUICtrlSetData($Input_SelDate, GUICtrlRead($Calendar))
GUICtrlSetColor($Input_SelDate, 0x990000)

$Label1 = GUICtrlCreateLabel("Selected Date:", 296, 28, 75, 17)
$Input_Quarter = GUICtrlCreateInput("", 450, 24, 20, 21, $ES_READONLY)
GUICtrlSetColor($Input_Quarter, 0x00994C)

$Input_Tag = GUICtrlCreateInput("", 296, 54, 175, 21) ;, $ES_READONLY)

$Button_CalendtarTag = GUICtrlCreateButton("Tag", 472, 52, 75, 25) ;## Calendar TAG

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
GUICtrlSetState($Button_Weekend, $gui_hide)

GUICtrlCreateLabel("Use Blank button for Weekends", 384, 180, 170)


$SelectLabel_1 = GUICtrlCreateLabel("", 494, 87, 46, 21) ;,$SS_BLACKFRAME)
$SelectLabel_2 = GUICtrlCreateLabel("", 496, 89, 42, 17) ;,$SS_BLACKFRAME)
GUICtrlSetBkColor($SelectLabel_1, $Color_bk_Today)
GUICtrlCreateLabel("Today", 497, 90, 40, 15, $SS_CENTER)

$TodayLabel_1 = GUICtrlCreateLabel("", 494, 116, 46, 21) ;,$SS_BLACKFRAME)
$TodayLabel_2 = GUICtrlCreateLabel("", 496, 118, 42, 17) ;,$SS_BLACKFRAME)
GUICtrlSetBkColor($TodayLabel_1, $Color_bk_Selected)
GUICtrlCreateLabel("Selected", 497, 119, 40, 15, $SS_CENTER)

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


GUICtrlCreateGroup("", 10, 303, 1120, 9)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlCreateGroup("", 10, 388, 1120, 9)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlCreateGroup("", 10, 473, 1120, 9)
GUICtrlSetColor(-1, 0x0000FF)

$StatusBar1 = _GUICtrlStatusBar_Create($Form_WorkDays)

$sMessage = "Developed by Fabricio Zambroni - VERSION: " & FileGetVersion(@ScriptFullPath) & " - Today: " & @YEAR & "/" & @MON & "/" & @MDAY
_GUICtrlStatusBar_SetText($StatusBar1, $sMessage)

_CriaINI(@YEAR)

_ReadINI(@YEAR)

_CheckQuarter()

_AutoBKP()

_CreateMenu()

$SelDate = GUICtrlRead($Calendar)
$SelDate_slipt = StringSplit($SelDate, "/")

$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
$Status = StringTrimLeft($Status1, 1)

GUICtrlSetData($Input_Tag, $Status)

GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

GUISetState(@SW_SHOW)


While 1
	$nMsg = GUIGetMsg()

	For $j = 1 To 12

		If $nMsg = $DBpMenu_Report_Year[$j] And $DBpMenu_Report_Year[$j] <> 0 Then
			$DBpMenu_Report_Date = GUICtrlRead($DBpMenu_Report_Year[$j], 1)
			GenerateWorkdaysReportHTML($DBpMenu_Report_Date)

		EndIf

		If $nMsg = $DBpMenu_Delete_Year[$j] And $DBpMenu_Delete_Year[$j] <> 0 Then
			$DBpMenu_Delete_Date = GUICtrlRead($DBpMenu_Delete_Year[$j], 1)
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Delete Year", "WARNING" & @CRLF & "" & @CRLF & "You are about To delete the year " & $DBpMenu_Delete_Date & " from the database." & @CRLF & "" & @CRLF & "All data associated With this year will be permanently removed And cannot be recovered." & @CRLF & "" & @CRLF & "Are you sure you want To proceed ?")
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					$BKPDB = @ScriptDir & "\autosave.db"
					_CreateBackup($BKPDB)
					$FOO = RegDelete($DB & "\" & $DBpMenu_Delete_Date)
					If Not @error Then
						If $DBpMenu_Delete_Date = @YEAR Then
							_CriaINI(@YEAR)
						EndIf
						GUICtrlSetData($Calendar, @YEAR & "/" & @MON & "/" & @MDAY)
						$SelDate = GUICtrlRead($Calendar)
						$SelDate_slipt = StringSplit($SelDate, "/")
						GUICtrlSetData($Input_SelDate, $SelDate)
						_ReadINI($SelDate_slipt[1])
						GUICtrlSetData($Calendar, @YEAR & "/" & @MON & "/" & @MDAY)
						GUICtrlSetData($Input_SelDate, $SelDate)
						_CreateMenu()
						MsgBox(262208, "Delete Year", "Year Deleted with Success")
					Else
						MsgBox(262160, "Year Delete", "An error occurred while attempting to delete this value from the database.")
					EndIf

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		EndIf

		For $i = 1 To 31
			If $Inputs[$i][$j] <> 0 And $nMsg = $Inputs[$i][$j] Then
				If $i < 10 Then
					$n = "0" & $i
				Else
					$n = $i
				EndIf

				If $j < 10 Then
					$s = "0" & $j
				Else
					$s = $j
				EndIf
				$FullDate = GUICtrlRead($Input_SelDate)
				$FullDate_Split = StringSplit($FullDate, "/")
				$ClickedDate = $FullDate_Split[1] & "/" & $s & "/" & $n

				GUICtrlSetData($Calendar, $ClickedDate)
				_CalendarRead($i, $j)

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
				If @error Then $Color_bk_Weekend = 0xF0F4A1

				$Color_bk_Today = RegRead($DB, "Color_Today")
				If @error Then $Color_bk_Today = 0xA0A0A0

				$Color_bk_Selected = RegRead($DB, "Color_Selected")
				If @error Then $Color_bk_Selected = 0x00FFA0

				GUICtrlSetBkColor($Button_OnSite, $Color_bk_OnSite)
				GUICtrlSetBkColor($Button_Remote, $Color_bk_Remote)
				GUICtrlSetBkColor($Button_holiday, $Color_bk_holiday)
				GUICtrlSetBkColor($Button_PTO, $Color_bk_PTO)
				GUICtrlSetBkColor($Button_Travel, $Color_bk_Travel)
				GUICtrlSetBkColor($Button_Sick, $Color_bk_Sick)
				GUICtrlSetBkColor($Button_Blank, $Color_bk_Blank)
				GUICtrlSetBkColor($Button_Weekend, $Color_bk_Weekend)

				GUICtrlSetBkColor($SelectLabel_1, $Color_bk_Today)
				GUICtrlSetBkColor($TodayLabel_1, $Color_bk_Selected)

				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				_ReadINI($SelDate_slipt[1])
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
				$Status = StringTrimLeft($Status1, 1)
				GUICtrlSetData($Input_Tag, $Status)
				GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)
			EndIf

		Case $BkpMenu_help_help
			$HelpFile = @ScriptDir & "\help.pdf"
			If Not FileExists($HelpFile) Then
				MsgBox(262160, "Work Days", "Help file not found in the application folder.")
			Else
				ShellExecute($HelpFile)
			EndIf

		Case $BkpMenu_help_About
			MsgBox(262144 + 64, "Work Days", "Work Days is a user-friendly calendar-based application for managing and categorizing your workdays—On Site, Remote, and Holiday—throughout the year." & @CRLF & @CRLF & $About & @CRLF & @CRLF & "Developed by Fabricio Zambroni - CURRENT VERSION: " & FileGetVersion(@ScriptFullPath))

		Case $GUI_EVENT_CLOSE
			Exit

		Case $Button_CalendtarTag
			$DateToTag = GUICtrlRead($Calendar)
			_CalendarTag($DateToTag)
			_Update($DateToTag)

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
			_ReadINI($SelDate_slipt[1])
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
			$Status = StringTrimLeft($Status1, 1)
			GUICtrlSetData($Input_Tag, $Status)
			GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

		Case $Button_OnSite
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "O")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$holidayName = GUICtrlRead($Input_Tag)
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "O" & $holidayName)
				_Update($SelDate)
			EndIf

		Case $Button_Blank
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$WeekDayNum = _DateToDayOfWeek($SelDate_slipt[1], $SelDate_slipt[2], $SelDate_slipt[3])
				$holidayName = GUICtrlRead($Input_Tag)
				If $WeekDayNum = "1" Or $WeekDayNum = "7" Then
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "W" & $holidayName)
					_Update($SelDate)
				Else
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "B" & $holidayName)
					_Update($SelDate)
				EndIf
			EndIf

		Case $Button_Remote
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "R")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$holidayName = GUICtrlRead($Input_Tag)
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "R" & $holidayName)
				_Update($SelDate)
			EndIf

		Case $Button_Travel
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "T")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$holidayName = GUICtrlRead($Input_Tag)
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "T" & $holidayName)
				_Update($SelDate)
			EndIf

		Case $Button_PTO
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "P")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$holidayName = GUICtrlRead($Input_Tag)
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "P" & $holidayName)
				_Update($SelDate)
			EndIf

		Case $Button_holiday
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "H")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$holidayName = GUICtrlRead($Input_Tag)
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "H" & $holidayName)
				_Update($SelDate)
			EndIf

		Case $Button_Sick
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "S")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")
				$holidayName = GUICtrlRead($Input_Tag)
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "S" & $holidayName)
				_Update($SelDate)
			EndIf

		Case $Button_Weekend
			$SelDate = GUICtrlRead($Calendar)
			$CheckDate_Return = _CheckDate($SelDate, "W")
			If $CheckDate_Return = 0 Then
				$SelDate_slipt = StringSplit($SelDate, "/")

				$WeekDayNum = _DateToDayOfWeek($SelDate_slipt[1], $SelDate_slipt[2], $SelDate_slipt[3])

				$WeekEnd = 0
				If $WeekDayNum <> "1" And $WeekDayNum <> "7" Then
					$WeekEnd = 1
				EndIf

				If $WeekEnd = 1 Then
					MsgBox(262160, "Weekend", "This date is not a weekend.")
				Else
					$holidayName = GUICtrlRead($Input_Tag)
					RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "W" & $holidayName)
					_Update($SelDate)
				EndIf
			EndIf

	EndSwitch

WEnd

Func _CalendarTag($DateToTag)

	$SelDate_slipt = StringSplit($DateToTag, "/")
	$holidayName = GUICtrlRead($Input_Tag)
	$Register = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", StringLeft($Register, 1) & $holidayName)
	_Update($SelDate)

EndFunc   ;==>_CalendarTag

Func _CreateMenu()

	GUICtrlDelete($DBpMenu_Report)
	GUICtrlDelete($DBpMenu_Delete)
	GUICtrlDelete($BkpMenu_reset_1)
	GUICtrlDelete($BkpMenu_Exit)
	Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
	Global $DBpMenu_Report = GUICtrlCreateMenu("Report", $DBpMenu_db)

	Local $sSubKey = ""
	For $i = 1 To 12

		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

		$DBpMenu_Delete_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Delete)
		$DBpMenu_Report_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Report)

	Next

	Global $BkpMenu_reset_1 = GUICtrlCreateMenuItem("", $DBpMenu_db)
	Global $BkpMenu_Exit = GUICtrlCreateMenuItem("&Exit", $DBpMenu_db)

EndFunc   ;==>_CreateMenu

Func _CheckDate($DateToCheck, $NewStatus)

	$DateToCheck_split = StringSplit($DateToCheck, "/")

	$DateToCheck_Value = RegRead($DB & "\" & $DateToCheck_split[1] & "\" & $DateToCheck_split[2], $DateToCheck_split[3])
;~ 	ConsoleWrite("## - %" & $DateToCheck_Value & "% - %" & $NewStatus & "%" & @CRLF)

	If $NewStatus = "" Then
		$WeekDayNum = _DateToDayOfWeek($DateToCheck_split[1], $DateToCheck_split[2], $DateToCheck_split[3])
		If $WeekDayNum = "1" Or $WeekDayNum = "7" Then
			$NewStatus = "W"
		EndIf
	EndIf

	$DateToCheck_Value = StringLeft($DateToCheck_Value, 1)
;~ MsgBox(262144,"aqui","%" & $DateToCheck_Value & "%")

	If $DateToCheck_Value <> "" And $DateToCheck_Value <> "B" And StringLeft($DateToCheck_Value, 1) <> $NewStatus Then
		If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(262436, "Replace current value", "You're about to replace the current status for the selected date. " & @CRLF & @CRLF & "Current Status: " & _Label(StringLeft($DateToCheck_Value, 1)) & @CRLF & "New Status: " & _Label($NewStatus) & @CRLF & @CRLF & "Do you want to continue?")
		Select
			Case $iMsgBoxAnswer = 6         ;Yes
				Return 0

			Case $iMsgBoxAnswer = 7         ;No
				Return 1

		EndSelect

	Else
		Return 0
	EndIf

EndFunc   ;==>_CheckDate

Func _Label($LabelName)

	If $LabelName = "" Then Return "Blank"
	If $LabelName = "O" Then Return "On Site"
	If $LabelName = "R" Then Return "Remote"
	If $LabelName = "H" Then Return "Holiday"
	If $LabelName = "P" Then Return "PTO"
	If $LabelName = "T" Then Return "Travel"
	If $LabelName = "S" Then Return "Sick"
	If $LabelName = "B" Then Return "Blank"
	If $LabelName = "W" Then Return "Weekend"

EndFunc   ;==>_Label

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
						$HolidaysLine_Setting = StringSplit($HolidaysLine, "=")
						$RegError = RegWrite($DB, $HolidaysLine_Setting[1], "REG_SZ", $HolidaysLine_Setting[2])
						If @error Then
							$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
						Else
							$ImportCount += 1
						EndIf
					EndIf
				Else
					$HolidaysLine_key = StringSplit($HolidaysLine, "\")
					$HolidaysLine_Value = StringSplit($HolidaysLine_key[3], "=")
					$RegError = RegWrite($DB & "\" & $HolidaysLine_key[1] & "\" & $HolidaysLine_key[2], $HolidaysLine_Value[1], "REG_SZ", $HolidaysLine_Value[2])
					If @error Then
						$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
					Else
						$ImportCount += 1
					EndIf
				EndIf

			WEnd

			If $HolidaysError <> "" Then
				MsgBox(262160, "Import", "Oops! Something went wrong when read the file." & @CRLF & "The following lines was not imported:" & @CRLF & @CRLF & $HolidaysError & @CRLF & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess)
			Else
				If $ImportCount > 15 Then
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & @CRLF & $ImportCount & " lines imported.")
				Else
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess)
				EndIf
			EndIf
		EndIf

	EndIf

EndFunc   ;==>_RestoreBackup

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
		GUICtrlSetData($Input_Tag, StringTrimLeft($Data_Register1, 1))
	Else
		$tip = ""
		GUICtrlSetData($Input_Tag, $tip)
	EndIf
	$WeekDayNum = _DateToDayOfWeek($Data_year, $Data_month, $Data_day)
	$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
	If $Data_Register = "B" Then
		If $tip <> "" Then
			$Data_Register = "   "
		Else
			$Data_Register = ""
		EndIf
	EndIf

	GUICtrlSetData($Inputs[$Data_day][$Data_month], $Data_Register)
	GUICtrlSetTip($Inputs[$Data_day][$Data_month], $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & $tip)
	If $tip <> "" Then
		GUICtrlSetFont($Inputs[$Data_day][$Data_month], 9, 900, 6, "", 2)
	Else
		GUICtrlSetFont($Inputs[$Data_day][$Data_month], 9, 100, 0, "", 2)
	EndIf

	If $Data_Register = "T" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Travel) ; Travel
		$Font_Travel = $Black
		If $Picker_Font_Travel_Read = 1 Then
			$Font_Travel = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_Travel)
	EndIf

	If $Data_Register = "W" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Weekend) ; Weekend
		$Font_Weekend = $Black
		If $Picker_Font_Weekend_Read = 1 Then
			$Font_Weekend = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_Weekend)
	EndIf

	If $Data_Register = "O" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_OnSite) ; On-site
		$Font_OnSite = $Black
		If $Picker_Font_OnSite_Read = 1 Then
			$Font_OnSite = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_OnSite)
	EndIf

	If $Data_Register = "R" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Remote) ; Remote
		$Font_Remote = $Black
		If $Picker_Font_Remote_Read = 1 Then
			$Font_Remote = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_Remote)
	EndIf

	If $Data_Register = "P" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_PTO) ; PTO
		$Font_PTO = $Black
		If $Picker_Font_PTO_Read = 1 Then
			$Font_PTO = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_PTO)
	EndIf

	If $Data_Register = "H" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_holiday) ; holiday
		$Font_Holiday = $Black
		If $Picker_Font_Holiday_Read = 1 Then
			$Font_Holiday = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_Holiday)
	EndIf

	If $Data_Register = "S" Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Sick) ; Sick
		$Font_Sick = $Black
		If $Picker_Font_Sick_Read = 1 Then
			$Font_Sick = $White
		EndIf
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_Sick)
	EndIf

	If $Data_Register = "" Or $Data_Register = "B" Or $Data_Register = "   " Then
		GUICtrlSetBkColor($Inputs[$Data_day][$Data_month], $Color_bk_Blank) ; Blank
		GUICtrlSetColor($Inputs[$Data_day][$Data_month], $Font_Blank)
	EndIf

	If $Data_year & "/" & $Data_month & "/" & $Data_day = @YEAR & "/" & @MON & "/" & @MDAY Then

		If $tip <> "" Then
			GUICtrlSetFont($Inputs[$Data_day][$Data_month], 9, 900, 6, "", 2)
		Else
			GUICtrlSetFont($Inputs[$Data_day][$Data_month], 9, 100, 0, "", 2)
		EndIf

		GUICtrlSetTip($Inputs[$Data_day][$Data_month], $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & " - TODAY" & $tip)
		If $Data_Register = "" Then
			$Data_Register = "X"
			GUICtrlSetColor($Inputs[$Data_day][$Data_month], 0xFF0000)             ; today
		EndIf

	EndIf

	_ReadStatistics($Data_year)
	_CreateMenu()


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
		If $FileHolidays_hwd <> -1 Then

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
		_CreateMenu()
	EndIf

EndFunc   ;==>_ImportHolidays

Func _ResetDatabase($step = "0")

	_CreateMenu()

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
			Return 1
		EndIf

	EndIf


	Return


EndFunc   ;==>_ResetDatabase

Func _CalendarRead($i = 0, $j = 0)

	For $a = 1 To 12
		For $b = 1 To 31
;~ 			ConsoleWrite($b & "-" & $a & "-" & GUICtrlGetState($SelectLabel[$b][$a]) & @CRLF)
			If GUICtrlGetState($SelectLabel[$b][$a]) = 144 Then
;~ 				ConsoleWrite($b & "-" & $a & "-" & GUICtrlGetState($SelectLabel[$b][$a]) & @CRLF)
				GUICtrlSetState($SelectLabel[$b][$a], $gui_hide)
			EndIf
		Next
	Next



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

;~ 	ConsoleWrite($SelDate_slipt[2] & "-" & $SelDate_slipt[3] & "-" & GUICtrlGetState($SelectLabel[$SelDate_slipt[2]][$SelDate_slipt[3]]) & @CRLF)
	GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

	$Status_Tip = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	GUICtrlSetData($Input_Tag, StringTrimLeft($Status_Tip, 1))

	Return

EndFunc   ;==>_CalendarRead

Func _ClearScreen()

	For $j = 1 To 12
		For $i = 1 To 31
			GUICtrlDelete($Inputs[$i][$j])
			GUICtrlDelete($TodayLabel[$i][$j])
			GUICtrlDelete($SelectLabel[$i][$j])
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

	GUICtrlSetBkColor($Input_Remaining_q1, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q2, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q3, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q4, 0xFFFFFF)

	; Criar ListView com colunas para os dias do mes
	$Headers = ""
	For $i = 1 To 31
		$Headers &= "|" & $i
	Next

	; Criar Inputs para cabecalhos (dias do mes)
	For $i = 1 To 31
		If $i < 10 Then
			$n = "0" & $i
		Else
			$n = $i
		EndIf
	Next
	$C = 0
	$Skip = 0
	For $j = 1 To 12
		If $j < 10 Then
			$X = "0" & $j
		Else
			$X = $j
		EndIf

		For $i = 1 To 31
			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf
			$IniSection[$j][$i] = RegEnumVal($DB & "\" & $Year & "\" & $X, $n)
			If @error Then ExitLoop
		Next

		$Return = _DateToMonth($X, 1)

		If @error Then ContinueLoop

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

				If $j = "01" Or $j = "02" Or $j = "03" Then
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

				If $j = "04" Or $j = "05" Or $j = "06" Then
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

				If $j = "07" Or $j = "08" Or $j = "09" Then
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

				If $j = "10" Or $j = "11" Or $j = "12" Then
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
					If $j = "01" Or $j = "02" Or $j = "03" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then

							$Counta_WD_q1 += 1
						EndIf
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

					If $j = "04" Or $j = "05" Or $j = "06" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q2 += 1
						EndIf
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

					If $j = "07" Or $j = "08" Or $j = "09" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q3 += 1
						EndIf
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

					If $j = "10" Or $j = "11" Or $j = "12" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q4 += 1
						EndIf
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
					If $j = "01" Or $j = "02" Or $j = "03" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q1 += 1
						EndIf

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
					If $j = "04" Or $j = "05" Or $j = "06" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q2 += 1
						EndIf
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
					If $j = "07" Or $j = "08" Or $j = "09" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q3 += 1
						EndIf
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
					If $j = "10" Or $j = "11" Or $j = "12" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q4 += 1
						EndIf
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
					If $j = "01" Or $j = "02" Or $j = "03" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q1 += 1
						EndIf
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
					If $j = "04" Or $j = "05" Or $j = "06" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q2 += 1
						EndIf
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

					If $j = "07" Or $j = "08" Or $j = "09" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q3 += 1
						EndIf
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

					If $j = "10" Or $j = "11" Or $j = "12" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q4 += 1
						EndIf
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
				If $Status = "" Or $Status = "B" Then
					If $j = "01" Or $j = "02" Or $j = "03" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q1 += 1
						EndIf

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

					If $j = "04" Or $j = "05" Or $j = "06" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q2 += 1
						EndIf

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

					If $j = "07" Or $j = "08" Or $j = "09" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q3 += 1
						EndIf

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

					If $j = "10" Or $j = "11" Or $j = "12" Then
						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
							$Counta_WD_q4 += 1
						EndIf

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

	GUICtrlSetData($Input_E_Onsite_q1, Ceiling(($Counta_WD_q1 / 5) * 3)) ;## Estm.On-Site ##
	GUICtrlSetData($Input_E_Onsite_q2, Ceiling(($Counta_WD_q2 / 5) * 3))
	GUICtrlSetData($Input_E_Onsite_q3, Ceiling(($Counta_WD_q3 / 5) * 3))
	GUICtrlSetData($Input_E_Onsite_q4, Ceiling(($Counta_WD_q4 / 5) * 3))

	GUICtrlSetData($Input_R_Onsite_q1, $Counta_R_Onsite_q1) ;## Real On-Site ##
	GUICtrlSetData($Input_R_Onsite_q2, $Counta_R_Onsite_q2)
	GUICtrlSetData($Input_R_Onsite_q3, $Counta_R_Onsite_q3)
	GUICtrlSetData($Input_R_Onsite_q4, $Counta_R_Onsite_q4)

	$Remaining_q1 = (Ceiling(($Counta_WD_q1 / 5) * 3)) - $Counta_R_Onsite_q1
	$Remaining_q2 = (Ceiling(($Counta_WD_q2 / 5) * 3)) - $Counta_R_Onsite_q2
	$Remaining_q3 = (Ceiling(($Counta_WD_q3 / 5) * 3)) - $Counta_R_Onsite_q3
	$Remaining_q4 = (Ceiling(($Counta_WD_q4 / 5) * 3)) - $Counta_R_Onsite_q4

	GUICtrlSetData($Input_Remaining_q1, $Remaining_q1) ;## Remaining ##
	GUICtrlSetData($Input_Remaining_q2, $Remaining_q2)
	GUICtrlSetData($Input_Remaining_q3, $Remaining_q3)
	GUICtrlSetData($Input_Remaining_q4, $Remaining_q4)

	$Ratio_R_Q1 = Round(($Counta_R_Onsite_q1 / Ceiling($Counta_WD_q1 / 5)), 2)
	$Ratio_R_Q2 = Round(($Counta_R_Onsite_q2 / Ceiling($Counta_WD_q2 / 5)), 2)
	$Ratio_R_Q3 = Round(($Counta_R_Onsite_q3 / Ceiling($Counta_WD_q3 / 5)), 2)
	$Ratio_R_Q4 = Round(($Counta_R_Onsite_q4 / Ceiling($Counta_WD_q4 / 5)), 2)

	GUICtrlSetData($Input_RT_q1, $Ratio_R_Q1) ; ## Ration ##
	GUICtrlSetBkColor($Input_RT_q1, _GetColorGradient($Ratio_R_Q1))


	GUICtrlSetData($Input_RT_q2, $Ratio_R_Q2)
	GUICtrlSetBkColor($Input_RT_q2, _GetColorGradient($Ratio_R_Q2))


	GUICtrlSetData($Input_RT_q3, $Ratio_R_Q3)
	GUICtrlSetBkColor($Input_RT_q3, _GetColorGradient($Ratio_R_Q3))


	GUICtrlSetData($Input_RT_q4, $Ratio_R_Q4)
	GUICtrlSetBkColor($Input_RT_q4, _GetColorGradient($Ratio_R_Q4))


	$Ratio_Q1 = Round(($Counta_R_Onsite_Quarter_Q1 / Ceiling($Counta_WD_Quarter_Q1 / 5)), 2)
	$Ratio_Q2 = Round(($Counta_R_Onsite_Quarter_Q2 / Ceiling($Counta_WD_Quarter_Q2 / 5)), 2)
	$Ratio_Q3 = Round(($Counta_R_Onsite_Quarter_Q3 / Ceiling($Counta_WD_Quarter_Q3 / 5)), 2)
	$Ratio_Q4 = Round(($Counta_R_Onsite_Quarter_Q4 / Ceiling($Counta_WD_Quarter_Q4 / 5)), 2)

	GUICtrlSetData($Input_RaTio_q1, "")
	GUICtrlSetData($Input_RaTio_q2, "")
	GUICtrlSetData($Input_RaTio_q3, "")
	GUICtrlSetData($Input_RaTio_q4, "")

;~ 	ConsoleWrite("$Counta_R_Onsite_Quarter_Q1: " & $Counta_R_Onsite_Quarter_Q1 & @CRLF)
;~ 	ConsoleWrite("$Counta_WD_Quarter_Q1: " & $Counta_WD_Quarter_Q1 & @CRLF)
;~ 	ConsoleWrite("Ratio: " & ($Counta_R_Onsite_Quarter_Q1 / Ceiling($Counta_WD_Quarter_Q1 / 5)) & @CRLF)

	If $Year = @YEAR Then
		If @MON = "01" Or @MON = "02" Or @MON = "03" Then
			GUICtrlSetData($Input_RaTio_q1, $Ratio_Q1)
			GUICtrlSetBkColor($Input_RaTio_q1, _GetColorGradient($Ratio_Q1))
		EndIf

		If @MON = "04" Or @MON = "05" Or @MON = "06" Then
			GUICtrlSetData($Input_RaTio_q2, $Ratio_Q2)
			GUICtrlSetBkColor($Input_RaTio_q2, _GetColorGradient($Ratio_Q2))
		EndIf

		If @MON = "07" Or @MON = "08" Or @MON = "09" Then
			GUICtrlSetData($Input_RaTio_q3, $Ratio_Q3)
			GUICtrlSetBkColor($Input_RaTio_q3, _GetColorGradient($Ratio_Q3))
		EndIf

		If @MON = "10" Or @MON = "11" Or @MON = "12" Then
			GUICtrlSetData($Input_RaTio_q4, $Ratio_Q4)
			GUICtrlSetBkColor($Input_RaTio_q4, _GetColorGradient($Ratio_Q4))
		EndIf
	EndIf

	_CheckQuarter()

	Return

EndFunc   ;==>_ReadStatistics

Func _ReadINI($Year)

	GUICtrlSetData($Input_Tag, "")

	_ClearScreen()

	_ReadStatistics($Year)

	; Criar ListView com colunas para os dias do mÃªs
	$Headers = ""
	For $i = 1 To 31
		$Headers &= "|" & $i
	Next

	; Criar Inputs para cabeÃ§alhos (dias do mÃªs)
	For $i = 1 To 31
		If $i < 10 Then
			$n = "0" & $i
		Else
			$n = $i
		EndIf
		$LabelMonth[$i] = GUICtrlCreateLabel($n, 5 + ($i * 35), 216, 20, 20, $SS_CENTER)
	Next
	$C = 0
	$Skip = 0
	For $j = 1 To 12
		If $j < 10 Then
			$X = "0" & $j
		Else
			$X = $j
		EndIf

		For $i = 1 To 31
;~
			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf
			$IniSection[$j][$i] = RegEnumVal($DB & "\" & $Year & "\" & $X, $n)
			If @error Then ExitLoop
		Next


		$Return = _DateToMonth($X, 1)

		If @error Then ContinueLoop ; Se a seÃ§Ã£o nÃ£o existir, pula para o prÃ³ximo mÃªs

		; Month
		$LabelMonthX[$j] = GUICtrlCreateLabel($Return, 8, 208 + $Skip + ($j * 25), 20, 20, $SS_CENTER) ;,$SS_BLACKRECT)

		;Days
		For $i = 1 To 31

			If $i < 10 Then
				$n = "0" & $i
			Else
				$n = $i
			EndIf

			If _DateIsValid($Year & "/" & $X & "/" & $i) = 1 Then

				$SelectLabel[$i][$j] = GUICtrlCreateLabel("", -3 + ($i * 35), 202 + $Skip + ($j * 25), 36, 28) ;,$SS_BLACKFRAME)
				GUICtrlSetBkColor($SelectLabel[$i][$j], $Color_bk_Selected)
				GUICtrlSetState($SelectLabel[$i][$j], $gui_disable)
				GUICtrlSetState($SelectLabel[$i][$j], $gui_hide)


				$TodayLabel[$i][$j] = GUICtrlCreateLabel("", -1 + ($i * 35), 204 + $Skip + ($j * 25), 32, 24) ;,$SS_BLACKFRAME)
				GUICtrlSetBkColor($TodayLabel[$i][$j], $Color_bk_Today)
				GUICtrlSetColor($TodayLabel[$i][$j], $Color_bk_Today)
				GUICtrlSetState($TodayLabel[$i][$j], $gui_disable)
				GUICtrlSetState($TodayLabel[$i][$j], $gui_hide)


				$Inputs[$i][$j] = GUICtrlCreateButton("", 0 + ($i * 35), 205 + $Skip + ($j * 25), 30, 22, BitOR($ES_READONLY, $ES_CENTER, $BS_FLAT, $BS_BOTTOM))

				$WeekDayNum = _DateToDayOfWeek($Year, $X, $i)
				$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
				$Status1 = RegRead($DB & "\" & $Year & "\" & $X, $n)
				$Status = StringLeft($Status1, 1)
				If StringLen($Status1) > 1 Then
					$tip = " - " & StringTrimLeft($Status1, 1)
					GUICtrlSetFont($Inputs[$i][$j], 9, 900, 6, "", 2)
				Else
					$tip = ""
					GUICtrlSetFont($Inputs[$i][$j], 9, 100, 0, "", 2)
				EndIf

				GUICtrlSetTip($Inputs[$i][$j], $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & $tip)

				If $Status = "B" Then
					If $tip <> "" Then
						$Status = "   "
					Else
						$Status = ""
					EndIf
				EndIf


				GUICtrlSetData($Inputs[$i][$j], $Status)
				If $Status = "W" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Weekend) ; Weekend

					$Font_Weekend = $Black
					If $Picker_Font_Weekend_Read = 1 Then
						$Font_Weekend = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_Weekend)

				EndIf

				If $Status = "O" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_OnSite) ; On-site

					$Font_OnSite = $Black
					If $Picker_Font_OnSite_Read = 1 Then
						$Font_OnSite = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_OnSite)

				EndIf

				If $Status = "R" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Remote) ; Remote

					$Font_Remote = $Black
					If $Picker_Font_Remote_Read = 1 Then
						$Font_Remote = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_Remote)


				EndIf

				If $Status = "T" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Travel) ; Travel

					$Font_Travel = $Black
					If $Picker_Font_Travel_Read = 1 Then
						$Font_Travel = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_Travel)

				EndIf

				If $Status = "P" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_PTO) ; PTO

					$Font_PTO = $Black
					If $Picker_Font_PTO_Read = 1 Then
						$Font_PTO = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_PTO)

				EndIf

				If $Status = "H" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_holiday) ; holiday

					$Font_Holiday = $Black
					If $Picker_Font_Holiday_Read = 1 Then
						$Font_Holiday = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_Holiday)

				EndIf

				If $Status = "" Or $Status = "   " Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Blank) ; Weekend

					$Font_Blank = $Black
					If $Picker_Font_Blank_Read = 1 Then
						$Font_Blank = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_Blank)

				EndIf

				If $Status = "S" Then
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Sick) ; Sick

					$Font_Sick = $Black
					If $Picker_Font_Sick_Read = 1 Then
						$Font_Sick = $White
					EndIf
					GUICtrlSetColor($Inputs[$i][$j], $Font_Sick)

				EndIf

				If $Year & "/" & $X & "/" & $n = @YEAR & "/" & @MON & "/" & @MDAY Then

					GUICtrlSetTip($Inputs[$i][$j], $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & " - TODAY" & $tip)
					GUICtrlSetState($TodayLabel[$i][$j], $gui_show)


				EndIf

			EndIf
		Next

		$C += 1
		If $C > 2 Then
			$C = 0
			$Skip = $Skip + 10
		EndIf

	Next
	_CreateMenu()
	Return

EndFunc   ;==>_ReadINI

Func _CheckQuarter()

	$SelDate = GUICtrlRead($Calendar)
	$Color_bk_Black = 0x000000

	GUICtrlSetBkColor($Input_Remaining_q1, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q2, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q3, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q4, 0xFFFFFF)

	If $Ratio_R_Q1 > 0 Or $Ratio_R_Q1 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q1, _GetColorFromValue($Remaining_q1))
	EndIf
	If $Ratio_R_Q2 > 0 Or $Ratio_R_Q2 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q2, _GetColorFromValue($Remaining_q2))
	EndIf
	If $Ratio_R_Q3 > 0 Or $Ratio_R_Q3 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q3, _GetColorFromValue($Remaining_q3))
	EndIf
	If $Ratio_R_Q4 > 0 Or $Ratio_R_Q4 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q4, _GetColorFromValue($Remaining_q4))
	EndIf

;~ 	ConsoleWrite("$Ratio_Q1: " & $Ratio_Q1 & @CRLF)
;~ 	ConsoleWrite("$Ratio_Q2: " & $Ratio_Q2 & @CRLF)
;~ 	ConsoleWrite("$Ratio_Q3: " & $Ratio_Q3 & @CRLF)
;~ 	ConsoleWrite("$Ratio_Q4: " & $Ratio_Q4 & @CRLF)

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

;~ 			GUICtrlSetBkColor($Input_Remaining_q1, _GetColorFromValue($Remaining_q1))
;~ 			GUICtrlSetBkColor($Input_Remaining_q2, _GetColorFromValue($Remaining_q2))
;~ 			GUICtrlSetBkColor($Input_Remaining_q3, _GetColorFromValue($Remaining_q3))
;~ 			GUICtrlSetBkColor($Input_Remaining_q4, _GetColorFromValue($Remaining_q4))


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

;~ 			GUICtrlSetBkColor($Input_Remaining_q1, _GetColorFromValue($Remaining_q1))
;~ 			GUICtrlSetBkColor($Input_Remaining_q2, _GetColorFromValue($Remaining_q2))
;~ 			GUICtrlSetBkColor($Input_Remaining_q3, _GetColorFromValue($Remaining_q3))
;~ 			GUICtrlSetBkColor($Input_Remaining_q4, _GetColorFromValue($Remaining_q4))

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

;~ 			GUICtrlSetBkColor($Input_Remaining_q1, _GetColorFromValue($Remaining_q1))
;~ 			GUICtrlSetBkColor($Input_Remaining_q2, _GetColorFromValue($Remaining_q2))
;~ 			GUICtrlSetBkColor($Input_Remaining_q3, _GetColorFromValue($Remaining_q3))
;~ 			GUICtrlSetBkColor($Input_Remaining_q4, _GetColorFromValue($Remaining_q4))

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

;~ 			GUICtrlSetBkColor($Input_Remaining_q1, _GetColorFromValue($Remaining_q1))
;~ 			GUICtrlSetBkColor($Input_Remaining_q2, _GetColorFromValue($Remaining_q2))
;~ 			GUICtrlSetBkColor($Input_Remaining_q3, _GetColorFromValue($Remaining_q3))
;~ 			GUICtrlSetBkColor($Input_Remaining_q4, _GetColorFromValue($Remaining_q4))

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
	_CreateMenu()
	Return

EndFunc   ;==>_CriaINI

Func _BKColorPallet()

	; Create custom (4 x 5) color palette
	Dim $aPalette[20] = _
			[0xFFFFFF, 0x000000, 0xC0C0C0, 0x808080, _
			0xFF0000, 0x800000, 0xFFFF00, 0x808000, _
			0x00FF00, 0x008000, 0x00FFFF, 0x008080, _
			0x0000FF, 0x000080, 0xFF00FF, 0x800080, _
			0xC0DCC0, 0xA6CAF0, 0xFFFBF0, 0xA0A0A4]

	$WinPos = WinGetPos("Work Days")
	$Form_Colors = GUICreate('Colors', 220, 400, $WinPos[0] + 300, $WinPos[1] + 100, $DS_MODALFRAME, $WS_EX_TOPMOST)
	GUICtrlSetBkColor(-1, 0x50CA1B)

	GUICtrlCreateLabel("On Site:", 10, 15)
	GUICtrlCreateLabel("Remote:", 10, 45)
	GUICtrlCreateLabel("Holiday:", 10, 75)
	GUICtrlCreateLabel("PTO:", 10, 105)
	GUICtrlCreateLabel("Travel:", 10, 135)
	GUICtrlCreateLabel("Sick:", 10, 165)
	GUICtrlCreateLabel("Blank:", 10, 195)
	GUICtrlCreateLabel("Weekend:", 10, 225)
	GUICtrlCreateLabel("Today:", 10, 255)
	GUICtrlCreateLabel("Selected:", 10, 285)

	$Picker_OnSite = _GUIColorPicker_Create('', 65, 10, 60, 23, $Color_bk_OnSite, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Remote = _GUIColorPicker_Create('', 65, 40, 60, 23, $Color_bk_Remote, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Holiday = _GUIColorPicker_Create('', 65, 70, 60, 23, $Color_bk_holiday, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_PTO = _GUIColorPicker_Create('', 65, 100, 60, 23, $Color_bk_PTO, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Travel = _GUIColorPicker_Create('', 65, 130, 60, 23, $Color_bk_Travel, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Sick = _GUIColorPicker_Create('', 65, 160, 60, 23, $Color_bk_Sick, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Blank = _GUIColorPicker_Create('', 65, 190, 60, 23, $Color_bk_Blank, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Weekend = _GUIColorPicker_Create('', 65, 220, 60, 23, $Color_bk_Weekend, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Today = _GUIColorPicker_Create('', 65, 250, 60, 23, $Color_bk_Today, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Selected = _GUIColorPicker_Create('', 65, 280, 60, 23, $Color_bk_Selected, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')

	$Picker_Font_OnSite = GUICtrlCreateCheckbox("White Font", 130, 10)
	$Picker_Font_Remote = GUICtrlCreateCheckbox("White Font", 130, 40)
	$Picker_Font_Holiday = GUICtrlCreateCheckbox("White Font", 130, 70)
	$Picker_Font_PTO = GUICtrlCreateCheckbox("White Font", 130, 100)
	$Picker_Font_Travel = GUICtrlCreateCheckbox("White Font", 130, 130)
	$Picker_Font_Sick = GUICtrlCreateCheckbox("White Font", 130, 160)
	$Picker_Font_Blank = GUICtrlCreateCheckbox("White Font", 130, 190)
	$Picker_Font_Weekend = GUICtrlCreateCheckbox("White Font", 130, 220)

	GUICtrlSetState($Picker_Font_OnSite, $Picker_Font_OnSite_Read)
	GUICtrlSetState($Picker_Font_Remote, $Picker_Font_Remote_Read)
	GUICtrlSetState($Picker_Font_Holiday, $Picker_Font_Holiday_Read)
	GUICtrlSetState($Picker_Font_PTO, $Picker_Font_PTO_Read)
	GUICtrlSetState($Picker_Font_Travel, $Picker_Font_Travel_Read)
	GUICtrlSetState($Picker_Font_Sick, $Picker_Font_Sick_Read)
	GUICtrlSetState($Picker_Font_Blank, $Picker_Font_Blank_Read)
	GUICtrlSetState($Picker_Font_Weekend, $Picker_Font_Weekend_Read)

	$Original_Color_1 = $Color_bk_OnSite & $Color_bk_Remote & $Color_bk_holiday & $Color_bk_PTO & $Color_bk_Travel & $Color_bk_Sick & $Color_bk_Blank & $Color_bk_Weekend & $Color_bk_Today & $Color_bk_Selected & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & $Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read
;~ 	ConsoleWrite($Original_Color_1 & @CRLF)

	$Colors_Close = GUICtrlCreateButton("Close", 80, 330, 70, 30)

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
				$Picker_Color_Today = _GUIColorPicker_GetColor($Picker_Today)
				$Picker_Color_Selected = _GUIColorPicker_GetColor($Picker_Selected)

				RegWrite($DB, "Color_OnSite", "REG_SZ", $Picker_Color_OnSite)
				RegWrite($DB, "Color_Remote", "REG_SZ", $Picker_Color_Remote)
				RegWrite($DB, "Color_holiday", "REG_SZ", $Picker_Color_Holiday)
				RegWrite($DB, "Color_PTO", "REG_SZ", $Picker_Color_PTO)
				RegWrite($DB, "Color_Travel", "REG_SZ", $Picker_Color_Travel)
				RegWrite($DB, "Color_Sick", "REG_SZ", $Picker_Color_Sick)
				RegWrite($DB, "Color_Blank", "REG_SZ", $Picker_Color_Blank)
				RegWrite($DB, "Color_Weekend", "REG_SZ", $Picker_Color_Weekend)
				RegWrite($DB, "Color_Today", "REG_SZ", $Picker_Color_Today)
				RegWrite($DB, "Color_Selected", "REG_SZ", $Picker_Color_Selected)

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

				$Original_Color_2 = $Picker_Color_OnSite & $Picker_Color_Remote & $Picker_Color_Holiday & $Picker_Color_PTO & $Picker_Color_Travel & $Picker_Color_Sick & $Picker_Color_Blank & $Picker_Color_Weekend & $Picker_Color_Today & $Picker_Color_Selected & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & $Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read

;~ 				ConsoleWrite($Original_Color_2 & @CRLF)

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

	Global $Color_bk_Today = RegRead($DB, "Color_Today")
	If @error Then $Color_bk_Today = 0xFF000000


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

	GUICtrlSetBkColor($SelectLabel_1, $Color_bk_Today)
	GUICtrlSetBkColor($TodayLabel_1, $Color_bk_Selected)


EndFunc   ;==>_ReadColors

Func _CreateBackup($DBBKP = "")

	Local $sRegPath = $DB & "\"

	If $DBBKP = "" Then
		Local $sFilePath = FileSaveDialog("Save backup file", @ScriptDir, "All (*.*)", 18, "Backup_" & @YEAR & "_" & @MON & "_" & @MDAY & ".bkp", $Form_WorkDays)
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
		FileWriteLine($sFilePath_hwd, $sSubKey_settings & "=" & $RegRead)
;~ 		ConsoleWrite($sSubKey_settings & "=" & $RegRead & @CRLF)
	Next

	; Loop from 1 to 10 times, displaying registry keys at the particular instance value.
	For $i = 1 To 10000
		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

;~ 		ConsoleWrite($DB & "\" & $sSubKey & @CRLF)

		For $r = 1 To 10000
			$sSubKey_month = RegEnumKey($DB & "\" & $sSubKey, $r)
			If @error Then ExitLoop

;~ 			ConsoleWrite($DB & "\" & $sSubKey & "\" & $sSubKey_month & @CRLF)

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

Func _Interpolate($v1, $v2, $ratio)
	Return Round($v1 + ($v2 - $v1) * $ratio)
EndFunc   ;==>_Interpolate

Func _GetColorGradient($value)
	; Limita o valor mínimo
	If $value < 0.1 Then
		If $value = 0 Then
			$value = 0
		Else
			$value = 0.1
		EndIf
	EndIf

	; Verde escuro fixo para valores acima de 3.0
	If $value > 3.0 Then
		Return "0x" & StringFormat("%02X%02X%02X", 0, 200, 0)
	EndIf

	If $value = 0 Then
		Return "0x" & StringFormat("%02X%02X%02X", 255, 255, 255)
	EndIf

	; Define os pontos de controle (valor, RGB)
	Local $points[5][4] = [ _
			[0.1, 255, 0, 0], _       ; Vermelho
			[1.0, 255, 128, 0], _     ; Laranja-avermelhado
			[2.0, 200, 165, 0], _     ; Laranja
			[2.5, 173, 255, 47], _    ; Amarelo-esverdeado
			[3.0, 0, 255, 0] _        ; Verde claro
			]

	; Procura os dois pontos entre os quais o valor se encontra
	Local $i
	For $i = 0 To UBound($points) - 2
		If $value >= $points[$i][0] And $value <= $points[$i + 1][0] Then
			ExitLoop
		EndIf
	Next

	Local $v1 = $points[$i][0]
	Local $r1 = $points[$i][1]
	Local $g1 = $points[$i][2]
	Local $b1 = $points[$i][3]

	Local $v2 = $points[$i + 1][0]
	Local $r2 = $points[$i + 1][1]
	Local $g2 = $points[$i + 1][2]
	Local $b2 = $points[$i + 1][3]

	; Calcula a razão de interpolação entre os dois pontos
	Local $ratio = ($value - $v1) / ($v2 - $v1)

	; Interpola cada canal de cor
	Local $r = _Interpolate($r1, $r2, $ratio)
	Local $g = _Interpolate($g1, $g2, $ratio)
	Local $b = _Interpolate($b1, $b2, $ratio)

	; Retorna em formato hexadecimal
	Return "0x" & StringFormat("%02X%02X%02X", $r, $g, $b)
EndFunc   ;==>_GetColorGradient


Func _GetColorFromValue($iValue)
;~
	; Limita o intervalo
	If $iValue < -30 Then $iValue = -30
	If $iValue > 31 Then $iValue = 31

	; Valores de 0 ou menores = verde
	If $iValue <= 0 Then
		Return 0x00FF00 ; RGB(0,255,0)
	EndIf

	; Valores de 1 a 31: gradiente de amarelo a vermelho
	; Verde varia de 255 (no valor 1) até 0 (no valor 31)
	Local $fRatio = ($iValue - 1) / 30
	Local $iRed = 255
	Local $iGreen = Int(255 * (1 - $fRatio))
	Local $iBlue = 0

	; Retorna no formato 0xRRGGBB
	Return BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue
EndFunc   ;==>_GetColorFromValue

; Função: GetColorFromValue
; Descrição: Retorna uma cor RGB em função de um valor entre -30 e 30.
; Valores >= 0 variam do verde (0) ao vermelho (30).
; Valores < 0 retornam um verde fixo.

Func _oldGetColorFromValue($iValue)
	; Garante que o valor está no intervalo permitido
	If $iValue < -30 Then $iValue = -30
	If $iValue > 30 Then $iValue = 30

	; Para valores negativos, retorna verde fixo
	If $iValue < 0 Then
		Return 0x00FF00 ; Verde puro (R=0, G=255, B=0)
	EndIf

	; Para valores entre 0 e 30, calcular do verde ao vermelho
	; Quanto maior o valor, mais vermelho e menos verde
	Local $fRatio = $iValue / 30
	Local $iRed = Int(255 * $fRatio)
	Local $iGreen = Int(255 * (1 - $fRatio))
	Local $iBlue = 0

	; Converte para RGB no formato 0xRRGGBB
	Return BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue
EndFunc   ;==>_oldGetColorFromValue


