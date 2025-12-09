#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=xcalendar4.ico
#AutoIt3Wrapper_Res_Description=Work Day management
#AutoIt3Wrapper_Res_Fileversion=1.0.5.2
#AutoIt3Wrapper_Res_ProductName=Work Days
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\splash.jpg
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\Help.pdf
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\about.jpg
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#pragma compile(inputboxres, true)
Opt("TrayIconHide", 1)
Opt("TrayAutoPause", 0)

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
#include <FontConstants.au3>
#include "GenerateWorkdaysReportHTML.au3"
#include "wkhtmltox.au3"
#include <ProgressConstants.au3>

Global $HelpFile = @TempDir & "\Help.pdf"
Global $sSplashPath = @TempDir & "\splash.jpg"
Global $AboutFile = @TempDir & "\splash.jpg"

Global $Progress_Splash, $Form_Splash

FileInstall("splash.jpg", $sSplashPath, 1)
_splash("on")

Global $About = "1.0.1.3 - Custom colors and bug fixes" & @CRLF _
		 & "1.0.1.4 - Code polishing and new custom color palette" & @CRLF _
		 & "1.0.1.8 - Today custom color option" & @CRLF _
		 & "1.0.1.9 - Report Functionality" & @CRLF _
		 & "1.0.2.2 - Report and Tag Multiline" & @CRLF _
		 & "1.0.3.0 - New Contextual Menu, Splash Screen and about" & @CRLF _
		 & "1.0.4.0 - Report in PDF" & @CRLF _
		 & "1.0.4.1 - Bug fix: 'Ratio to date' metric now calculates correctly." & @CRLF _
		 & "1.0.4.2 - Bug fix: Import full database when in a different year." & @CRLF _
		 & "1.0.5.0 - Layout update." & @CRLF _
		 & "1.0.5.2 - Minor updates."

Global $XCount = 0



Global $IniSection[999][999]
Global $LabelMonth[99999]
Global $LabelMonthX[99999]
Global $Inputs[32][32]


Global $Context[32][32]
Global $ContextItem_Date[32][32]
Global $ContextItem_Separator[32][32]
Global $ContextItem_Tag[32][32]
Global $ContextItem_OnSite[32][32]
Global $ContextItem_Remote[32][32]
Global $ContextItem_Holiday[32][32]
Global $ContextItem_PTO[32][32]
Global $ContextItem_Travel[32][32]
Global $ContextItem_Sick[32][32]
Global $ContextItem_Blank[32][32]


Global $TodayLabel[32][32]
Global $SelectLabel[32][32]
Global $DBpMenu_Delete_Year[20]
Global $DBpMenu_Delete_Date[15]

Global $DBpMenu_Report_simple_Year[20]
Global $DBpMenu_Report_detailed_Year[20]
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


; Caminho de extração (tempo de execução)




$DB = "HKEY_CURRENT_USER\Software\WorkDays"


Global $CalendarTag = RegRead($DB, "caltag")
If @error Then $CalendarTag = "1"

Global $Debug = RegRead($DB, "Debug")
If @error Then $Debug = "0"

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
Global $Window_X = 1140
Global $Window_Y = 620
$Form_WorkDays = GUICreate("Work Days", $Window_X, $Window_Y, -1, -1)
;~ $Form_WorkDays = GUICreate("Work Days", 1140, 620, -1, -1, $WS_SYSMENU)

Global $DBpMenu_db = GUICtrlCreateMenu("File")
Global $DBpMenu_backup_Data = GUICtrlCreateMenu("Data")
Global $DBpMenu_backup = GUICtrlCreateMenuItem("Create Backup", $DBpMenu_backup_Data)
Global $BkpMenu_Batch = GUICtrlCreateMenuItem("Restore Backup", $DBpMenu_backup_Data)
Global $DBpMenu_backup_2 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
Global $BkpMenu_reset_all1 = GUICtrlCreateMenu("Data Management", $DBpMenu_backup_Data)
Global $BkpMenu_reset_all = GUICtrlCreateMenuItem("Reset Entire Database", $BkpMenu_reset_all1)
Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
Global $DBpMenu_backup_Data_Holidays = GUICtrlCreateMenuItem("Import Holidays File", $DBpMenu_backup_Data)
Global $DBpMenu_Report = GUICtrlCreateMenu("Report")
Global $DBpMenu_Report_Simple = GUICtrlCreateMenu("Simple", $DBpMenu_Report)
Global $DBpMenu_Report_Detailed = GUICtrlCreateMenu("Detailed", $DBpMenu_Report)


Global $BkpMenu_Exit = GUICtrlCreateMenuItem("&Exit", $DBpMenu_db)

Global $DBpMenu_settings = GUICtrlCreateMenu("Settings")
Global $BkpMenu_settings_BKcolors = GUICtrlCreateMenuItem("Options", $DBpMenu_settings)
Global $BkpMenu_help = GUICtrlCreateMenu("?")
Global $BkpMenu_help_help = GUICtrlCreateMenuItem("Help", $BkpMenu_help)
Global $BkpMenu_help_space = GUICtrlCreateMenuItem("", $BkpMenu_help)
Global $BkpMenu_help_About = GUICtrlCreateMenuItem("About", $BkpMenu_help)


$Calendar = GUICtrlCreateMonthCal(@YEAR & "/" & @MON & "/" & @MDAY, 8, 8, 273, 201, $MCS_WEEKNUMBERS)

$Group1 = GUICtrlCreateGroup("", 288, 8, 270, 200)

$Input_SelDate = GUICtrlCreateInput("", 376, 24, 70, 21, $ES_READONLY)
GUICtrlSetData($Input_SelDate, GUICtrlRead($Calendar))
GUICtrlSetColor($Input_SelDate, 0x990000)
GUICtrlSetState($Input_SelDate, $gui_disable)

$Label1 = GUICtrlCreateLabel("Selected Date:", 296, 28, 75, 17)
$Input_Quarter = GUICtrlCreateInput("", 450, 24, 20, 21, $ES_READONLY)
GUICtrlSetColor($Input_Quarter, 0x00994C)
GUICtrlSetState($Input_Quarter, $gui_disable)

$Input_Tag = GUICtrlCreateInput("", 296, 54, 175, 21) ;, $ES_READONLY)
GUICtrlSetState($Input_Tag, $gui_hide)

;~ $Button_CalendtarTag = GUICtrlCreateButton("Tag", 472, 52, 75, 25) ;## Calendar TAG
;~ GUICtrlSetTip($Button_CalendtarTag, "Use /n as linebreak.")
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

$Label_4_q1 = GUICtrlCreateLabel("Estim.On-Site:", 700, 30, 65, 21, $SS_RIGHT)
$Label_5_q1 = GUICtrlCreateLabel("Real On-Site:", 700, 50, 65, 21, $SS_RIGHT)
$Label_6_q1 = GUICtrlCreateLabel("Remaining:", 700, 70, 65, 21, $SS_RIGHT)

$Input_TD_q1 = GUICtrlCreateLabel("", 651, 30, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q1 = GUICtrlCreateLabel("", 651, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q1 = GUICtrlCreateLabel("", 651, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q1 = GUICtrlCreateLabel("", 651, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q1 = GUICtrlCreateLabel("", 770, 30, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q1 = GUICtrlCreateLabel("", 770, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q1 = GUICtrlCreateLabel("", 770, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q2x = GUICtrlCreateGroup(" Q2 - " & @YEAR, 855, 8, 270, 100)

$Label_1_q2 = GUICtrlCreateLabel("Total Days:", 856, 30, 79, 21, $SS_RIGHT)
$Label_2_q2 = GUICtrlCreateLabel("Work Days:", 856, 50, 79, 21, $SS_RIGHT)
$Label_3_q2 = GUICtrlCreateLabel("Ratio:", 856, 70, 79, 21, $SS_RIGHT)
$Label_Ratio_q2 = GUICtrlCreateLabel("Ratio to Date:", 856, 90, 79, 16, $SS_RIGHT)

$Label_4_q2 = GUICtrlCreateLabel("Estim.On-Site:", 985, 30, 65, 21, $SS_RIGHT)
$Label_5_q2 = GUICtrlCreateLabel("Real On-Site:", 985, 50, 65, 21, $SS_RIGHT)
$Label_6_q2 = GUICtrlCreateLabel("Remaining:", 985, 70, 65, 21, $SS_RIGHT)

$Input_TD_q2 = GUICtrlCreateLabel("", 936, 30, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q2 = GUICtrlCreateLabel("", 936, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q2 = GUICtrlCreateLabel("", 936, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q2 = GUICtrlCreateLabel("", 936, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q2 = GUICtrlCreateLabel("", 1055, 30, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q2 = GUICtrlCreateLabel("", 1055, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q2 = GUICtrlCreateLabel("", 1055, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q3 = GUICtrlCreateGroup(" Q3 - " & @YEAR, 570, 108, 270, 100)

$Label_1_q3 = GUICtrlCreateLabel("Total Days:", 571, 130, 79, 21, $SS_RIGHT)
$Label_2_q3 = GUICtrlCreateLabel("Work Days:", 571, 150, 79, 21, $SS_RIGHT)
$Label_3_q3 = GUICtrlCreateLabel("Ratio:", 571, 170, 79, 21, $SS_RIGHT)
$Label_Ratio_q3 = GUICtrlCreateLabel("Ratio to Date:", 571, 190, 79, 16, $SS_RIGHT)

$Label_4_q3 = GUICtrlCreateLabel("Estim.On-Site:", 700, 130, 65, 21, $SS_RIGHT)
$Label_5_q3 = GUICtrlCreateLabel("Real On-Site:", 700, 150, 65, 21, $SS_RIGHT)
$Label_6_q3 = GUICtrlCreateLabel("Remaining:", 700, 170, 65, 21, $SS_RIGHT)

$Input_TD_q3 = GUICtrlCreateLabel("", 651, 130, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q3 = GUICtrlCreateLabel("", 651, 150, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q3 = GUICtrlCreateLabel("", 651, 170, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q3 = GUICtrlCreateLabel("", 651, 190, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q3 = GUICtrlCreateLabel("", 770, 130, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q3 = GUICtrlCreateLabel("", 770, 150, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q3 = GUICtrlCreateLabel("", 770, 170, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $Group_Q4 = GUICtrlCreateGroup(" Q4 - " & @YEAR, 855, 108, 270, 100)

$Label_1_q4 = GUICtrlCreateLabel("Total Days:", 856, 130, 79, 21, $SS_RIGHT)
$Label_2_q4 = GUICtrlCreateLabel("Work Days:", 856, 150, 79, 21, $SS_RIGHT)
$Label_3_q4 = GUICtrlCreateLabel("Ratio:", 856, 170, 79, 21, $SS_RIGHT)
$Label_Ratio_q4 = GUICtrlCreateLabel("Ratio to Date:", 856, 190, 79, 16, $SS_RIGHT)

$Label_4_q4 = GUICtrlCreateLabel("Estim.On-Site:", 985, 130, 65, 21, $SS_RIGHT)
$Label_5_q4 = GUICtrlCreateLabel("Real On-Site:", 985, 150, 65, 21, $SS_RIGHT)
$Label_6_q4 = GUICtrlCreateLabel("Remaining:", 985, 170, 65, 21, $SS_RIGHT)

$Input_TD_q4 = GUICtrlCreateLabel("", 936, 130, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q4 = GUICtrlCreateLabel("", 936, 150, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q4 = GUICtrlCreateLabel("", 936, 170, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q4 = GUICtrlCreateLabel("", 936, 190, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q4 = GUICtrlCreateLabel("", 1055, 130, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q4 = GUICtrlCreateLabel("", 1055, 150, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q4 = GUICtrlCreateLabel("", 1055, 170, 40, 15, BitOR($ES_CENTER, $ES_READONLY))


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
GUICtrlSetData($Progress_Splash, 30)
_CriaINI(@YEAR)
GUICtrlSetData($Progress_Splash, 40)
_DBRepair()
GUICtrlSetData($Progress_Splash, 50)
_ReadINI(@YEAR)
GUICtrlSetData($Progress_Splash, 80)
_CheckQuarter()
GUICtrlSetData($Progress_Splash, 90)
Sleep(100)
_AutoBKP()
GUICtrlSetData($Progress_Splash, 95)
_CreateMenu()
Sleep(100)
GUICtrlSetData($Progress_Splash, 100)
$SelDate = GUICtrlRead($Calendar)
$SelDate_slipt = StringSplit($SelDate, "/")

$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
$Status = StringTrimLeft($Status1, 1)

GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

GUISetState(@SW_SHOW, $Form_WorkDays)

GUIDelete($Form_Splash)
FileDelete($sSplashPath)

While 1
	$nMsg = GUIGetMsg()

	For $j = 1 To 12

		If $nMsg = $DBpMenu_Report_simple_Year[$j] And $DBpMenu_Report_simple_Year[$j] <> 0 Then
			$DBpMenu_Report_Date = GUICtrlRead($DBpMenu_Report_simple_Year[$j], 1)
			GenerateWorkdaysReportHTML($DBpMenu_Report_Date, 0)
		EndIf

		If $nMsg = $DBpMenu_Report_detailed_Year[$j] And $DBpMenu_Report_detailed_Year[$j] <> 0 Then
			$DBpMenu_Report_Date = GUICtrlRead($DBpMenu_Report_detailed_Year[$j], 1)
			GenerateWorkdaysReportHTML($DBpMenu_Report_Date, 1)
		EndIf

		If $nMsg = $DBpMenu_Delete_Year[$j] And $DBpMenu_Delete_Year[$j] <> 0 Then
			$DBpMenu_Delete_Date = GUICtrlRead($DBpMenu_Delete_Year[$j], 1)
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Delete Year", "WARNING" & @CRLF & "" & @CRLF & "You are about To delete the year " & $DBpMenu_Delete_Date & " from the database." & @CRLF & "" & @CRLF & "All data associated With this year will be permanently removed And cannot be recovered." & @CRLF & "" & @CRLF & "Are you sure you want To proceed ?", 0, $Form_WorkDays)
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
						MsgBox(262208, "Delete Year", "Year Deleted with Success", 0, $Form_WorkDays)
					Else
						MsgBox(262160, "Year Delete", "An error occurred while attempting to delete this value from the database.", 0, $Form_WorkDays)
					EndIf

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		EndIf

		For $i = 1 To 31
			If $Inputs[$i][$j] <> 0 And $nMsg = $Inputs[$i][$j] Then ;_CalendarRead
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

			If $ContextItem_Tag[$i][$j] <> 0 And $nMsg = $ContextItem_Tag[$i][$j] Then ;_Button_Tag
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_Tag($XV, $XU, $SelDate_slipt[1])
			EndIf


			If $ContextItem_OnSite[$i][$j] <> 0 And $nMsg = $ContextItem_OnSite[$i][$j] Then ;_Button_OnSite
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_OnSite($XV, $XU, $SelDate_slipt[1])
			EndIf

			If $ContextItem_Remote[$i][$j] <> 0 And $nMsg = $ContextItem_Remote[$i][$j] Then ;_Button_Remote
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_Remote($XV, $XU, $SelDate_slipt[1])
			EndIf

			If $ContextItem_Holiday[$i][$j] <> 0 And $nMsg = $ContextItem_Holiday[$i][$j] Then ;_Button_holiday
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_holiday($XV, $XU, $SelDate_slipt[1])
			EndIf

			If $ContextItem_PTO[$i][$j] <> 0 And $nMsg = $ContextItem_PTO[$i][$j] Then ;_Button_PTO
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_PTO($XV, $XU, $SelDate_slipt[1])
			EndIf

			If $ContextItem_Travel[$i][$j] <> 0 And $nMsg = $ContextItem_Travel[$i][$j] Then ;_Button_Travel
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_Travel($XV, $XU, $SelDate_slipt[1])
			EndIf

			If $ContextItem_Sick[$i][$j] <> 0 And $nMsg = $ContextItem_Sick[$i][$j] Then ;_Button_Sick
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_Sick($XV, $XU, $SelDate_slipt[1])
			EndIf

			If $ContextItem_Blank[$i][$j] <> 0 And $nMsg = $ContextItem_Blank[$i][$j] Then ;_Button_Blank
				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				If Number($j) < Number("10") Then
					$XV = "0" & $j
				Else
					$XV = $j
				EndIf

				If Number($i) < Number("10") Then
					$XU = "0" & $i
				Else
					$XU = $i
				EndIf
				_Button_Blank($XV, $XU, $SelDate_slipt[1])
			EndIf

		Next
	Next

	Switch $nMsg
		Case $GUI_EVENT_SECONDARYDOWN
			ConsoleWrite("********************************" & @CRLF)
			Global $mousePosX = MouseGetPos(0)
			Global $mousePosY = MouseGetPos(1)
			ConsoleWrite("$mousePosX :" & $mousePosX & @CRLF)
			ConsoleWrite("$mousePosY :" & $mousePosY & @CRLF)
			ConsoleWrite("********************************" & @CRLF)


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
;~ 				GUICtrlSetData($Input_Tag, $Status)
				GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)
			EndIf

		Case $BkpMenu_help_help

			FileInstall("Help.pdf", $HelpFile, 1)

			If Not FileExists($HelpFile) Then
				MsgBox(262160, "Work Days", "Help file not found in the application folder.", 0, $Form_WorkDays)
			Else
				ShellExecute($HelpFile)
			EndIf
;~ 			FileDelete($HelpFile)

		Case $BkpMenu_help_About
			_About()


		Case $GUI_EVENT_CLOSE
			Exit

			#cs
			Case $Button_CalendtarTag
				$DateToTag = GUICtrlRead($Calendar)
				_CalendarTag($DateToTag)
				_Update($DateToTag)
			#ce

		Case $BkpMenu_Batch
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Batch Import", "**WARNING** Importing data will overwrite any existing records. Do you want to proceed?" & @CRLF & @CRLF & "Check the help file for more details.", 0, $Form_WorkDays)
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					GUICtrlSetData($Calendar, @YEAR & "/" & @MON & "/" & @MDAY)
					_RestoreBackup()
					_CalendarRead()

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		Case $DBpMenu_backup_Data_Holidays
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Holidays Import", "**WARNING** Importing data will overwrite any existing records for the selected dates. Do you want to proceed?", 0, $Form_WorkDays)
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					_ImportHolidays()
					_Reload()
					#cs
					_ClearScreen()
					_ReadColors()
					_CriaINI(@YEAR)
					_ReadINI(@YEAR)
					#ce

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
			_CriaINI(@YEAR)
			_Reload()
			#cs
			_ClearScreen()
			_ReadColors()
			_CriaINI(@YEAR)
			_ReadINI(@YEAR)
			#ce

		Case $Button_Reload
			_Reload()
			#cs
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			_ReadINI($SelDate_slipt[1])
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
			$Status = StringTrimLeft($Status1, 1)
			GUICtrlSetData($Input_Tag, $Status)
			GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)
			#ce

		Case $Button_OnSite
			_Button_OnSite()


		Case $Button_Blank
			_Button_Blank()


		Case $Button_Remote
			_Button_Remote()


		Case $Button_Travel
			_Button_Travel()


		Case $Button_PTO
			_Button_PTO()


		Case $Button_holiday
			_Button_holiday()


		Case $Button_Sick
			_Button_Sick()


		Case $Button_Weekend
			_Button_Weekend()


	EndSwitch

WEnd

Func _DBRepair()

	For $i = 1 To 20
		$sSubKey_Year = RegEnumKey($DB, $i)
		If @error Then ExitLoop
;~ 			ConsoleWrite($DB & "\" & $sSubKey_Year & @CRLF)
		For $j = 1 To 12
			$sSubKey_Month = RegEnumKey($DB & "\" & $sSubKey_Year, $j)
			If @error Then ExitLoop
;~ 			ConsoleWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month & @CRLF)
			For $z = 1 To 31
				$sSubKey_Day = RegEnumVal($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month, $z)
				If @error Then ExitLoop
				$sSubKey_Day_Value = RegRead($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month, $sSubKey_Day)
				If StringInStr($sSubKey_Day_Value, " /n") Then
					RegWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month, $sSubKey_Day, "REG_SZ", StringReplace($sSubKey_Day_Value, " /n", @CRLF))
					ConsoleWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month & "\" & $sSubKey_Day & "\" & $sSubKey_Day_Value & @CRLF)
				EndIf
;~ 				If StringInStr($sSubKey_Day_Value," /n ") Then
;~ 					RegWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month,$sSubKey_Day,"REG_SZ",StringReplace($sSubKey_Day_Value," /n ",@CRLF))
;~ 					ConsoleWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month & "\" & $sSubKey_Day & "\" & $sSubKey_Day_Value & @CRLF)
;~ 				EndIf


			Next
		Next
	Next

EndFunc   ;==>_DBRepair


Func _splash($Mode = "on")

	If $Mode = "on" Then

		Global $Form_Splash = GUICreate("", 640, 360, -1, -1, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW, $WS_EX_LAYERED))
		Global $Pic_Splash = GUICtrlCreatePic(@TempDir & "\splash.jpg", 5, 5, 630, 350)
		Global $Progress_Splash = GUICtrlCreateProgress(104, 288, 430, 17)
		Global $Label_version = GUICtrlCreateLabel(FileGetVersion(@ScriptFullPath), 560, 330, -1, -1, $SS_SIMPLE)
		GUICtrlSetColor($Label_version, 0xFFFFFF)
		GUICtrlSetBkColor($Label_version, 0x5b90b2)
		GUISetState(@SW_SHOW, $Form_Splash)
		Return
	Else
		If $Mode = "off" Then
			GUIDelete($Form_Splash)
			GUISetState(@SW_SHOW, $Form_WorkDays)
			Return
		EndIf
	EndIf

EndFunc   ;==>_splash


Func _Reload()

	$ValueConsole = "_ClearScreen() start..."
	ConsoleWrite($ValueConsole & @CRLF)
	_ClearScreen()
	$ValueConsole = "_ClearScreen() end..."
	ConsoleWrite($ValueConsole & @CRLF)

	$ValueConsole = "_ReadColors() start..."
	ConsoleWrite($ValueConsole & @CRLF)
	_ReadColors()
	$ValueConsole = "_ReadColors() end..."
	ConsoleWrite($ValueConsole & @CRLF)

	$SelDate = GUICtrlRead($Calendar)
	$SelDate_slipt = StringSplit($SelDate, "/")

	$ValueConsole = "_ReadINI() start..."
	ConsoleWrite($ValueConsole & @CRLF)
	_ReadINI($SelDate_slipt[1])
	$ValueConsole = "_ReadINI() end..."
	ConsoleWrite($ValueConsole & @CRLF)

	$SelDate = GUICtrlRead($Calendar)
	$SelDate_slipt = StringSplit($SelDate, "/")
	$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	$Status = StringTrimLeft($Status1, 1)
	GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

	$ValueConsole = "end of function _Reload..."
	ConsoleWrite($ValueConsole & @CRLF)


EndFunc   ;==>_Reload



Func _MenuContextual($U, $V)

	$SelDate = GUICtrlRead($Calendar)
	$SelDate_slipt = StringSplit($SelDate, "/") ;Risco de bug em potencial #### Fique atento ####


	If Number($V) < Number("10") Then
		$XV = "0" & $V
	Else
		$XV = $V
	EndIf

	If Number($U) < Number("10") Then
		$XU = "0" & $U
	Else
		$XU = $U
	EndIf




	$Context[$U][$V] = GUICtrlCreateContextMenu($Inputs[$U][$V])
	$ContextItem_Date[$U][$V] = GUICtrlCreateMenuItem("Date: " & $XV & "/" & $XU & "/" & $SelDate_slipt[1], $Context[$U][$V])
	GUICtrlSetState($ContextItem_Date[$U][$V], $gui_disable)
	$ContextItem_Separator[$U][$V] = GUICtrlCreateMenuItem("", $Context[$U][$V])
	$ContextItem_Tag[$U][$V] = GUICtrlCreateMenuItem("Add/Edit Tag", $Context[$U][$V])
	$ContextItem_Separator[$U][$V] = GUICtrlCreateMenuItem("", $Context[$U][$V])
	$ContextItem_OnSite[$U][$V] = GUICtrlCreateMenuItem("On-Site", $Context[$U][$V])
;~ 	GUICtrlSetColor($ContextItem_OnSite[$U][$V],"0xFF0033")
	$ContextItem_Remote[$U][$V] = GUICtrlCreateMenuItem("Remote", $Context[$U][$V])
	$ContextItem_Holiday[$U][$V] = GUICtrlCreateMenuItem("Holiday", $Context[$U][$V])
	$ContextItem_PTO[$U][$V] = GUICtrlCreateMenuItem("PTO", $Context[$U][$V])
	$ContextItem_Travel[$U][$V] = GUICtrlCreateMenuItem("Travel", $Context[$U][$V])
	$ContextItem_Sick[$U][$V] = GUICtrlCreateMenuItem("Sick", $Context[$U][$V])
	$ContextItem_Blank[$U][$V] = GUICtrlCreateMenuItem("Blank / Weekends", $Context[$U][$V])


EndFunc   ;==>_MenuContextual


Func _Button_Tag($Month = "-1", $Day = "-1", $CYear = "-1")

	$Window_Tag_Pos = WinGetPos("Work Days")

	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf

	$RegReadTag = ""

	$Mouse_Tag_Pos_X_New = $mousePosX - 125
	$Mouse_Tag_Pos_Y_New = $mousePosY

	ConsoleWrite("$Window_Tag_Pos[0]: " & $Window_Tag_Pos[0] & @CRLF)
	ConsoleWrite("$Window_Tag_Pos[1]: " & $Window_Tag_Pos[1] & @CRLF)
	ConsoleWrite("$Window_Tag_Pos[2]: " & $Window_Tag_Pos[2] & @CRLF)
	ConsoleWrite("$Window_Tag_Pos[3]: " & $Window_Tag_Pos[3] & @CRLF)
	ConsoleWrite("$Mouse_Tag_Pos_X_New: " & $Mouse_Tag_Pos_X_New & @CRLF)
	ConsoleWrite("$Mouse_Tag_Pos_Y_New: " & $Mouse_Tag_Pos_Y_New & @CRLF)

	If $Mouse_Tag_Pos_X_New + 200 > $Window_Tag_Pos[0] + $Window_Tag_Pos[2] Then
		ConsoleWrite("XXXXXXXXXXXXXXXXXX" & @CRLF)
		$Mouse_Tag_Pos_X_New_calc = ($Mouse_Tag_Pos_X_New + 200) - ($Window_Tag_Pos[0] + $Window_Tag_Pos[2])
		ConsoleWrite("$Mouse_Tag_Pos_X_New_calc: " & $Mouse_Tag_Pos_X_New_calc & @CRLF)
		$Mouse_Tag_Pos_X_New = ($Mouse_Tag_Pos_X_New - $Mouse_Tag_Pos_X_New_calc) - 70
		ConsoleWrite("New position $Mouse_Tag_Pos_X_New: " & $Mouse_Tag_Pos_X_New & @CRLF)
		ConsoleWrite("XXXXXXXXXXXXXXXXXX" & @CRLF)
	Else
		$Mouse_Tag_Pos_X_New = $mousePosX - 20
	EndIf


	If $Mouse_Tag_Pos_Y_New + 150 > $Window_Tag_Pos[1] + $Window_Tag_Pos[3] Then
		ConsoleWrite("YYYYYYYYYYYYYYYYYY" & @CRLF)
		$Mouse_Tag_Pos_Y_New_calc = ($Mouse_Tag_Pos_Y_New + 200) - ($Window_Tag_Pos[1] + $Window_Tag_Pos[3])
		ConsoleWrite("$Mouse_Tag_Pos_Y_New_calc: " & $Mouse_Tag_Pos_Y_New_calc & @CRLF)
		$Mouse_Tag_Pos_Y_New = ($Mouse_Tag_Pos_Y_New - $Mouse_Tag_Pos_Y_New_calc) ; - 70
		ConsoleWrite("New position $Mouse_Tag_Pos_Y_New: " & $Mouse_Tag_Pos_Y_New & @CRLF)
		ConsoleWrite("YYYYYYYYYYYYYYYYYY" & @CRLF)
	Else
		$Mouse_Tag_Pos_Y_New = $mousePosY
	EndIf


	$SelDate_slipt = StringSplit($SelDate, "/")
	$holidayName = GUICtrlRead($Input_Tag)
	$RegReadTag = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])

	If $RegReadTag <> "" Then
		If _CheckDateReturn($SelDate) <> "" Then
			$RegReadTag = StringTrimLeft($RegReadTag, 1)
		EndIf
	EndIf

	Global $Form_Tag = GUICreate("Add/Edit Tag", 249, 181, $Mouse_Tag_Pos_X_New, $Mouse_Tag_Pos_Y_New, BitOR($WS_BORDER, $WS_POPUP, $DS_SETFOREGROUND, $DS_MODALFRAME), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW), $Form_WorkDays)
	$Label_Tag = GUICtrlCreateLabel("Selected Date (YYYY/MM/DD): " & $CYear & "/" & $Month & "/" & $Day, 8, 10, 233, 105)
	$Edit_Tag = GUICtrlCreateEdit("", 8, 38, 233, 105,BitOR($ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL, $ES_AUTOVSCROLL, $ES_AUTOHSCROLL,$ES_NOHIDESEL))
	$Button_Tag_Cancel = GUICtrlCreateButton("Cancel", 8, 150, 75, 25)
	$Button_Tag_Save = GUICtrlCreateButton("Save", 165, 150, 75, 25,$BS_DEFPUSHBUTTON)



	GUICtrlSetData($Edit_Tag, $RegReadTag,1)


	GUISetState(@SW_SHOW)


	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($Form_Tag)
				Return
			Case $Button_Tag_Cancel
				GUIDelete($Form_Tag)
				Return
			Case $Button_Tag_Save
				$DateToTag = $CYear & "/" & $Month & "/" & $Day
;~ 				MsgBox(262144,"",$DateToTag & @CRLF & $CYear & "/" & $Month & "/" & $Day)
				$SelDate_slipt = StringSplit($DateToTag, "/")
				$holidayName = GUICtrlRead($Edit_Tag)
				$Register = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
				If $Register = "" Then $Register = "B"
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", StringLeft($Register, 1) & $holidayName)
				_Update($SelDate)

;~ 	 			_CalendarTag($DateToTag)
				_Update($DateToTag)
				GUIDelete($Form_Tag)
				Return

		EndSwitch
	WEnd


	#cs
	$CheckDate_Return = _CheckDate($SelDate, "O")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "O" & $holidayName)
		_Update($SelDate)
	EndIf
	#ce
EndFunc   ;==>_Button_Tag


Func _Button_OnSite($Month = "-1", $Day = "-1", $CYear = "-1")
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf

	$CheckDate_Return = _CheckDate($SelDate, "O")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "O" & $holidayName)
		_Update($SelDate)
	EndIf
EndFunc   ;==>_Button_OnSite

Func _Button_Blank($Month = "-1", $Day = "-1", $CYear = "-1")
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
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

EndFunc   ;==>_Button_Blank

Func _Button_Remote($Month = "-1", $Day = "-1", $CYear = "-1")
	ConsoleWrite("##### ----->>>>> Remote" & @CRLF)
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
	$CheckDate_Return = _CheckDate($SelDate, "R")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "R" & $holidayName)
		_Update($SelDate)
	EndIf
EndFunc   ;==>_Button_Remote

Func _Button_Travel($Month = "-1", $Day = "-1", $CYear = "-1")
	ConsoleWrite("##### ----->>>>> Travel" & @CRLF)
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
	$CheckDate_Return = _CheckDate($SelDate, "T")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "T" & $holidayName)
		_Update($SelDate)
	EndIf
EndFunc   ;==>_Button_Travel

Func _Button_PTO($Month = "-1", $Day = "-1", $CYear = "-1")
	ConsoleWrite("##### ----->>>>> PTO" & @CRLF)
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
	$CheckDate_Return = _CheckDate($SelDate, "P")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "P" & $holidayName)
		_Update($SelDate)
	EndIf
EndFunc   ;==>_Button_PTO

Func _Button_holiday($Month = "-1", $Day = "-1", $CYear = "-1")
	ConsoleWrite("##### ----->>>>> Holiday" & @CRLF)
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
	$CheckDate_Return = _CheckDate($SelDate, "H")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "H" & $holidayName)
		_Update($SelDate)
	EndIf
EndFunc   ;==>_Button_holiday

Func _Button_Sick($Month = "-1", $Day = "-1", $CYear = "-1")
	ConsoleWrite("##### ----->>>>> Sick" & @CRLF)
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
	$CheckDate_Return = _CheckDate($SelDate, "S")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")
		$holidayName = GUICtrlRead($Input_Tag)
		RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "S" & $holidayName)
		_Update($SelDate)
	EndIf
EndFunc   ;==>_Button_Sick

Func _Button_Weekend($Month = "-1", $Day = "-1", $CYear = "-1")
	ConsoleWrite("##### ----->>>>> Weekend" & @CRLF)
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf
	$CheckDate_Return = _CheckDate($SelDate, "W")
	If $CheckDate_Return = 0 Then
		$SelDate_slipt = StringSplit($SelDate, "/")

		$WeekDayNum = _DateToDayOfWeek($SelDate_slipt[1], $SelDate_slipt[2], $SelDate_slipt[3])

		$WeekEnd = 0
		If $WeekDayNum <> "1" And $WeekDayNum <> "7" Then
			$WeekEnd = 1
		EndIf

		If $WeekEnd = 1 Then
			MsgBox(262160, "Weekend", "This date is not a weekend.", 0, $Form_WorkDays)
		Else
			$holidayName = GUICtrlRead($Input_Tag)
			RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", "W" & $holidayName)
			_Update($SelDate)
		EndIf
	EndIf
EndFunc   ;==>_Button_Weekend

Func _CreateMenu()

	GUICtrlDelete($DBpMenu_Report_Simple)
	GUICtrlDelete($DBpMenu_Report_Detailed)
	GUICtrlDelete($DBpMenu_Delete)
	GUICtrlDelete($BkpMenu_Exit)

	Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
	Global $DBpMenu_Report_Simple = GUICtrlCreateMenu("Simple", $DBpMenu_Report)
	Global $DBpMenu_Report_Detailed = GUICtrlCreateMenu("Detailed", $DBpMenu_Report)

	Local $sSubKey = ""
	For $i = 1 To 12

		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

		$DBpMenu_Delete_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Delete)
		$DBpMenu_Report_simple_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Report_Simple)
		$DBpMenu_Report_detailed_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Report_Detailed)

	Next

	Global $BkpMenu_Exit = GUICtrlCreateMenuItem("&Exit", $DBpMenu_db)

EndFunc   ;==>_CreateMenu

Func _CheckDateReturn($DateToCheck)

	$DateToCheck_split = StringSplit($DateToCheck, "/")

	$DateToCheck_Value = RegRead($DB & "\" & $DateToCheck_split[1] & "\" & $DateToCheck_split[2], $DateToCheck_split[3])

	$DateToCheck_Value = StringLeft($DateToCheck_Value, 1)

	Return $DateToCheck_Value

EndFunc   ;==>_CheckDateReturn

Func _CheckDate($DateToCheck, $NewStatus)

	$DateToCheck_split = StringSplit($DateToCheck, "/")

	$DateToCheck_Value = RegRead($DB & "\" & $DateToCheck_split[1] & "\" & $DateToCheck_split[2], $DateToCheck_split[3])

	If $NewStatus = "" Then
		$WeekDayNum = _DateToDayOfWeek($DateToCheck_split[1], $DateToCheck_split[2], $DateToCheck_split[3])
		If $WeekDayNum = "1" Or $WeekDayNum = "7" Then
			$NewStatus = "W"
		EndIf
	EndIf

	$DateToCheck_Value = StringLeft($DateToCheck_Value, 1)

	If $DateToCheck_Value <> "" And $DateToCheck_Value <> "B" And StringLeft($DateToCheck_Value, 1) <> $NewStatus Then
		If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(262436, "Replace current value", "You're about to replace the current status for the selected date. " & @CRLF & @CRLF & "Current Status: " & _Label(StringLeft($DateToCheck_Value, 1)) & @CRLF & "New Status: " & _Label($NewStatus) & @CRLF & @CRLF & "Do you want to continue?", 0, $Form_WorkDays)
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
		If @error = 1 Then
			Return
		Else

			MsgBox(262160, "Import", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: 1." & @error, 0, $Form_WorkDays)
		EndIf
	Else
		$FileHolidays_hwd = FileOpen($HolidaysFile, 0)
		If $FileHolidays_hwd = -1 Then
			MsgBox(262160, "Import", "Oops! Something went wrong when read the file. Please try again." & @CRLF & "Error code: 2." & @error, 0, $Form_WorkDays)
			Return

		Else

			$ResetReturn = _ResetDatabase("1")
			If $ResetReturn = "1" Then
				_CriaINI(@YEAR)
				While 1
					$HolidaysLine = FileReadLine($FileHolidays_hwd)
					If @error = -1 Then ExitLoop
					If @error = 1 Then
						MsgBox(262160, "Import", "Oops! Something went wrong when read the file. Please try again." & @CRLF & "Error code: 3." & @error, 0, $Form_WorkDays)
						Return
					EndIf
					If Not StringInStr($HolidaysLine, "\") Then
						If Not StringInStr($HolidaysLine, "=") Then
							$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
						Else
							$HolidaysLine_Setting = StringSplit($HolidaysLine, "=")
							$RegError = RegWrite($DB, $HolidaysLine_Setting[1], "REG_SZ", StringReplace($HolidaysLine_Setting[2], " /n", @CRLF))
							If @error Then
								$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
							Else
								$ImportCount += 1
							EndIf
						EndIf
					Else
						$HolidaysLine_key = StringSplit($HolidaysLine, "\")
						$HolidaysLine_Value = StringSplit($HolidaysLine_key[3], "=")
						$RegError = RegWrite($DB & "\" & $HolidaysLine_key[1] & "\" & $HolidaysLine_key[2], $HolidaysLine_Value[1], "REG_SZ", StringReplace($HolidaysLine_Value[2], " /n", @CRLF))
						If @error Then
							$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
						Else
							$ImportCount += 1
						EndIf
					EndIf

				WEnd

				If $HolidaysError <> "" Then
					_DBRepair()
					_Reload()
					MsgBox(262160, "Import", "Oops! Something went wrong when read the file." & @CRLF & "The following lines was not imported:" & @CRLF & @CRLF & $HolidaysError & @CRLF & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess, 0, $Form_WorkDays)
				Else
					If $ImportCount > 15 Then
						_DBRepair()
						_Reload()
						MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & @CRLF & $ImportCount & " lines imported.", 0, $Form_WorkDays)

					Else
						_DBRepair()
						_Reload()
						MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess, 0, $Form_WorkDays)

					EndIf
				EndIf

			Else

				_DBRepair()
				_Reload()
				MsgBox(262160, "Reset Database", "Oops! Something went wrong. Please try again.", 0, $Form_WorkDays)
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
		$tip = "- " & StringTrimLeft($Data_Register1, 1)
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

	If $Data_Register = "W" Then $StatusName = "WEEKEND"
	If $Data_Register = "O" Then $StatusName = "ON-SITE"
	If $Data_Register = "R" Then $StatusName = "REMOTE"
	If $Data_Register = "T" Then $StatusName = "TRAVEL"
	If $Data_Register = "P" Then $StatusName = "PTO"
	If $Data_Register = "H" Then $StatusName = "HOLIDAY"
	If $Data_Register = "S" Then $StatusName = "SICK DAY"
	If $Data_Register = "B" Then $StatusName = "BLANK"
	If $Data_Register = "" Then $StatusName = "BLANK"
	If $Data_Register = "   " Then $StatusName = "BLANK"

	ConsoleWrite("$Data_Register: %" & $Data_Register & "%" & @CRLF)

	If $tip <> "" Then
		GUICtrlSetTip($Inputs[$Data_day][$Data_month], StringReplace($tip, @CRLF, @CRLF & "- "), $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & " - " & $StatusName) ; & " - TODAY")
	Else
		GUICtrlSetTip($Inputs[$Data_day][$Data_month], $StatusName, $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day) ; & " - TODAY")
	EndIf



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

;~ 		GUICtrlSetTip($Inputs[$Data_day][$Data_month], StringReplace($tip, "/n", @CRLF & "-"), $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & " - TODAY")

		If $tip <> "" Then
			GUICtrlSetTip($Inputs[$Data_day][$Data_month], StringReplace($tip, @CRLF, @CRLF & "- "), $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & " - " & $StatusName & " - TODAY")
		Else
			GUICtrlSetTip($Inputs[$Data_day][$Data_month], $StatusName, $WeekDayName & " - " & $Data_year & "/" & $Data_month & "/" & $Data_day & " - TODAY")
		EndIf

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
		If _DateDiff('D', $AutoSaveDate[0] & "/" & $AutoSaveDate[1] & "/" & $AutoSaveDate[2], @YEAR & "/" & @MON & "/" & @MDAY) > 1 Then
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
		MsgBox(262160, "Import", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error, 0, $Form_WorkDays)
	Else
		$FileHolidays_hwd = FileOpen($HolidaysFile, 0)
		If $FileHolidays_hwd <> -1 Then

			While 1
				$HolidaysLine = FileReadLine($FileHolidays_hwd)
				If @error = -1 Then ExitLoop
				If @error = 1 Then
					MsgBox(262160, "Import", "Oops! Something went wrong when read the file. Please try again." & @CRLF & "Error code: " & @error, 0, $Form_WorkDays)
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
				MsgBox(262160, "Import", "Oops! Something went wrong when read the file." & @CRLF & "The following lines was not imported:" & @CRLF & @CRLF & $HolidaysError & @CRLF & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess, 0, $Form_WorkDays)
			Else
				If $ImportCount > 10 Then
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & @CRLF & $ImportCount & " lines imported.", 0, $Form_WorkDays)
				Else
					MsgBox(262208, "Import", "**Success!** The command was executed successfully." & @CRLF & "The following lines was imported:" & @CRLF & @CRLF & $HolidaysSucess, 0, $Form_WorkDays)
				EndIf
			EndIf
		EndIf
		_CreateMenu()
	EndIf

EndFunc   ;==>_ImportHolidays

Func _ResetDatabase($step = "0")

	$sKey = $DB & "\"
	If $step = "0" Then
		If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(262452, "Reset Database", "**Warning!** " & @CRLF & "Are you sure you want to permanently delete all data from the database? This action cannot be undone.", 0, $Form_WorkDays)
		Select
			Case $iMsgBoxAnswer = 6 ;Yes
				RegDelete($sKey)
				If @error Then
					_CreateMenu()
					MsgBox(262160, "Reset Database", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error, 0, $Form_WorkDays)
					Return
				Else
					_CreateMenu()
					MsgBox(262208, "Reset Database", "**Success!** The command was executed successfully. All data has been removed.", 0, $Form_WorkDays)
					Return
				EndIf

			Case $iMsgBoxAnswer = 7 ;No
				_CreateMenu()
				Return

		EndSelect
	Else
		RegDelete($sKey)
		If @error Then
			_CreateMenu()
			MsgBox(262160, "Reset Database", "Oops! Something went wrong. Please try again." & @CRLF & "Error code: " & @error, 0, $Form_WorkDays)
			Return 0
		Else
			_CreateMenu()
			Return 1
		EndIf

	EndIf


	Return


EndFunc   ;==>_ResetDatabase

Func _CalendarRead($i = 0, $j = 0)

	For $a = 1 To 12
		For $b = 1 To 31
			If GUICtrlGetState($SelectLabel[$b][$a]) = 144 Then
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

	GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

	$Status_Tip = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	GUICtrlSetData($Input_Tag, StringTrimLeft($Status_Tip, 1))

	Return

EndFunc   ;==>_CalendarRead

Func _ClearScreen()

	For $j = 1 To 12
		For $i = 1 To 31
			GUICtrlDelete($Inputs[$i][$j])

			If $Debug = 9 Then
				GUICtrlDelete($Context[$i][$j])
				GUICtrlDelete($ContextItem_Date[$i][$j])
				GUICtrlDelete($ContextItem_Tag[$i][$j])
				GUICtrlDelete($ContextItem_OnSite[$i][$j])
				GUICtrlDelete($ContextItem_Remote[$i][$j])
				GUICtrlDelete($ContextItem_Holiday[$i][$j])
				GUICtrlDelete($ContextItem_PTO[$i][$j])
				GUICtrlDelete($ContextItem_Travel[$i][$j])
				GUICtrlDelete($ContextItem_Sick[$i][$j])
				GUICtrlDelete($ContextItem_Blank[$i][$j])
			EndIf

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
					$tip = "- " & StringTrimLeft($Status1, 1)
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

	GUICtrlSetData($Input_R_Onsite_q1, Round($Counta_R_Onsite_q1, 2)) ;## Real On-Site ##
	GUICtrlSetData($Input_R_Onsite_q2, Round($Counta_R_Onsite_q2, 2))
	GUICtrlSetData($Input_R_Onsite_q3, Round($Counta_R_Onsite_q3, 2))
	GUICtrlSetData($Input_R_Onsite_q4, Round($Counta_R_Onsite_q4, 2))

	$Remaining_q1 = Ceiling(($Counta_WD_q1 / 5) * 3) - $Counta_R_Onsite_q1
	$Remaining_q2 = Ceiling(($Counta_WD_q2 / 5) * 3) - $Counta_R_Onsite_q2
	$Remaining_q3 = Ceiling(($Counta_WD_q3 / 5) * 3) - $Counta_R_Onsite_q3
	$Remaining_q4 = Ceiling(($Counta_WD_q4 / 5) * 3) - $Counta_R_Onsite_q4

	GUICtrlSetData($Input_Remaining_q1, $Remaining_q1) ;## Remaining ##
	GUICtrlSetData($Input_Remaining_q2, $Remaining_q2)
	GUICtrlSetData($Input_Remaining_q3, $Remaining_q3)
	GUICtrlSetData($Input_Remaining_q4, $Remaining_q4)

	$Ratio_R_Q1 = Round(($Counta_R_Onsite_q1 / ($Counta_WD_q1 / 5)), 2)
	$Ratio_R_Q2 = Round(($Counta_R_Onsite_q2 / ($Counta_WD_q2 / 5)), 2)
	$Ratio_R_Q3 = Round(($Counta_R_Onsite_q3 / ($Counta_WD_q3 / 5)), 2)
	$Ratio_R_Q4 = Round(($Counta_R_Onsite_q4 / ($Counta_WD_q4 / 5)), 2)

	GUICtrlSetData($Input_RT_q1, $Ratio_R_Q1) ; ## Ration ##
	GUICtrlSetBkColor($Input_RT_q1, _GetColorGradient($Ratio_R_Q1))


	GUICtrlSetData($Input_RT_q2, $Ratio_R_Q2)
	GUICtrlSetBkColor($Input_RT_q2, _GetColorGradient($Ratio_R_Q2))


	GUICtrlSetData($Input_RT_q3, $Ratio_R_Q3)
	GUICtrlSetBkColor($Input_RT_q3, _GetColorGradient($Ratio_R_Q3))


	GUICtrlSetData($Input_RT_q4, $Ratio_R_Q4)
	GUICtrlSetBkColor($Input_RT_q4, _GetColorGradient($Ratio_R_Q4))



	$Ratio_Q1 = Round(($Counta_R_Onsite_Quarter_Q1 / ($Counta_WD_Quarter_Q1 / 5)), 2)
	$Ratio_Q2 = Round(($Counta_R_Onsite_Quarter_Q2 / ($Counta_WD_Quarter_Q2 / 5)), 2)
	$Ratio_Q3 = Round(($Counta_R_Onsite_Quarter_Q3 / ($Counta_WD_Quarter_Q3 / 5)), 2)
	$Ratio_Q4 = Round(($Counta_R_Onsite_Quarter_Q4 / ($Counta_WD_Quarter_Q4 / 5)), 2)

	GUICtrlSetData($Input_RaTio_q1, "")
	GUICtrlSetData($Input_RaTio_q2, "")
	GUICtrlSetData($Input_RaTio_q3, "")
	GUICtrlSetData($Input_RaTio_q4, "")

	ConsoleWrite(@CRLF & _
			"Dias úteis em Q2 (total): " & $Counta_WD_q3 & @CRLF & _
			"Dias úteis em Q2 (to date): " & $Counta_WD_Quarter_Q3 & @CRLF & _
			"Dias on-site Q2 (total): " & $Counta_R_Onsite_q3 & @CRLF & _
			"Dias on-site Q2 (to date): " & $Counta_R_Onsite_Quarter_Q3 & @CRLF & _
			"Ratio Q2 (total): " & $Ratio_R_Q3 & @CRLF & _
			"Ratio Q2 (to date): " & $Ratio_Q3 & @CRLF)




	If $Year = @YEAR Then
		If @MON = "01" Or @MON = "02" Or @MON = "03" Then
			If $Counta_WD_Quarter_Q1 < 4 Then
				$Ratio_Q1 = "-"
				$Counta_WD_Quarter_Q1 = $Counta_WD_Quarter_Q1 & @CRLF & "Insufficient data to generate a reliable metric."
			EndIf
			GUICtrlSetData($Input_RaTio_q1, $Ratio_Q1)
			GUICtrlSetBkColor($Input_RaTio_q1, _GetColorGradient($Ratio_Q1))
;~ 			GUICtrlSetState($Input_RaTio_q1,$gui_disable)
;~ 			GUICtrlSetTip($Input_RaTio_q1, "Work Days to date: " & $Counta_WD_Quarter_Q1)
;~ 			GUICtrlSetData($Input_WD_q1,$Counta_WD_q1 & "/" & $Counta_WD_q1 - $Counta_WD_Quarter_Q1)
;~ 			GUICtrlSetTip($Input_WD_q1, "Work Days Remaining: " & $Counta_WD_q1 - $Counta_WD_Quarter_Q1)
			GUICtrlSetData($Input_WD_q1, $Counta_WD_Quarter_Q1 & "/" & $Counta_WD_q1)
			GUICtrlSetTip($Input_WD_q1, "Work Days to date: " & $Counta_WD_Quarter_Q1 & @CRLF & "Work Days Remaining: " & $Counta_WD_q1 - $Counta_WD_Quarter_Q1 & @CRLF & "Work Days Total: " & $Counta_WD_q1)

;~ 			GUICtrlSetData($Input_WD_q1, $Counta_WD_q1) ;## Work Days ##
		EndIf

		If @MON = "04" Or @MON = "05" Or @MON = "06" Then
			If $Counta_WD_Quarter_Q2 < 4 Then
				$Ratio_Q2 = "-"
				$Counta_WD_Quarter_Q2 = $Counta_WD_Quarter_Q2 & @CRLF & "Insufficient data to generate a reliable metric."
			EndIf
			GUICtrlSetData($Input_RaTio_q2, $Ratio_Q2)
			GUICtrlSetBkColor($Input_RaTio_q2, _GetColorGradient($Ratio_Q2))
;~ 			GUICtrlSetState($Input_RaTio_q2,$gui_disable)
;~ 			GUICtrlSetTip($Input_RaTio_q2, "Work Days to date: " & $Counta_WD_Quarter_Q2)
;~ 			GUICtrlSetData($Input_WD_q2,$Counta_WD_q2 & "/" & $Counta_WD_q2 - $Counta_WD_Quarter_Q2)
;~ 			GUICtrlSetTip($Input_WD_q2, "Work Days Remaining: " & $Counta_WD_q2 - $Counta_WD_Quarter_Q2)
			GUICtrlSetData($Input_WD_q2, $Counta_WD_Quarter_Q2 & "/" & $Counta_WD_q2)
			GUICtrlSetTip($Input_WD_q2, "Work Days to date: " & $Counta_WD_Quarter_Q2 & @CRLF & "Work Days Remaining: " & $Counta_WD_q2 - $Counta_WD_Quarter_Q2 & @CRLF & "Work Days Total: " & $Counta_WD_q2)
		EndIf

		If @MON = "07" Or @MON = "08" Or @MON = "09" Then
			If $Counta_WD_Quarter_Q3 < 4 Then
				$Ratio_Q3 = "-"
				$Counta_WD_Quarter_Q3 = $Counta_WD_Quarter_Q3 & @CRLF & "Insufficient data to generate a reliable metric."
			EndIf
			GUICtrlSetData($Input_RaTio_q3, $Ratio_Q3)
			GUICtrlSetBkColor($Input_RaTio_q3, _GetColorGradient($Ratio_Q3))
;~ 			GUICtrlSetState($Input_RaTio_q3,$gui_disable)
;~ 			GUICtrlSetTip($Input_RaTio_q3, "Work Days to date: " & $Counta_WD_Quarter_Q3)
;~ 			GUICtrlSetData($Input_WD_q3,$Counta_WD_q3 & "/" & $Counta_WD_q3 - $Counta_WD_Quarter_Q3)
;~ 			GUICtrlSetTip($Input_WD_q3, "Work Days Remaining: " & $Counta_WD_q3 - $Counta_WD_Quarter_Q3)
			GUICtrlSetData($Input_WD_q3, $Counta_WD_Quarter_Q3 & "/" & $Counta_WD_q3)
			GUICtrlSetTip($Input_WD_q3, "Work Days to date: " & $Counta_WD_Quarter_Q3 & @CRLF & "Work Days Remaining: " & $Counta_WD_q3 - $Counta_WD_Quarter_Q3 & @CRLF & "Work Days Total: " & $Counta_WD_q3)
		EndIf

		If @MON = "10" Or @MON = "11" Or @MON = "12" Then
			If $Counta_WD_Quarter_Q4 < 4 Then
				$Ratio_Q4 = "-"
				$Counta_WD_Quarter_Q4 = $Counta_WD_Quarter_Q4 & @CRLF & "Insufficient data to generate a reliable metric."
			EndIf

			GUICtrlSetData($Input_RaTio_q4, $Ratio_Q4)
			GUICtrlSetBkColor($Input_RaTio_q4, _GetColorGradient($Ratio_Q4))
;~ 			GUICtrlSetState($Input_RaTio_q4,$gui_disable)
			#cs
			GUICtrlSetTip($Input_RaTio_q4, "Work Days to date: " & $Counta_WD_Quarter_Q4)
			GUICtrlSetData($Input_WD_q4,$Counta_WD_q4 & "/" & $Counta_WD_q4 - $Counta_WD_Quarter_Q4)
			GUICtrlSetTip($Input_WD_q4, "Work Days Remaining: " & $Counta_WD_q4 - $Counta_WD_Quarter_Q4)
			#ce
			GUICtrlSetData($Input_WD_q4, $Counta_WD_Quarter_Q4 & "/" & $Counta_WD_q4)
			GUICtrlSetTip($Input_WD_q4, "Work Days to date: " & $Counta_WD_Quarter_Q4 & @CRLF & "Work Days Remaining: " & $Counta_WD_q4 - $Counta_WD_Quarter_Q4 & @CRLF & "Work Days Total: " & $Counta_WD_q4)

		EndIf
	EndIf

	_CheckQuarter()

	Return

EndFunc   ;==>_ReadStatistics


Func _ReadINI($iYear)

	GUICtrlSetData($Input_Tag, "")

	_ClearScreen()



	_ReadStatistics($iYear)

	Local Const $iMaxDays = 31
	Local Const $iMaxMonths = 12

	Local Static $aDayCaption[32]
	Local Static $aMonthCaption[13]

	Local $i, $j

	; Pre-compute "01".."31"
	If $aDayCaption[1] = "" Then
		For $i = 1 To $iMaxDays
			$aDayCaption[$i] = StringFormat("%02d", $i)
		Next
	EndIf

	; Pre-compute "01".."12"
	If $aMonthCaption[1] = "" Then
		For $j = 1 To $iMaxMonths
			$aMonthCaption[$j] = StringFormat("%02d", $j)
		Next
	EndIf

	; Top day labels
	For $i = 1 To $iMaxDays
		$LabelMonth[$i] = GUICtrlCreateLabel($aDayCaption[$i], 5 + ($i * 35), 216, 20, 20, $SS_CENTER)
	Next

	Local $C = 0, $Skip = 0
	Local $sMonthStr, $iDaysInMonth
	Local $Status, $Status1, $StatusName, $tip
	Local $WeekDayNum, $WeekDayName
	Local $sDayStr
	Local $bCurrentYear = ($iYear = @YEAR)

	For $j = 1 To $iMaxMonths

		$sMonthStr = $aMonthCaption[$j]

		; *** FIXED HERE ***
		; How many days in this month
		$iDaysInMonth = _DateDaysInMonth($iYear, $j)
		If @error Or $iDaysInMonth = 0 Then ContinueLoop

		; Month name (e.g. January)
		Local $sMonthName = _DateToMonth($sMonthStr, 1)
		If @error Then ContinueLoop

		; Month label on the left
		$LabelMonthX[$j] = GUICtrlCreateLabel($sMonthName, 8, 208 + $Skip + ($j * 25), 20, 20, $SS_CENTER)

		For $i = 1 To $iDaysInMonth

			$sDayStr = $aDayCaption[$i]

			; Select overlay label
			$SelectLabel[$i][$j] = GUICtrlCreateLabel("", -3 + ($i * 35), 202 + $Skip + ($j * 25), 36, 28)
			GUICtrlSetBkColor($SelectLabel[$i][$j], $Color_bk_Selected)
			GUICtrlSetState($SelectLabel[$i][$j], BitOR($GUI_DISABLE, $GUI_HIDE))

			; Today overlay label
			$TodayLabel[$i][$j] = GUICtrlCreateLabel("", -1 + ($i * 35), 204 + $Skip + ($j * 25), 32, 24)
			GUICtrlSetBkColor($TodayLabel[$i][$j], $Color_bk_Today)
			GUICtrlSetColor($TodayLabel[$i][$j], $Color_bk_Today)
			GUICtrlSetState($TodayLabel[$i][$j], BitOR($GUI_DISABLE, $GUI_HIDE))

			; Main clickable cell
			$Inputs[$i][$j] = GUICtrlCreateButton("", 0 + ($i * 35), 205 + $Skip + ($j * 25), 30, 22, _
					BitOR($ES_READONLY, $ES_CENTER, $BS_FLAT, $BS_BOTTOM))

			_MenuContextual($i, $j)

			; Read registry once
			$Status1 = RegRead($DB & "\" & $iYear & "\" & $sMonthStr, $sDayStr)
			If @error Then
				GUICtrlDelete($Inputs[$i][$j])
				ContinueLoop
			EndIf

			; Keep your original IniSection logic (if used elsewhere)
			$IniSection[$j][$i] = RegEnumVal($DB & "\" & $iYear & "\" & $sMonthStr, $i)
			; if you need the old behaviour of using $sDayStr ("01"), change $i -> $sDayStr above

			; Decode status / notes
			$Status = StringLeft($Status1, 1)

			If StringLen($Status1) > 1 Then
				$tip = "- " & StringTrimLeft($Status1, 1)
				GUICtrlSetFont($Inputs[$i][$j], 9, 900, 6, "", 2)
			Else
				$tip = ""
				GUICtrlSetFont($Inputs[$i][$j], 9, 100, 0, "", 2)
			EndIf

			Switch $Status
				Case "W"
					$StatusName = "WEEKEND"
				Case "O"
					$StatusName = "ON-SITE"
				Case "R"
					$StatusName = "REMOTE"
				Case "T"
					$StatusName = "TRAVEL"
				Case "P"
					$StatusName = "PTO"
				Case "H"
					$StatusName = "HOLIDAY"
				Case "S"
					$StatusName = "SICK DAY"
				Case "B", ""
					$StatusName = "BLANK"
				Case Else
					$StatusName = "BLANK"
			EndSwitch

			; weekday info
			$WeekDayNum = _DateToDayOfWeek($iYear, $sMonthStr, $i)
			$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)

			If $tip <> "" Then
				GUICtrlSetTip($Inputs[$i][$j], StringReplace($tip, @CRLF, @CRLF & "- "), _
						$WeekDayName & " - " & $iYear & "/" & $sMonthStr & "/" & $sDayStr & " - " & $StatusName)
			Else
				GUICtrlSetTip($Inputs[$i][$j], $StatusName, _
						$WeekDayName & " - " & $iYear & "/" & $sMonthStr & "/" & $sDayStr)
			EndIf

			; blank display logic
			If $Status = "B" Then
				If $tip <> "" Then
					$Status = "   "
				Else
					$Status = ""
				EndIf
			EndIf

			GUICtrlSetData($Inputs[$i][$j], $Status)

			; Colors / font
			Switch $Status
				; Weekend
				Case "W"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Weekend)
					If $Picker_Font_Weekend_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; On-Site
				Case "O"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_OnSite)
					If $Picker_Font_OnSite_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; Remote
				Case "R"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Remote)
					If $Picker_Font_Remote_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; Travel
				Case "T"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Travel)
					If $Picker_Font_Travel_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; PTO
				Case "P"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_PTO)
					If $Picker_Font_PTO_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; Holiday
				Case "H"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_holiday)
					If $Picker_Font_Holiday_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; Sick Day
				Case "S"
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Sick)
					If $Picker_Font_Sick_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

					; Blank / empty
				Case "", "   "
					GUICtrlSetBkColor($Inputs[$i][$j], $Color_bk_Blank)
					If $Picker_Font_Blank_Read = 1 Then
						GUICtrlSetColor($Inputs[$i][$j], $White)
					Else
						GUICtrlSetColor($Inputs[$i][$j], $Black)
					EndIf

			EndSwitch

			; TODAY highlight
			If $bCurrentYear And ($sMonthStr = @MON) And ($sDayStr = @MDAY) Then

				GUICtrlSetState($TodayLabel[$i][$j], $GUI_SHOW)

				If $tip <> "" Then
					GUICtrlSetTip($Inputs[$i][$j], StringReplace($tip, @CRLF, @CRLF & "- "), _
							$WeekDayName & " - " & $iYear & "/" & $sMonthStr & "/" & $sDayStr & _
							" - " & $StatusName & " - TODAY")
				Else
					GUICtrlSetTip($Inputs[$i][$j], $StatusName, _
							$WeekDayName & " - " & $iYear & "/" & $sMonthStr & "/" & $sDayStr & " - TODAY")
				EndIf

			EndIf

		Next ; day loop

		$C += 1
		If $C > 2 Then
			$C = 0
			$Skip += 10
		EndIf

	Next ; month loop

	$XCount += 1
	ConsoleWrite("$XCount: " & $XCount & @CRLF)
	_CreateMenu()
	Return

EndFunc   ;==>_ReadINI



Func _oldReadINI($Year)

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


			$SelectLabel[$i][$j] = GUICtrlCreateLabel("", -3 + ($i * 35), 202 + $Skip + ($j * 25), 36, 28)     ;Select Day label
			GUICtrlSetBkColor($SelectLabel[$i][$j], $Color_bk_Selected)
			GUICtrlSetState($SelectLabel[$i][$j], $gui_disable)
			GUICtrlSetState($SelectLabel[$i][$j], $gui_hide)

			$TodayLabel[$i][$j] = GUICtrlCreateLabel("", -1 + ($i * 35), 204 + $Skip + ($j * 25), 32, 24)     ;Today Label
			GUICtrlSetBkColor($TodayLabel[$i][$j], $Color_bk_Today)
			GUICtrlSetColor($TodayLabel[$i][$j], $Color_bk_Today)
			GUICtrlSetState($TodayLabel[$i][$j], $gui_disable)
			GUICtrlSetState($TodayLabel[$i][$j], $gui_hide)


			$Inputs[$i][$j] = GUICtrlCreateButton("", 0 + ($i * 35), 205 + $Skip + ($j * 25), 30, 22, BitOR($ES_READONLY, $ES_CENTER, $BS_FLAT, $BS_BOTTOM))

			_MenuContextual($i, $j)



			$WeekDayNum = _DateToDayOfWeek($Year, $X, $i)
			$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
			$Status1 = RegRead($DB & "\" & $Year & "\" & $X, $n)
			If @error Then
				GUICtrlDelete($Inputs[$i][$j])
			Else

				$Status = StringLeft($Status1, 1)
				If StringLen($Status1) > 1 Then
					$tip = "- " & StringTrimLeft($Status1, 1)
					GUICtrlSetFont($Inputs[$i][$j], 9, 900, 6, "", 2)
				Else
					$tip = ""
					GUICtrlSetFont($Inputs[$i][$j], 9, 100, 0, "", 2)
				EndIf

				If $Status = "W" Then $StatusName = "WEEKEND"
				If $Status = "O" Then $StatusName = "ON-SITE"
				If $Status = "R" Then $StatusName = "REMOTE"
				If $Status = "T" Then $StatusName = "TRAVEL"
				If $Status = "P" Then $StatusName = "PTO"
				If $Status = "H" Then $StatusName = "HOLIDAY"
				If $Status = "S" Then $StatusName = "SICK DAY"
				If $Status = "B" Then $StatusName = "BLANK"
				If $Status = "" Then $StatusName = "BLANK"

				If $tip <> "" Then
;~ 					GUICtrlSetTip($Inputs[$i][$j], StringReplace(StringReplace($tip,@CRLF,@CRLF & "- "), "/n", @CRLF & "- "), $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & " - " & $StatusName)
					GUICtrlSetTip($Inputs[$i][$j], StringReplace($tip, @CRLF, @CRLF & "- "), $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & " - " & $StatusName)
				Else
					GUICtrlSetTip($Inputs[$i][$j], $StatusName, $WeekDayName & " - " & $Year & "/" & $X & "/" & $n)
				EndIf


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

					GUICtrlSetState($TodayLabel[$i][$j], $gui_show)

					If $tip <> "" Then
						GUICtrlSetTip($Inputs[$i][$j], StringReplace($tip, @CRLF, @CRLF & "- "), $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & " - " & $StatusName & " - TODAY")
					Else
						GUICtrlSetTip($Inputs[$i][$j], $StatusName, $WeekDayName & " - " & $Year & "/" & $X & "/" & $n & " - TODAY")
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
	$XCount += 1
	ConsoleWrite("$XCount: " & $XCount & @CRLF)
	_CreateMenu()
	Return

EndFunc   ;==>_oldReadINI

Func _CheckQuarter()

	$SelDate = GUICtrlRead($Calendar)
	$Color_bk_Black = 0x000000

	GUICtrlSetBkColor($Input_Remaining_q1, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q2, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q3, 0xFFFFFF)
	GUICtrlSetBkColor($Input_Remaining_q4, 0xFFFFFF)

	GUICtrlSetColor($Input_Remaining_q1, 0x000000)
	GUICtrlSetColor($Input_Remaining_q2, 0x000000)
	GUICtrlSetColor($Input_Remaining_q3, 0x000000)
	GUICtrlSetColor($Input_Remaining_q4, 0x000000)

	If $Ratio_R_Q1 > 0 Or $Ratio_R_Q1 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q1, _GetColorFromValue($Remaining_q1))
		If $Ratio_R_Q1 > 2.99 Then
			GUICtrlSetBkColor($Input_Remaining_q1, 0x009900)
			GUICtrlSetColor($Input_Remaining_q1, $Color_bk_Blank)
		Else
			GUICtrlSetColor($Input_Remaining_q1, $Color_bk_Black)
		EndIf
	EndIf

	If $Ratio_R_Q2 > 0 Or $Ratio_R_Q2 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q2, _GetColorFromValue($Remaining_q2))
		If $Ratio_R_Q2 > 2.99 Then
			GUICtrlSetBkColor($Input_Remaining_q2, 0x009900)
			GUICtrlSetColor($Input_Remaining_q2, $Color_bk_Blank)
		Else
			GUICtrlSetColor($Input_Remaining_q2, $Color_bk_Black)
		EndIf
	EndIf

	If $Ratio_R_Q3 > 0 Or $Ratio_R_Q3 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q3, _GetColorFromValue($Remaining_q3))
		If $Ratio_R_Q3 > 2.99 Then
			GUICtrlSetBkColor($Input_Remaining_q3, 0x009900)
			GUICtrlSetColor($Input_Remaining_q3, $Color_bk_Blank)
		Else
			GUICtrlSetColor($Input_Remaining_q3, $Color_bk_Black)
		EndIf
	EndIf

	If $Ratio_R_Q4 > 0 Or $Ratio_R_Q4 < 0 Then
		GUICtrlSetBkColor($Input_Remaining_q4, _GetColorFromValue($Remaining_q4))
		If $Ratio_R_Q4 > 2.99 Then
			GUICtrlSetBkColor($Input_Remaining_q4, 0x009900)
			GUICtrlSetColor($Input_Remaining_q4, $Color_bk_Blank)
		Else
			GUICtrlSetColor($Input_Remaining_q4, $Color_bk_Black)
		EndIf
	EndIf


	$SelDate_slipt = StringSplit($SelDate, "/")
	If $SelDate_slipt[1] = @YEAR Then

		If @MON = "01" Or @MON = "02" Or @MON = "03" Then

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


	#cs
	For reference only:
	###### $Form_WorkDays = GUICreate("Work Days", 1140, 620, -1, -1)
	###### $Form_About = GUICreate("About", 655, 617, 280, -40, $WS_SYSMENU,$WS_EX_MDICHILD,$Form_WorkDays)
	#ce

	$Form_Colors = GUICreate('Colors', 220, 400, 300, 100, $DS_MODALFRAME, BitOR($WS_EX_TOPMOST, $WS_EX_MDICHILD), $Form_WorkDays)

;~ 	$WinPos = WinGetPos("Work Days")
;~ 	$Form_Colors = GUICreate('Colors', 220, 400, $WinPos[0] + 300, $WinPos[1] + 100, $DS_MODALFRAME, $WS_EX_TOPMOST)
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

	$debug_check = GUICtrlCreateCheckbox("Dev tools", 65, 305)
	If $Debug = 1 Then
		GUICtrlSetState($debug_check, $gui_checked)
	Else
		GUICtrlSetState($debug_check, $gui_unchecked)
	EndIf
	GUICtrlSetState($debug_check, $gui_hide)



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
			Case $debug_check
				RegWrite($DB, "Debug", "REG_SZ", GUICtrlRead($debug_check))
				$Debug = GUICtrlRead($debug_check)


			Case $Colors_Close
;~ 				Exit
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
		FileWriteLine($sFilePath_hwd, $sSubKey_settings & "=" & StringReplace($RegRead, @CRLF, " /n"))
	Next

	; Loop from 1 to 10 times, displaying registry keys at the particular instance value.
	For $i = 1 To 10000
		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

;~ 		ConsoleWrite($DB & "\" & $sSubKey & @CRLF)

		For $r = 1 To 10000
			$sSubKey_Month = RegEnumKey($DB & "\" & $sSubKey, $r)
			If @error Then ExitLoop

;~ 			ConsoleWrite($DB & "\" & $sSubKey & "\" & $sSubKey_month & @CRLF)

			For $D = 1 To 10000

				If $D < 10 Then
					$D1 = "0" & $D
				Else
					$D1 = $D
				EndIf

				$sSubKey_Day = RegEnumVal($DB & "\" & $sSubKey & "\" & $sSubKey_Month, $D1)
				If @error Then ExitLoop
				$RegRead = RegRead($DB & "\" & $sSubKey & "\" & $sSubKey_Month, $sSubKey_Day)
				FileWriteLine($sFilePath_hwd, $sSubKey & "\" & $sSubKey_Month & "\" & $sSubKey_Day & "=" & StringReplace($RegRead, @CRLF, " /n"))
			Next
		Next
	Next

	FileClose($sFilePath_hwd)

	If $DBBKP = "" Then
		MsgBox(64, "Sucess", "Backup saved: " & $sFilePath, 0, $Form_WorkDays)
	EndIf

	Return

EndFunc   ;==>_CreateBackup

Func _Interpolate($v1, $v2, $ratio)
	Return Round($v1 + ($v2 - $v1) * $ratio)
EndFunc   ;==>_Interpolate

Func _GetColorGradient($value)
	If $value = "-" Then
		Return "0x007ECD"
	Else
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
				[0.1, 255, 0, 0], _   ; Vermelho
				[1.0, 255, 128, 0], _ ; Laranja-avermelhado
				[2.0, 200, 165, 0], _ ; Laranja
				[2.5, 173, 255, 47], _ ; Amarelo-esverdeado
				[3.0, 0, 255, 0] _    ; Verde claro
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
	EndIf
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

Func _About()

	#cs
	For reference only:
	###### $Form_WorkDays = GUICreate("Work Days", 1140, 620, -1, -1)
	###### $Form_About = GUICreate("About", 655, 617, 280, -40, $WS_SYSMENU,$WS_EX_MDICHILD,$Form_WorkDays)
	#ce
;~ $AboutFile

;~ 	Global $AboutFile = @TempDir & "\about.jpg"

	FileInstall("splash.jpg", $AboutFile, 1)

;~ 	FileInstall("about.jpg", $AboutFile, 1)

;~ 	$Form_About = GUICreate("About", 655, 617, $aPos[0], $aPos[1], $WS_SYSMENU,-1,$Form_WorkDays)
	$Form_About = GUICreate("About", 655, 617, 280, -40, $WS_SYSMENU, $WS_EX_MDICHILD, $Form_WorkDays)
	$Pic_About = GUICtrlCreatePic($AboutFile, 5, 5, 640, 360)
	$About_Text = "Work Days is a user-friendly calendar-based application for managing and categorizing your workdays—On Site, Remote, and Holiday—throughout the year." & @CRLF & @CRLF & "Developed by Fabricio Zambroni - CURRENT VERSION: " & FileGetVersion(@ScriptFullPath)
	$Text_About = GUICtrlCreateEdit($About_Text, 5, 293, 640, 90, BitOR($ES_MULTILINE, $ES_READONLY), -1)
	GUICtrlSetFont($Text_About, 12)
	GUICtrlSetColor($Text_About, 0x2211FF)
	$Edit_About = GUICtrlCreateEdit($About, 5, 396, 640, 180, BitOR($ES_MULTILINE, $ES_READONLY), -1)

	GUISetState(@SW_SHOW)


	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($Form_About)
;~ 				exit
				Return

		EndSwitch
	WEnd


EndFunc   ;==>_About
