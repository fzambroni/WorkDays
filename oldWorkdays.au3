#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=xcalendar4.ico
#AutoIt3Wrapper_Res_Description=Work Day management
#AutoIt3Wrapper_Res_Fileversion=2.0.0.1
#AutoIt3Wrapper_Res_ProductName=Work Days
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\splash.jpg
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\Help.pdf
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\Updater.exe
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
#include <WinAPI.au3>
#include <FontConstants.au3>
#include <ProgressConstants.au3>
#include <GuiTab.au3>
#include <GDIPlus.au3>
#include <GUISlider.au3>
#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <StructureConstants.au3>
#include <WinAPIConstants.au3>
#include <Misc.au3>
#include <GuiMonthCal.au3>
#include <WindowsStylesConstants.au3>
#include <WinAPIGdi.au3>
#include "Workdays_Report_HTML_UTF8.au3"
#include "Workdays_HTML_TOX.au3"
#include "Workdays_Monitor_UDF.au3"
#include "Workdays_ColorChooser.au3"
#include "Workdays_ColorPicker.au3"

#Region GLOBAL
; =======================
; Config
; =======================
Global Const $g_iYear = @YEAR ; set fixed year if you want: 2026
Global $g_aMonths[12] = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

; Per-cell colors: [row 0..11][day 1..31]
Global $g_aCellColor[20][50]
Global $g_aCellColorBK[20][50]
Global $g_aCellStatus[20][50]
Global $g_aCellTip[20][50]
Global $iItem[30][30]
Global $SubItem[50][50]
Global $hImage[50]

Global $g_clrTodayBorder = 0xFF0000 ; vermelho
Global $g_clrSelectedBorder = 0x00AA00 ; verde
;~ Global $g_clrQuarterBorder = 0x000000 ; preto
;~ Global $g_iQuarterBorderSize = 2
;~ Global $g_iListViewFontHeight = 14

;Set BOLD and Underline

Global $g_hFontNormal = 0
Global $g_hFontBold = 0
Global $g_hFontUnderline = 0
Global $g_hFontBoldUnderline = 0
Global $g_hFontHeaderBold = 0

;Tip Management
Global $g_iTipRow = -1
Global $g_iTipCol = -1
Global $g_sTipText = ""
Global $g_bTipVisible = False

;Context menu
Global $mousePosX = 0
Global $mousePosY = 0

Global Const $MF_STRING = 0x00000000
Global Const $MF_GRAYED = 0x00000001
Global Const $MF_DISABLED = 0x00000002
Global Const $MF_SEPARATOR = 0x00000800

Global Const $TPM_RIGHTBUTTON = 0x0002
Global Const $TPM_NONOTIFY = 0x0080
Global Const $TPM_RETURNCMD = 0x0100

Global $g_bShowCellMenu = False
Global $g_iMenuDay = 0
Global $g_iMenuMonth = 0
Global $g_iMenuYear = 0


; Handles
Global $g_hGUI = 0, $g_hLV = 0, $g_idLV = 0

Global $iYear = @YEAR
Global $UpdatePath = "\\lp16-fzi1-dsa\WorkDays"

Global $HelpFile = @TempDir & "\Help.pdf"
Global $sSplashPath = @TempDir & "\splash.jpg"
Global $AboutFile = @TempDir & "\splash.jpg"
Global $ResetPosition = 0
Global $Progress_Splash, $Form_Splash, $Label_Percentage, $Splash, $Button_Close_Splash

;Chart Variables
Global $Total, $Count_O, $Count_R, $Count_H, $Count_P, $Count_T, $Count_S, $Count_B, $Count_W, $Percentage_O, $Degrees_O, $Percentage_R, $Degrees_R, $Percentage_H, $Degrees_H, $Percentage_P, $Degrees_P, $Percentage_T, $Degrees_T, $Percentage_S, $Degrees_S, $Percentage_B, $Degrees_B, $Percentage_W, $Degrees_W, $Chart, $Color_Graphic_Transparent = 1


;Colors Variables
Global $Color_bk_OnSite, $Color_bk_Remote, $Color_bk_holiday, $Color_bk_PTO, $Color_bk_Travel, $Color_bk_Sick, $Color_bk_Blank, $Color_bk_Weekend


Global $DB = "HKEY_CURRENT_USER\Software\WorkDays"

Global $WinPos_X = RegRead($DB, "WinPosX")
If @error Then $WinPos_X = -1
Global $WinPos_Y = RegRead($DB, "WinPosY")
If @error Then $WinPos_Y = -1

If $WinPos_X = "" Then $WinPos_X = -1
If $WinPos_Y = "" Then $WinPos_Y = -1

$cnt = _Monitor_GetCount()
ConsoleWrite("Monitor Count:" & $cnt & @CRLF)

If $cnt = 1 And $WinPos_X < 0 Then $WinPos_X = -1
If $cnt = 1 And $WinPos_Y < 0 Then $WinPos_Y = -1

ConsoleWrite("$WinPos_X:" & $WinPos_X & @CRLF)
ConsoleWrite("$WinPos_Y:" & $WinPos_Y & @CRLF)

Global $Window_X = 1140
Global $Window_Y = 620

$Left_coordinate = 0
$Top_coordinate = 0



$MonitorInfo = _Monitor_GetLayout()

$Count = 1
While 1
;~ 	ConsoleWrite($Count & " - Monitor index: " & $MonitorInfo[$Count][0] & @CRLF)
;~ 	ConsoleWrite($Count & " - Left coordinate: " & $MonitorInfo[$Count][1] & @CRLF)
;~ 	ConsoleWrite($Count & " - Top coordinate: " & $MonitorInfo[$Count][2] & @CRLF)
;~ 	ConsoleWrite($Count & " - Width: " & $MonitorInfo[$Count][3] & @CRLF)
;~ 	ConsoleWrite($Count & " - Height: " & $MonitorInfo[$Count][4] & @CRLF)
;~ 	ConsoleWrite($Count & " - IsPrimary: " & $MonitorInfo[$Count][5] & @CRLF)

	If $MonitorInfo[$Count][1] < $Left_coordinate Then
		$Left_coordinate = $MonitorInfo[$Count][1]
	EndIf

	If $MonitorInfo[$Count][2] < $Top_coordinate Then
		$Top_coordinate = $MonitorInfo[$Count][2]
	EndIf


	If $MonitorInfo[0][0] = $Count Then ExitLoop
	$Count += 1
WEnd
$Monitor_qnty = $Count


ConsoleWrite("$Monitor_qnty: " & $Monitor_qnty & @CRLF)
ConsoleWrite("$Left_coordinate: " & $Left_coordinate & @CRLF)
ConsoleWrite("$Top_coordinate: " & $Top_coordinate & @CRLF)

$Restart = ""

If Number($WinPos_X) < Number($Left_coordinate) And Number($WinPos_X) <> -1 Then
	$Restart &= "A"
EndIf

If Number($WinPos_Y) < Number($Top_coordinate) And Number($WinPos_Y) <> -1 Then
	$Restart &= "B"
EndIf

If Number($WinPos_X) > (Number($MonitorInfo[1][3]) - $Window_X) And $WinPos_X <> -1 Then
	$Restart &= "C"
EndIf

If Number($WinPos_Y) > (Number($MonitorInfo[1][4]) - $Window_Y) And $WinPos_Y <> -1 Then
	$Restart &= "D"
EndIf

ConsoleWrite("$Restart: " & $Restart & @CRLF)
If $Restart <> "" Then

	$WinPos_X = -1
	$WinPos_Y = -1

;~ 	RegWrite($DB, "WinPosX", "REG_SZ", $WinPos_X)
;~ 	RegWrite($DB, "WinPosY", "REG_SZ", $WinPos_Y)

	RegDelete($DB, "WinPosY") ;, "REG_SZ", $WinPos_Y)
	RegDelete($DB, "WinPosX") ;, "REG_SZ", $WinPos_X)

	ConsoleWrite("@ScriptFullPath: " & @ScriptFullPath & @CRLF)
;~ 	Run(@ScriptFullPath)
;~ 	Exit

EndIf


FileInstall("splash.jpg", $sSplashPath, 1)
_splash("on")

Global $About = "1.0.1.3 - Custom colors and bug fixes" & @CRLF _
		 & "1.0.1.9 - Report Functionality" & @CRLF _
		 & "1.0.2.2 - Tag Multiline" & @CRLF _
		 & "1.0.3.0 - New Contextual Menu, Splash Screen and about" & @CRLF _
		 & "1.0.4.0 - Report in PDF" & @CRLF _
		 & "1.0.4.1 - Bug fix: 'Ratio to date' metric now calculates correctly." & @CRLF _
		 & "1.0.4.2 - Bug fix: Import full database when in a different year." & @CRLF _
		 & "1.0.5.0 - Layout update." & @CRLF _
		 & "1.0.6.0 - New KPI screen and graphic" & @CRLF _
		 & "1.0.6.9 - Screen adjustments" & @CRLF _
		 & "2.0.0.1 - New layout and adjustments"

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
Global $Color_bk_Black = 0x000000

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

Global $Color_bk_Graphic = RegRead($DB, "Color_Graphic")
If @error Then $Color_bk_Graphic = 0x000000

Global $g_clrQuarterBorder = RegRead($DB, "Color_Quarter")
If @error Then $g_clrQuarterBorder = 0xE0E0E0

Global $g_iQuarterBorderSize = RegRead($DB, "Quarter_Border_Size")
If @error Then $g_iQuarterBorderSize = 2

Global $g_iListViewFontHeight = RegRead($DB, "Font_Size")
If @error Then $g_iListViewFontHeight = 14


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

Global $Picker_Font_Graphic_Read = RegRead($DB, "Font_Graphic")
If @error Then $Picker_Font_Graphic_Read = 1
Global $Font_Graphic = 1


Global $Picker_Grid_Size_X_Read = RegRead($DB, "Grid_Size_X")
If @error Then $Picker_Grid_Size_X_Read = 34

Global $Picker_Grid_Size_Y_Read = RegRead($DB, "Grid_Size_Y")
If @error Then $Picker_Grid_Size_Y_Read = 0xFF0000


$Form_WorkDays = GUICreate("Work Days", $Window_X, $Window_Y, $WinPos_X, $WinPos_Y)
If $Form_WorkDays = 0 Then Exit MsgBox(16, "Error", "Failed to create main window.")

$g_hGUI = $Form_WorkDays
If $g_hGUI = 0 Then Exit MsgBox(16, "Error", "Failed to store GUI handle.")

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
Global $BkpMenu_settings_ResetScreen = GUICtrlCreateMenuItem("Reset Screen Position", $DBpMenu_settings)
Global $BkpMenu_help = GUICtrlCreateMenu("?")
Global $BkpMenu_help_help = GUICtrlCreateMenuItem("Help", $BkpMenu_help)
Global $BkpMenu_help_space = GUICtrlCreateMenuItem("", $BkpMenu_help)
Global $BkpMenu_help_About = GUICtrlCreateMenuItem("About", $BkpMenu_help)

#EndRegion GLOBAL

$Calendar = GUICtrlCreateMonthCal(@YEAR & "/" & @MON & "/" & @MDAY, 8, 8, 273, 201, $MCS_WEEKNUMBERS)

$Group_Buttons = GUICtrlCreateGroup("", 288, 2, 270, 208)

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

$Button_Update = GUICtrlCreateButton("UPDATE AVAILABLE - Click to execute", 296, 50, 245, 30, $SS_CENTER)
GUICtrlSetColor($Button_Update, 0xFF0000)
GUICtrlSetFont($Button_Update, 8, 700)
GUICtrlSetState($Button_Update, $GUI_HIDE)


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

GUICtrlCreateLabel("Use Blank button for Weekends." & @CRLF & "Left-click on the grid for the menu.", 384, 175, 170, 30)


$SelectLabel_1 = GUICtrlCreateLabel("", 494, 87, 46, 21) ;,$SS_BLACKFRAME)
$SelectLabel_2 = GUICtrlCreateLabel("", 496, 89, 42, 17) ;,$SS_BLACKFRAME)
GUICtrlSetBkColor($SelectLabel_1, $Color_bk_Today)
GUICtrlCreateLabel("Today", 497, 90, 40, 15, $SS_CENTER)

$TodayLabel_1 = GUICtrlCreateLabel("", 494, 116, 46, 21) ;,$SS_BLACKFRAME)
$TodayLabel_2 = GUICtrlCreateLabel("", 496, 118, 42, 17) ;,$SS_BLACKFRAME)
GUICtrlSetBkColor($TodayLabel_1, $Color_bk_Selected)
GUICtrlCreateLabel("Selected", 497, 119, 40, 15, $SS_CENTER)

$Button_Reload = GUICtrlCreateButton("Reload Data", 472, 22, 75, 25)

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("", -99, -99, 1, 1)

$TabQuarters = GUICtrlCreateTab(565, 8, 295, 202)

$Group_Q1 = GUICtrlCreateTabItem(" Q1 - " & @YEAR)

Global $Group_Q1x = GUICtrlCreateGroup("", 573, 30, 277, 172)

$Label_1_q1 = GUICtrlCreateLabel("Total Days:", 576, 50, 75, 21, $SS_RIGHT)
$Label_2_q1 = GUICtrlCreateLabel("Work Days:", 576, 70, 75, 21, $SS_RIGHT)
$Label_3_q1 = GUICtrlCreateLabel("Ratio:", 576, 90, 75, 21, $SS_RIGHT)
$Label_ratio_q1 = GUICtrlCreateLabel("Ratio to Date:", 576, 110, 75, 21, $SS_RIGHT)

$Label_4_q1 = GUICtrlCreateLabel("Estim.On-Site: ", 705, 50, 65, 21, $SS_RIGHT)
$Label_5_q1 = GUICtrlCreateLabel("Real On-Site: ", 705, 70, 65, 21, $SS_RIGHT)
$Label_6_q1 = GUICtrlCreateLabel("Remaining:", 705, 90, 65, 21, $SS_RIGHT)

$Input_TD_q1 = GUICtrlCreateLabel("", 651, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q1 = GUICtrlCreateLabel("", 651, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q1 = GUICtrlCreateLabel("", 651, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q1 = GUICtrlCreateLabel("", 651, 110, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q1 = GUICtrlCreateLabel("", 770, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q1 = GUICtrlCreateLabel("", 770, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q1 = GUICtrlCreateLabel("", 770, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Label_Q1_Sumary_OnSite = GUICtrlCreateLabel("On Site: ", 587, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Label_Q1_Sumary_OnSite, $Font_OnSite)
$Label_Q1_Sumary_Value_OnSite = GUICtrlCreateLabel("XXX", 651, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_OnSite, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_OnSite, 10, 700)

$Label_Q1_Sumary_Holiday = GUICtrlCreateLabel("Holiday: ", 587, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_Holiday, $Color_bk_holiday)
GUICtrlSetColor($Label_Q1_Sumary_Holiday, $Font_Holiday)
$Label_Q1_Sumary_Value_Holiday = GUICtrlCreateLabel("XXX", 651, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_Holiday, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_Holiday, 10, 700)

$Label_Q1_Sumary_Travel = GUICtrlCreateLabel("Travel: ", 587, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_Travel, $Color_bk_Travel)
GUICtrlSetColor($Label_Q1_Sumary_Travel, $Font_Travel)
$Label_Q1_Sumary_Value_Travel = GUICtrlCreateLabel("XXX", 651, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_Travel, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_Travel, 10, 700)

$Label_Q1_Sumary_Blank = GUICtrlCreateLabel("Blank: ", 587, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_Blank, $Color_bk_Blank)
GUICtrlSetColor($Label_Q1_Sumary_Blank, $Font_Blank)
$Label_Q1_Sumary_Value_Blank = GUICtrlCreateLabel("XXX", 651, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_Blank, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_Blank, 10, 700)

$Label_Q1_Sumary_Remote = GUICtrlCreateLabel("Remote: ", 705, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_Remote, $Color_bk_Remote)
GUICtrlSetColor($Label_Q1_Sumary_Remote, $Font_Remote)
$Label_Q1_Sumary_Value_Remote = GUICtrlCreateLabel("XXX", 770, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_Remote, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_Remote, 10, 700)

$Label_Q1_Sumary_PTO = GUICtrlCreateLabel("PTO: ", 705, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_PTO, $Color_bk_PTO)
GUICtrlSetColor($Label_Q1_Sumary_PTO, $Font_PTO)
$Label_Q1_Sumary_Value_PTO = GUICtrlCreateLabel("XXX", 770, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_PTO, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_PTO, 10, 700)

$Label_Q1_Sumary_Sick = GUICtrlCreateLabel("Sick: ", 705, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_Sick, $Color_bk_Sick)
GUICtrlSetColor($Label_Q1_Sumary_Sick, $Font_Sick)
$Label_Q1_Sumary_Value_Sick = GUICtrlCreateLabel("XXX", 770, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_Sick, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_Sick, 10, 700)

$Label_Q1_Sumary_Weekend = GUICtrlCreateLabel("Weekend: ", 705, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q1_Sumary_Weekend, $Color_bk_Weekend)
GUICtrlSetColor($Label_Q1_Sumary_Weekend, $Font_Weekend)
$Label_Q1_Sumary_Value_Weekend = GUICtrlCreateLabel("XXX", 770, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q1_Sumary_Value_Weekend, $Color_bk_Blank)
GUICtrlSetFont($Label_Q1_Sumary_Value_Weekend, 10, 700)

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateTabItem("")

$Group_Q2 = GUICtrlCreateTabItem(" Q2 - " & @YEAR)

Global $Group_Q2x = GUICtrlCreateGroup("", 573, 30, 277, 172)

$Label_1_q2 = GUICtrlCreateLabel("Total Days:", 576, 50, 75, 21, $SS_RIGHT)
$Label_2_q2 = GUICtrlCreateLabel("Work Days:", 576, 70, 75, 21, $SS_RIGHT)
$Label_3_q2 = GUICtrlCreateLabel("Ratio:", 576, 90, 75, 21, $SS_RIGHT)
$Label_ratio_q2 = GUICtrlCreateLabel("Ratio to Date:", 576, 110, 75, 21, $SS_RIGHT)

$Label_4_q2 = GUICtrlCreateLabel("Estim.On-Site: ", 705, 50, 65, 21, $SS_RIGHT)
$Label_5_q2 = GUICtrlCreateLabel("Real On-Site: ", 705, 70, 65, 21, $SS_RIGHT)
$Label_6_q2 = GUICtrlCreateLabel("Remaining:", 705, 90, 65, 21, $SS_RIGHT)

$Input_TD_q2 = GUICtrlCreateLabel("", 651, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q2 = GUICtrlCreateLabel("", 651, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q2 = GUICtrlCreateLabel("", 651, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q2 = GUICtrlCreateLabel("", 651, 110, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q2 = GUICtrlCreateLabel("", 770, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q2 = GUICtrlCreateLabel("", 770, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q2 = GUICtrlCreateLabel("", 770, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Label_Q2_Sumary_OnSite = GUICtrlCreateLabel("On Site: ", 587, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Label_Q2_Sumary_OnSite, $Font_OnSite)
$Label_Q2_Sumary_Value_OnSite = GUICtrlCreateLabel("XXX", 651, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_OnSite, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_OnSite, 10, 700)

$Label_Q2_Sumary_Holiday = GUICtrlCreateLabel("Holiday: ", 587, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_Holiday, $Color_bk_holiday)
GUICtrlSetColor($Label_Q2_Sumary_Holiday, $Font_Holiday)
$Label_Q2_Sumary_Value_Holiday = GUICtrlCreateLabel("XXX", 651, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_Holiday, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_Holiday, 10, 700)

$Label_Q2_Sumary_Travel = GUICtrlCreateLabel("Travel: ", 587, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_Travel, $Color_bk_Travel)
GUICtrlSetColor($Label_Q2_Sumary_Travel, $Font_Travel)
$Label_Q2_Sumary_Value_Travel = GUICtrlCreateLabel("XXX", 651, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_Travel, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_Travel, 10, 700)

$Label_Q2_Sumary_Blank = GUICtrlCreateLabel("Blank: ", 587, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_Blank, $Color_bk_Blank)
GUICtrlSetColor($Label_Q2_Sumary_Blank, $Font_Blank)
$Label_Q2_Sumary_Value_Blank = GUICtrlCreateLabel("XXX", 651, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_Blank, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_Blank, 10, 700)

$Label_Q2_Sumary_Remote = GUICtrlCreateLabel("Remote: ", 705, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_Remote, $Color_bk_Remote)
GUICtrlSetColor($Label_Q2_Sumary_Remote, $Font_Remote)
$Label_Q2_Sumary_Value_Remote = GUICtrlCreateLabel("XXX", 770, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_Remote, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_Remote, 10, 700)

$Label_Q2_Sumary_PTO = GUICtrlCreateLabel("PTO: ", 705, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_PTO, $Color_bk_PTO)
GUICtrlSetColor($Label_Q2_Sumary_PTO, $Font_PTO)
$Label_Q2_Sumary_Value_PTO = GUICtrlCreateLabel("XXX", 770, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_PTO, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_PTO, 10, 700)

$Label_Q2_Sumary_Sick = GUICtrlCreateLabel("Sick: ", 705, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_Sick, $Color_bk_Sick)
GUICtrlSetColor($Label_Q2_Sumary_Sick, $Font_Sick)
$Label_Q2_Sumary_Value_Sick = GUICtrlCreateLabel("XXX", 770, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_Sick, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_Sick, 10, 700)

$Label_Q2_Sumary_Weekend = GUICtrlCreateLabel("Weekend: ", 705, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q2_Sumary_Weekend, $Color_bk_Weekend)
GUICtrlSetColor($Label_Q2_Sumary_Weekend, $Font_Weekend)
$Label_Q2_Sumary_Value_Weekend = GUICtrlCreateLabel("XXX", 770, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q2_Sumary_Value_Weekend, $Color_bk_Blank)
GUICtrlSetFont($Label_Q2_Sumary_Value_Weekend, 10, 700)

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateTabItem("")

$Group_Q3 = GUICtrlCreateTabItem(" Q3 - " & @YEAR)

Global $Group_Q3x = GUICtrlCreateGroup("", 573, 30, 277, 172)

$Label_1_q3 = GUICtrlCreateLabel("Total Days:", 576, 50, 75, 21, $SS_RIGHT)
$Label_2_q3 = GUICtrlCreateLabel("Work Days:", 576, 70, 75, 21, $SS_RIGHT)
$Label_3_q3 = GUICtrlCreateLabel("Ratio:", 576, 90, 75, 21, $SS_RIGHT)
$Label_ratio_q3 = GUICtrlCreateLabel("Ratio to Date:", 576, 110, 75, 21, $SS_RIGHT)

$Label_4_q3 = GUICtrlCreateLabel("Estim.On-Site: ", 705, 50, 65, 21, $SS_RIGHT)
$Label_5_q3 = GUICtrlCreateLabel("Real On-Site: ", 705, 70, 65, 21, $SS_RIGHT)
$Label_6_q3 = GUICtrlCreateLabel("Remaining:", 705, 90, 65, 21, $SS_RIGHT)

$Input_TD_q3 = GUICtrlCreateLabel("", 651, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q3 = GUICtrlCreateLabel("", 651, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q3 = GUICtrlCreateLabel("", 651, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q3 = GUICtrlCreateLabel("", 651, 110, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Input_E_Onsite_q3 = GUICtrlCreateLabel("", 770, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_R_Onsite_q3 = GUICtrlCreateLabel("", 770, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_Remaining_q3 = GUICtrlCreateLabel("", 770, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))

$Label_Q3_Sumary_OnSite = GUICtrlCreateLabel("On Site: ", 587, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Label_Q3_Sumary_OnSite, $Font_OnSite)
$Label_Q3_Sumary_Value_OnSite = GUICtrlCreateLabel("XXX", 651, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_OnSite, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_OnSite, 10, 700)

$Label_Q3_Sumary_Holiday = GUICtrlCreateLabel("Holiday: ", 587, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_Holiday, $Color_bk_holiday)
GUICtrlSetColor($Label_Q3_Sumary_Holiday, $Font_Holiday)
$Label_Q3_Sumary_Value_Holiday = GUICtrlCreateLabel("XXX", 651, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_Holiday, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_Holiday, 10, 700)

$Label_Q3_Sumary_Travel = GUICtrlCreateLabel("Travel: ", 587, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_Travel, $Color_bk_Travel)
GUICtrlSetColor($Label_Q3_Sumary_Travel, $Font_Travel)
$Label_Q3_Sumary_Value_Travel = GUICtrlCreateLabel("XXX", 651, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_Travel, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_Travel, 10, 700)

$Label_Q3_Sumary_Blank = GUICtrlCreateLabel("Blank: ", 587, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_Blank, $Color_bk_Blank)
GUICtrlSetColor($Label_Q3_Sumary_Blank, $Font_Blank)
$Label_Q3_Sumary_Value_Blank = GUICtrlCreateLabel("XXX", 651, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_Blank, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_Blank, 10, 700)

$Label_Q3_Sumary_Remote = GUICtrlCreateLabel("Remote: ", 705, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_Remote, $Color_bk_Remote)
GUICtrlSetColor($Label_Q3_Sumary_Remote, $Font_Remote)
$Label_Q3_Sumary_Value_Remote = GUICtrlCreateLabel("XXX", 770, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_Remote, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_Remote, 10, 700)

$Label_Q3_Sumary_PTO = GUICtrlCreateLabel("PTO: ", 705, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_PTO, $Color_bk_PTO)
GUICtrlSetColor($Label_Q3_Sumary_PTO, $Font_PTO)
$Label_Q3_Sumary_Value_PTO = GUICtrlCreateLabel("XXX", 770, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_PTO, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_PTO, 10, 700)

$Label_Q3_Sumary_Sick = GUICtrlCreateLabel("Sick: ", 705, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_Sick, $Color_bk_Sick)
GUICtrlSetColor($Label_Q3_Sumary_Sick, $Font_Sick)
$Label_Q3_Sumary_Value_Sick = GUICtrlCreateLabel("XXX", 770, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_Sick, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_Sick, 10, 700)

$Label_Q3_Sumary_Weekend = GUICtrlCreateLabel("Weekend: ", 705, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q3_Sumary_Weekend, $Color_bk_Weekend)
GUICtrlSetColor($Label_Q3_Sumary_Weekend, $Font_Weekend)
$Label_Q3_Sumary_Value_Weekend = GUICtrlCreateLabel("XXX", 770, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q3_Sumary_Value_Weekend, $Color_bk_Blank)
GUICtrlSetFont($Label_Q3_Sumary_Value_Weekend, 10, 700)

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateTabItem("")

$Group_Q4 = GUICtrlCreateTabItem(" Q4 - " & @YEAR)

Global $Group_Q4x = GUICtrlCreateGroup("", 573, 30, 277, 172)

;~ $Label_1_q4 = GUICtrlCreateLabel("Total Days: ", 571, 50, 79, 21, $SS_RIGHT)
$Label_1_q4 = GUICtrlCreateLabel("Total Days:", 576, 50, 75, 21, $SS_RIGHT)
$Label_2_q4 = GUICtrlCreateLabel("Work Days:", 576, 70, 75, 21, $SS_RIGHT)
$Label_3_q4 = GUICtrlCreateLabel("Ratio:", 576, 90, 75, 21, $SS_RIGHT)
$Label_ratio_q4 = GUICtrlCreateLabel("Ratio to Date:", 576, 110, 75, 21, $SS_RIGHT)

$Label_4_q4 = GUICtrlCreateLabel("Estim.On-Site:", 705, 50, 65, 21, $SS_RIGHT)
$Label_5_q4 = GUICtrlCreateLabel("Real On-Site:", 705, 70, 65, 21, $SS_RIGHT)
$Label_6_q4 = GUICtrlCreateLabel("Remaining:", 705, 90, 65, 21, $SS_RIGHT)

$Input_TD_q4 = GUICtrlCreateLabel("", 651, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_WD_q4 = GUICtrlCreateLabel("", 651, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RT_q4 = GUICtrlCreateLabel("", 651, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
$Input_RaTio_q4 = GUICtrlCreateLabel("", 651, 110, 40, 15, BitOR($ES_CENTER, $ES_READONLY))


$Input_E_Onsite_q4 = GUICtrlCreateLabel("", 770, 50, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
;~ $Input_E_Onsite_q4 = GUICtrlCreateLabel("", 770, 50, 40, 15, $SS_GRAYFRAME)
$Input_R_Onsite_q4 = GUICtrlCreateLabel("", 770, 70, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
;~ $Input_R_Onsite_q4 = GUICtrlCreateLabel("", 770, 70, 40, 15, $SS_GRAYFRAME)
$Input_Remaining_q4 = GUICtrlCreateLabel("", 770, 90, 40, 15, BitOR($ES_CENTER, $ES_READONLY))
;~ $Input_Remaining_q4 = GUICtrlCreateLabel("", 770, 90, 40, 15, $SS_GRAYFRAME)

$Label_Q4_Sumary_OnSite = GUICtrlCreateLabel("On Site: ", 587, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Label_Q4_Sumary_OnSite, $Font_OnSite)
$Label_Q4_Sumary_Value_OnSite = GUICtrlCreateLabel("XXX", 651, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_OnSite, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_OnSite, 10, 700)

$Label_Q4_Sumary_Holiday = GUICtrlCreateLabel("Holiday: ", 587, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_Holiday, $Color_bk_holiday)
GUICtrlSetColor($Label_Q4_Sumary_Holiday, $Font_Holiday)
$Label_Q4_Sumary_Value_Holiday = GUICtrlCreateLabel("XXX", 651, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_Holiday, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_Holiday, 10, 700)

$Label_Q4_Sumary_Travel = GUICtrlCreateLabel("Travel: ", 587, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_Travel, $Color_bk_Travel)
GUICtrlSetColor($Label_Q4_Sumary_Travel, $Font_Travel)
$Label_Q4_Sumary_Value_Travel = GUICtrlCreateLabel("XXX", 651, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_Travel, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_Travel, 10, 700)

$Label_Q4_Sumary_Blank = GUICtrlCreateLabel("Blank: ", 587, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_Blank, $Color_bk_Blank)
GUICtrlSetColor($Label_Q4_Sumary_Blank, $Font_Blank)
$Label_Q4_Sumary_Value_Blank = GUICtrlCreateLabel("XXX", 651, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_Blank, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_Blank, 10, 700)

$Label_Q4_Sumary_Remote = GUICtrlCreateLabel("Remote: ", 705, 130, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_Remote, $Color_bk_Remote)
GUICtrlSetColor($Label_Q4_Sumary_Remote, $Font_Remote)
$Label_Q4_Sumary_Value_Remote = GUICtrlCreateLabel("XXX", 770, 130, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_Remote, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_Remote, 10, 700)

$Label_Q4_Sumary_PTO = GUICtrlCreateLabel("PTO: ", 705, 145, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_PTO, $Color_bk_PTO)
GUICtrlSetColor($Label_Q4_Sumary_PTO, $Font_PTO)
$Label_Q4_Sumary_Value_PTO = GUICtrlCreateLabel("XXX", 770, 145, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_PTO, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_PTO, 10, 700)

$Label_Q4_Sumary_Sick = GUICtrlCreateLabel("Sick: ", 705, 160, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_Sick, $Color_bk_Sick)
GUICtrlSetColor($Label_Q4_Sumary_Sick, $Font_Sick)
$Label_Q4_Sumary_Value_Sick = GUICtrlCreateLabel("XXX", 770, 160, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_Sick, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_Sick, 10, 700)

$Label_Q4_Sumary_Weekend = GUICtrlCreateLabel("Weekend: ", 705, 175, 65, 15, $SS_RIGHT)
GUICtrlSetBkColor($Label_Q4_Sumary_Weekend, $Color_bk_Weekend)
GUICtrlSetColor($Label_Q4_Sumary_Weekend, $Font_Weekend)
$Label_Q4_Sumary_Value_Weekend = GUICtrlCreateLabel("XXX", 770, 175, 40, 15, $SS_CENTER)
;~ GUICtrlSetBkColor($Label_Q4_Sumary_Value_Weekend, $Color_bk_Blank)
GUICtrlSetFont($Label_Q4_Sumary_Value_Weekend, 10, 700)


GUICtrlCreateTabItem("")

GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlSetState($Input_RaTio_q1, $gui_hide)
GUICtrlSetState($Input_RaTio_q2, $gui_hide)
GUICtrlSetState($Input_RaTio_q3, $gui_hide)
GUICtrlSetState($Input_RaTio_q4, $gui_hide)

$Group_YSumary = GUICtrlCreateGroup("", 865, 2, 270, 208)

$Label_YSumary_width = 65

$Label_YSumary = GUICtrlCreateLabel("Year Summary:", 870, 20, 75, 15)

$Label_YSumary_OnSite = GUICtrlCreateButton("On Site: ", 870, 40, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Label_YSumary_OnSite, $Font_OnSite)
$Label_YSumary_Value_OnSite = GUICtrlCreateLabel("XXX", 935, 40, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_OnSite, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_OnSite, 10, 700)


$Label_YSumary_Remote = GUICtrlCreateButton("Remote: ", 870, 60, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_Remote, $Color_bk_Remote)
GUICtrlSetColor($Label_YSumary_Remote, $Font_Remote)
$Label_YSumary_Value_Remote = GUICtrlCreateLabel("XXX", 935, 60, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_Remote, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_Remote, 10, 700)

$Label_YSumary_Holiday = GUICtrlCreateButton("Holiday: ", 870, 80, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_Holiday, $Color_bk_holiday)
GUICtrlSetColor($Label_YSumary_Holiday, $Font_Holiday)
$Label_YSumary_Value_Holiday = GUICtrlCreateLabel("XXX", 935, 80, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_Holiday, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_Holiday, 10, 700)

$Label_YSumary_PTO = GUICtrlCreateButton("PTO: ", 870, 100, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_PTO, $Color_bk_PTO)
GUICtrlSetColor($Label_YSumary_PTO, $Font_PTO)
$Label_YSumary_Value_PTO = GUICtrlCreateLabel("XXX", 935, 100, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_PTO, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_PTO, 10, 700)

$Label_YSumary_Travel = GUICtrlCreateButton("Travel: ", 870, 120, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_Travel, $Color_bk_Travel)
GUICtrlSetColor($Label_YSumary_Travel, $Font_Travel)
$Label_YSumary_Value_Travel = GUICtrlCreateLabel("XXX", 935, 120, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_Travel, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_Travel, 10, 700)

$Label_YSumary_Sick = GUICtrlCreateButton("Sick: ", 870, 140, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_Sick, $Color_bk_Sick)
GUICtrlSetColor($Label_YSumary_Sick, $Font_Sick)
$Label_YSumary_Value_Sick = GUICtrlCreateLabel("XXX", 935, 140, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_Sick, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_Sick, 10, 700)

$Label_YSumary_Blank = GUICtrlCreateButton("Blank: ", 870, 160, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_Blank, $Color_bk_Blank)
GUICtrlSetColor($Label_YSumary_Blank, $Font_Blank)
$Label_YSumary_Value_Blank = GUICtrlCreateLabel("XXX", 935, 160, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_Blank, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_Blank, 10, 700)

$Label_YSumary_Weekend = GUICtrlCreateButton("Weekend: ", 870, 180, 65, 20)
;~ GUICtrlSetState(-1, $gui_checked)
GUICtrlSetBkColor($Label_YSumary_Weekend, $Color_bk_Weekend)
GUICtrlSetColor($Label_YSumary_Weekend, $Font_Weekend)
$Label_YSumary_Value_Weekend = GUICtrlCreateLabel("XXX", 935, 180, $Label_YSumary_width, 20, BitOR($SS_CENTER, $BS_PUSHLIKE))
;~ GUICtrlSetBkColor($Label_YSumary_Value_Weekend, $Color_bk_Blank)
GUICtrlSetFont($Label_YSumary_Value_Weekend, 10, 700)

$Label_YSumary_Reset = GUICtrlCreateButton("Reset", 1030, 180, 65, 20)

Global $Pie1_left = 945
Global $Pie1_top = 40
Global $Pie1_width = 240
Global $Pie1_height = 140

$Pie1 = GUICtrlCreateGraphic($Pie1_left, $Pie1_top, $Pie1_width, $Pie1_height) ;Create the main graphic area

GUICtrlCreateGroup("", -99, -99, 1, 1)

$sMessage1 = "Developed by Fabricio Zambroni - VERSION: " & FileGetVersion(@ScriptFullPath) & " - Today: " & @YEAR & "/" & @MON & "/" & @MDAY
;~ $sMessage2 = "Update Available - New Version: xxx"
$StatusBar1 = _GUICtrlStatusBar_Create($Form_WorkDays)
Dim $StatusBar1_PartsWidth[2] = [750, -1]
_GUICtrlStatusBar_SetParts($StatusBar1, $StatusBar1_PartsWidth)
_GUICtrlStatusBar_SetText($StatusBar1, $sMessage1, 0)

GUICtrlSetData($Progress_Splash, 30)
GUICtrlSetData($Label_Percentage, "30%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
_CriaINI(@YEAR)
GUICtrlSetData($Progress_Splash, 40)
GUICtrlSetData($Label_Percentage, "40%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
_DBRepair()
GUICtrlSetData($Progress_Splash, 50)
GUICtrlSetData($Label_Percentage, "50%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
_ReadINI(@YEAR, 1)
GUICtrlSetData($Progress_Splash, 80)
GUICtrlSetData($Label_Percentage, "80%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
_CheckQuarter()
GUICtrlSetData($Progress_Splash, 90)
GUICtrlSetData($Label_Percentage, "90%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
;~ Sleep(100)
_AutoBKP()
GUICtrlSetData($Progress_Splash, 95)
GUICtrlSetData($Label_Percentage, "95%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
_CreateMenu()
;~ Sleep(100)
GUICtrlSetData($Progress_Splash, 100)
GUICtrlSetData($Label_Percentage, "100%")
If GUICtrlRead($Button_Close_Splash) = $GUI_CHECKED Then
	Exit
EndIf
$SelDate = GUICtrlRead($Calendar)
$SelDate_slipt = StringSplit($SelDate, "/")

$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
$Status = StringTrimLeft($Status1, 1)

GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

$UpdatedVersion = FileGetVersion($UpdatePath & "\WorkDays.exe")

GUISetState(@SW_SHOW, $Form_WorkDays)

ConsoleWrite("Window is visible: " & _Monitor_IsVisibleWindow($Form_WorkDays) & @CRLF)

GUIDelete($Form_Splash)
FileDelete($sSplashPath)

$currentVersion = FileGetVersion(@ScriptDir & "\WorkDays.exe")

ConsoleWrite("$UpdatedVersion:" & $UpdatedVersion & @CRLF)
ConsoleWrite("$currentVersion:" & $currentVersion & @CRLF)

If $UpdatedVersion > $currentVersion Then
	FileCopy($UpdatePath & "\WorkDays.exe", @ScriptDir & "\WorkDays.tmp", 9)
EndIf
If FileExists(@ScriptDir & "\WorkDays.tmp") Then
	GUICtrlSetState($Button_Update, $GUI_SHOW)
Else
	GUICtrlSetState($Button_Update, $GUI_HIDE)
EndIf


ConsoleWrite("$Color_bk_Weekend :" & $Color_bk_Weekend & @CRLF)
ConsoleWrite("$Color_bk_OnSite :" & $Color_bk_OnSite & @CRLF)
ConsoleWrite("$Color_bk_Remote :" & $Color_bk_Remote & @CRLF)
ConsoleWrite("$Color_bk_Travel :" & $Color_bk_Travel & @CRLF)
ConsoleWrite("$Color_bk_PTO :" & $Color_bk_PTO & @CRLF)
ConsoleWrite("$Color_bk_holiday :" & $Color_bk_holiday & @CRLF)
ConsoleWrite("$Color_bk_Sick :" & $Color_bk_Sick & @CRLF)
ConsoleWrite("$Color_bk_Blank :" & $Color_bk_Blank & @CRLF)
;~ GUISetBkColor($Color_bk_Remote)
While 1



	$nMsg = GUIGetMsg()

	#cs
	If $g_bShowCellMenu Then
		$g_bShowCellMenu = False
		_MenuContextual($g_iMenuDay, $g_iMenuMonth, $g_iMenuYear)
	EndIf
	#ce

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
						_CheckQuarter()
;~ 						_ReadINI(@YEAR)
						GUICtrlSetData($Calendar, @YEAR & "/" & @MON & "/" & @MDAY)
						GUICtrlSetData($Input_SelDate, $SelDate)
						_Reload()
						MsgBox(262208, "Delete Year", "Year Deleted with Success", 0, $Form_WorkDays)

					Else
						_Reload()
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
				_Reload()
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

	_UpdateListViewCellTip()

	Switch $nMsg
		Case $Button_Update
			$Updater_File = @TempDir & "\Updater.exe"
			FileInstall("Updater.exe", $Updater_File, 1)
			Sleep(500)
			Run(@TempDir & "\Updater.exe '" & @ScriptDir & "'")
;~ 			Run($Updater_File)
			Sleep(500)
			_HideListViewCellTip()
			Exit

		Case $Label_YSumary_Reset
			_Chart("")

		Case $Label_YSumary_OnSite
			_Chart("O")

		Case $Label_YSumary_Remote
			_Chart("R")

		Case $Label_YSumary_Holiday
			_Chart("H")

		Case $Label_YSumary_PTO
			_Chart("P")

		Case $Label_YSumary_Travel
			_Chart("T")

		Case $Label_YSumary_Sick
			_Chart("S")

		Case $Label_YSumary_Blank
			_Chart("B")

		Case $Label_YSumary_Weekend
			_Chart("W")

		Case $GUI_EVENT_SECONDARYDOWN
;~ 			ConsoleWrite("********************************" & @CRLF)
			Global $mousePosX = MouseGetPos(0)
			Global $mousePosY = MouseGetPos(1)
;~ 			ConsoleWrite("$mousePosX :" & $mousePosX & @CRLF)
;~ 			ConsoleWrite("$mousePosY :" & $mousePosY & @CRLF)
;~ 			ConsoleWrite("********************************" & @CRLF)


		Case $BkpMenu_Exit
			If $ResetPosition = 0 Then
				$winPos = WinGetPos("Work Days")
				RegWrite($DB, "WinPosX", "REG_SZ", $winPos[0])
				RegWrite($DB, "WinPosY", "REG_SZ", $winPos[1])
			EndIf
			_HideListViewCellTip()
			Exit

		Case $BkpMenu_settings_ResetScreen
			RegWrite($DB, "WinPosX", "REG_SZ", "")
			RegWrite($DB, "WinPosY", "REG_SZ", "")
			$ResetPosition = 1
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262208, "Workdays", "Window position restored to the default value.", 30)
			Select
				Case $iMsgBoxAnswer = -1 ;Timeout

				Case Else                ;OK

			EndSelect


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

				GUICtrlSetBkColor($Label_YSumary_OnSite, $Color_bk_OnSite)
				GUICtrlSetColor($Label_YSumary_OnSite, $Font_OnSite)
				GUICtrlSetBkColor($Label_YSumary_Remote, $Color_bk_Remote)
				GUICtrlSetColor($Label_YSumary_Remote, $Font_Remote)
				GUICtrlSetBkColor($Label_YSumary_Holiday, $Color_bk_holiday)
				GUICtrlSetColor($Label_YSumary_Holiday, $Font_Holiday)
				GUICtrlSetBkColor($Label_YSumary_PTO, $Color_bk_PTO)
				GUICtrlSetColor($Label_YSumary_PTO, $Font_PTO)
				GUICtrlSetBkColor($Label_YSumary_Travel, $Color_bk_Travel)
				GUICtrlSetColor($Label_YSumary_Travel, $Font_Travel)
				GUICtrlSetBkColor($Label_YSumary_Sick, $Color_bk_Sick)
				GUICtrlSetColor($Label_YSumary_Sick, $Font_Sick)
				GUICtrlSetBkColor($Label_YSumary_Blank, $Color_bk_Blank)
				GUICtrlSetColor($Label_YSumary_Blank, $Font_Blank)
				GUICtrlSetBkColor($Label_YSumary_Weekend, $Color_bk_Weekend)
				GUICtrlSetColor($Label_YSumary_Weekend, $Font_Weekend)

				GUICtrlSetBkColor($Label_Q1_Sumary_OnSite, $Color_bk_OnSite)
				GUICtrlSetColor($Label_Q1_Sumary_OnSite, $Font_OnSite)
				GUICtrlSetBkColor($Label_Q1_Sumary_Holiday, $Color_bk_holiday)
				GUICtrlSetColor($Label_Q1_Sumary_Holiday, $Font_Holiday)
				GUICtrlSetBkColor($Label_Q1_Sumary_Travel, $Color_bk_Travel)
				GUICtrlSetColor($Label_Q1_Sumary_Travel, $Font_Travel)
				GUICtrlSetBkColor($Label_Q1_Sumary_Blank, $Color_bk_Blank)
				GUICtrlSetColor($Label_Q1_Sumary_Blank, $Font_Blank)
				GUICtrlSetBkColor($Label_Q1_Sumary_Remote, $Color_bk_Remote)
				GUICtrlSetColor($Label_Q1_Sumary_Remote, $Font_Remote)
				GUICtrlSetBkColor($Label_Q1_Sumary_PTO, $Color_bk_PTO)
				GUICtrlSetColor($Label_Q1_Sumary_PTO, $Font_PTO)
				GUICtrlSetBkColor($Label_Q1_Sumary_Sick, $Color_bk_Sick)
				GUICtrlSetColor($Label_Q1_Sumary_Sick, $Font_Sick)
				GUICtrlSetBkColor($Label_Q1_Sumary_Weekend, $Color_bk_Weekend)
				GUICtrlSetColor($Label_Q1_Sumary_Weekend, $Font_Weekend)

				GUICtrlSetBkColor($Label_Q2_Sumary_OnSite, $Color_bk_OnSite)
				GUICtrlSetColor($Label_Q2_Sumary_OnSite, $Font_OnSite)
				GUICtrlSetBkColor($Label_Q2_Sumary_Holiday, $Color_bk_holiday)
				GUICtrlSetColor($Label_Q2_Sumary_Holiday, $Font_Holiday)
				GUICtrlSetBkColor($Label_Q2_Sumary_Travel, $Color_bk_Travel)
				GUICtrlSetColor($Label_Q2_Sumary_Travel, $Font_Travel)
				GUICtrlSetBkColor($Label_Q2_Sumary_Blank, $Color_bk_Blank)
				GUICtrlSetColor($Label_Q2_Sumary_Blank, $Font_Blank)
				GUICtrlSetBkColor($Label_Q2_Sumary_Remote, $Color_bk_Remote)
				GUICtrlSetColor($Label_Q2_Sumary_Remote, $Font_Remote)
				GUICtrlSetBkColor($Label_Q2_Sumary_PTO, $Color_bk_PTO)
				GUICtrlSetColor($Label_Q2_Sumary_PTO, $Font_PTO)
				GUICtrlSetBkColor($Label_Q2_Sumary_Sick, $Color_bk_Sick)
				GUICtrlSetColor($Label_Q2_Sumary_Sick, $Font_Sick)
				GUICtrlSetBkColor($Label_Q2_Sumary_Weekend, $Color_bk_Weekend)
				GUICtrlSetColor($Label_Q2_Sumary_Weekend, $Font_Weekend)

				GUICtrlSetBkColor($Label_Q3_Sumary_OnSite, $Color_bk_OnSite)
				GUICtrlSetColor($Label_Q3_Sumary_OnSite, $Font_OnSite)
				GUICtrlSetBkColor($Label_Q3_Sumary_Holiday, $Color_bk_holiday)
				GUICtrlSetColor($Label_Q3_Sumary_Holiday, $Font_Holiday)
				GUICtrlSetBkColor($Label_Q3_Sumary_Travel, $Color_bk_Travel)
				GUICtrlSetColor($Label_Q3_Sumary_Travel, $Font_Travel)
				GUICtrlSetBkColor($Label_Q3_Sumary_Blank, $Color_bk_Blank)
				GUICtrlSetColor($Label_Q3_Sumary_Blank, $Font_Blank)
				GUICtrlSetBkColor($Label_Q3_Sumary_Remote, $Color_bk_Remote)
				GUICtrlSetColor($Label_Q3_Sumary_Remote, $Font_Remote)
				GUICtrlSetBkColor($Label_Q3_Sumary_PTO, $Color_bk_PTO)
				GUICtrlSetColor($Label_Q3_Sumary_PTO, $Font_PTO)
				GUICtrlSetBkColor($Label_Q3_Sumary_Sick, $Color_bk_Sick)
				GUICtrlSetColor($Label_Q3_Sumary_Sick, $Font_Sick)
				GUICtrlSetBkColor($Label_Q3_Sumary_Weekend, $Color_bk_Weekend)
				GUICtrlSetColor($Label_Q3_Sumary_Weekend, $Font_Weekend)

				GUICtrlSetBkColor($Label_Q4_Sumary_OnSite, $Color_bk_OnSite)
				GUICtrlSetColor($Label_Q4_Sumary_OnSite, $Font_OnSite)
				GUICtrlSetBkColor($Label_Q4_Sumary_Holiday, $Color_bk_holiday)
				GUICtrlSetColor($Label_Q4_Sumary_Holiday, $Font_Holiday)
				GUICtrlSetBkColor($Label_Q4_Sumary_Travel, $Color_bk_Travel)
				GUICtrlSetColor($Label_Q4_Sumary_Travel, $Font_Travel)
				GUICtrlSetBkColor($Label_Q4_Sumary_Blank, $Color_bk_Blank)
				GUICtrlSetColor($Label_Q4_Sumary_Blank, $Font_Blank)
				GUICtrlSetBkColor($Label_Q4_Sumary_Remote, $Color_bk_Remote)
				GUICtrlSetColor($Label_Q4_Sumary_Remote, $Font_Remote)
				GUICtrlSetBkColor($Label_Q4_Sumary_PTO, $Color_bk_PTO)
				GUICtrlSetColor($Label_Q4_Sumary_PTO, $Font_PTO)
				GUICtrlSetBkColor($Label_Q4_Sumary_Sick, $Color_bk_Sick)
				GUICtrlSetColor($Label_Q4_Sumary_Sick, $Font_Sick)
				GUICtrlSetBkColor($Label_Q4_Sumary_Weekend, $Color_bk_Weekend)
				GUICtrlSetColor($Label_Q4_Sumary_Weekend, $Font_Weekend)

				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")

				GUISetState(@SW_SHOW, $Form_WorkDays)
				WinSetState("Work Days", "", @SW_SHOW)

;~ 				_ReadINI($SelDate_slipt[1])
				_Reload()

				$SelDate = GUICtrlRead($Calendar)
				$SelDate_slipt = StringSplit($SelDate, "/")
				$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
				$Status = StringTrimLeft($Status1, 1)
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
			If $ResetPosition = 0 Then
				$winPos = WinGetPos("Work Days")
				RegWrite($DB, "WinPosX", "REG_SZ", $winPos[0])
				RegWrite($DB, "WinPosY", "REG_SZ", $winPos[1])
			EndIf
			_HideListViewCellTip()
			If $g_hFontHeaderBold <> 0 Then _WinAPI_DeleteObject($g_hFontHeaderBold)
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
					_Chart()

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		Case $DBpMenu_backup_Data_Holidays
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Holidays Import", "**WARNING** Importing data will overwrite any existing records for the selected dates. Do you want to proceed?", 0, $Form_WorkDays)
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					_ImportHolidays()
					_Reload()
					_Chart()


				Case $iMsgBoxAnswer = 7 ;No

			EndSelect


		Case $Calendar
			_CalendarRead()
;~ 			_Chart()

		Case $DBpMenu_backup
			_CreateBackup()

		Case $BkpMenu_reset_all
			_ResetDatabase()
			$SelDate = GUICtrlRead($Calendar)
			$SelDate_slipt = StringSplit($SelDate, "/")
			_CriaINI(@YEAR)
			_Reload()
			_Chart()

		Case $Button_Reload
			_Reload()
			_Chart()

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
;~ 					ConsoleWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month & "\" & $sSubKey_Day & "\" & $sSubKey_Day_Value & @CRLF)
				EndIf
;~ 				If StringInStr($sSubKey_Day_Value," /n ") Then
;~ 					RegWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month,$sSubKey_Day,"REG_SZ",StringReplace($sSubKey_Day_Value," /n ",@CRLF))
;~ 					ConsoleWrite($DB & "\" & $sSubKey_Year & "\" & $sSubKey_Month & "\" & $sSubKey_Day & "\" & $sSubKey_Day_Value & @CRLF)
;~ 				EndIf


			Next
		Next
	Next

EndFunc   ;==>_DBRepair

Func _Log($Inputs_log)
;~ 	$Log = FileOpen(@ScriptDir & "\Log.txt",9)
;~ 	FileWriteLine($Log,$Inputs_log)
;~ 	FileClose($Log)
	ConsoleWrite($Inputs_log & @CRLF)

EndFunc   ;==>_Log

Func _Monitor()

	_Log('+ TEST 1: Enumerate monitors -----------------------\')
	; Enumerate monitors
	_Monitor_GetList()
	Local $cnt = _Monitor_GetCount()
	If @error Then
		Local $sMsg = "TEST 1: FAILED" & @CRLF & "ERROR - Failed to enumerate monitors" & @CRLF & "@error=" & @error
		_Log("---> Example 1: ERROR - Failed to enumerate monitors")

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

	_Log('- End ------------------------------------------------/')
EndFunc   ;==>_Monitor

Func _splash($Mode = "on")

	If $Mode = "on" Then

		$splashWin_X = 640
		$splashWin_Y = 360

		If $WinPos_X = -1 And $WinPos_Y = -1 Then
			Global $Form_Splash = GUICreate("", $splashWin_X, $splashWin_Y, -1, -1, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW, $WS_EX_LAYERED))
		Else
			Global $Form_Splash = GUICreate("", $splashWin_X, $splashWin_Y, $WinPos_X + Round(($Window_X - $splashWin_X) - (($Window_X - $splashWin_X) / 2), 0), $WinPos_Y + Round(($Window_Y - $splashWin_Y) - (($Window_Y - $splashWin_Y) / 2), 0), $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW, $WS_EX_LAYERED))
		EndIf

		Global $Pic_Splash = GUICtrlCreatePic(@TempDir & "\splash.jpg", 5, 5, 630, 350)

		Global $Progress_Splash = GUICtrlCreateProgress(104, 288, 430, 17)
		Global $Label_Percentage = GUICtrlCreateLabel("0%", 540, 290, 100, -1, $SS_SIMPLE)
		GUICtrlSetColor($Label_Percentage, 0xFFFFFF)
		GUICtrlSetBkColor($Label_Percentage, 0x5b90b2)
		Global $Label_version = GUICtrlCreateLabel(FileGetVersion(@ScriptFullPath), 560, 330, -1, -1, $SS_SIMPLE)
		GUICtrlSetColor($Label_version, 0xFFFFFF)
		GUICtrlSetBkColor($Label_version, 0x5b90b2)
		Global $Button_Close_Splash = GUICtrlCreateCheckbox("X", 605, 15, 20, 20, $BS_PUSHLIKE)
		GUICtrlDelete($Button_Close_Splash)
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

	_ReadColors()

	$SelDate = GUICtrlRead($Calendar)
	$SelDate_slipt = StringSplit($SelDate, "/")


	$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	$Status = StringTrimLeft($Status1, 1)
	GUICtrlSetState($SelectLabel[$SelDate_slipt[3]][$SelDate_slipt[2]], $gui_show)

	_Chart()
	_Update($SelDate)


	Return


EndFunc   ;==>_Reload

Func _MenuContextual($U, $V, $SelYear)
	; U = dia
	; V = m�s

	Local $sDay = StringFormat("%02d", Number($U))
	Local $sMonth = StringFormat("%02d", Number($V))

	Local Const $IDM_DATE = 1001
	Local Const $IDM_TAG = 1002
	Local Const $IDM_ONSITE = 1003
	Local Const $IDM_REMOTE = 1004
	Local Const $IDM_HOLIDAY = 1005
	Local Const $IDM_PTO = 1006
	Local Const $IDM_TRAVEL = 1007
	Local Const $IDM_SICK = 1008
	Local Const $IDM_BLANK = 1009

	Local $hOwner = $Form_WorkDays
	If $hOwner = 0 Then
		ConsoleWrite("_MenuContextual: hOwner = 0" & @CRLF)
		Return SetError(100, 0, 0)
	EndIf

	Local $aMenu = DllCall("user32.dll", "handle", "CreatePopupMenu")
	If @error Or Not IsArray($aMenu) Or $aMenu[0] = 0 Then
		ConsoleWrite("_MenuContextual: CreatePopupMenu falhou" & @CRLF)
		Return SetError(1, 0, 0)
	EndIf

	Local $hMenu = $aMenu[0]
	ConsoleWrite("_MenuContextual: hMenu = " & $hMenu & @CRLF)

	If Not _AppendPopupMenu($hMenu, BitOR($MF_STRING, $MF_DISABLED, $MF_GRAYED), $IDM_DATE, "Date: " & $sDay & "/" & $sMonth & "/" & $SelYear) Then
		DllCall("user32.dll", "bool", "DestroyMenu", "handle", $hMenu)
		ConsoleWrite("Return 1" & @CRLF)
		Return SetError(2, 0, 0)
	EndIf

	If Not _AppendPopupMenu($hMenu, $MF_SEPARATOR, 0, "") Then
		DllCall("user32.dll", "bool", "DestroyMenu", "handle", $hMenu)
		ConsoleWrite("Return 2" & @CRLF)
		Return SetError(3, 0, 0)
	EndIf

	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_TAG, "Add/Edit Tag") Then Return _DestroyPopupFail($hMenu, 4)
	If Not _AppendPopupMenu($hMenu, $MF_SEPARATOR, 0, "") Then Return _DestroyPopupFail($hMenu, 5)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_ONSITE, "On-Site") Then Return _DestroyPopupFail($hMenu, 6)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_REMOTE, "Remote") Then Return _DestroyPopupFail($hMenu, 7)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_HOLIDAY, "Holiday") Then Return _DestroyPopupFail($hMenu, 8)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_PTO, "PTO") Then Return _DestroyPopupFail($hMenu, 9)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_TRAVEL, "Travel") Then Return _DestroyPopupFail($hMenu, 10)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_SICK, "Sick") Then Return _DestroyPopupFail($hMenu, 11)
	If Not _AppendPopupMenu($hMenu, $MF_STRING, $IDM_BLANK, "Blank / Weekends") Then Return _DestroyPopupFail($hMenu, 12)

;~ 	_WinAPI_SetForegroundWindow($hOwner)
	WinActivate($hOwner)

	WinActivate($hOwner)

	Local $aTrack = DllCall("user32.dll", "int", "TrackPopupMenu", _
			"handle", $hMenu, _
			"uint", BitOR($TPM_RETURNCMD, $TPM_RIGHTBUTTON), _
			"int", $mousePosX, _
			"int", $mousePosY, _
			"int", 0, _
			"hwnd", $hOwner, _
			"ptr", 0)

	DllCall("user32.dll", "lresult", "PostMessageW", _
			"hwnd", $hOwner, _
			"uint", 0, _
			"wparam", 0, _
			"lparam", 0)

	If @error Or Not IsArray($aTrack) Then
		DllCall("user32.dll", "bool", "DestroyMenu", "handle", $hMenu)
		ConsoleWrite("_MenuContextual: TrackPopupMenu falhou" & @CRLF)
		ConsoleWrite("Return 3" & @CRLF)
		Return SetError(13, 0, 0)
	EndIf

;~ 	ConsoleWrite("_MenuContextual: retorno TrackPopupMenu = " & $aTrack[0] & @CRLF)

	DllCall("user32.dll", "bool", "DestroyMenu", "handle", $hMenu)

	Switch $aTrack[0]
		Case 0
			Return 0

		Case $IDM_TAG
			_Button_Tag($sMonth, $sDay, $SelYear)
			_Reload()

		Case $IDM_ONSITE
			_Button_OnSite($sMonth, $sDay, $SelYear)

		Case $IDM_REMOTE
			_Button_Remote($sMonth, $sDay, $SelYear)

		Case $IDM_HOLIDAY
			_Button_holiday($sMonth, $sDay, $SelYear)

		Case $IDM_PTO
			_Button_PTO($sMonth, $sDay, $SelYear)

		Case $IDM_TRAVEL
			_Button_Travel($sMonth, $sDay, $SelYear)

		Case $IDM_SICK
			_Button_Sick($sMonth, $sDay, $SelYear)

		Case $IDM_BLANK
			_Button_Blank($sMonth, $sDay, $SelYear)
	EndSwitch

	Return 1
EndFunc   ;==>_MenuContextual

Func _AppendPopupMenu($hMenu, $iFlags, $iID, $sText)
	Local $aRet

	If BitAND($iFlags, $MF_SEPARATOR) = $MF_SEPARATOR Then
		$aRet = DllCall("user32.dll", "bool", "AppendMenuW", _
				"handle", $hMenu, _
				"uint", $iFlags, _
				"uint_ptr", 0, _
				"ptr", 0)
	Else
		$aRet = DllCall("user32.dll", "bool", "AppendMenuW", _
				"handle", $hMenu, _
				"uint", $iFlags, _
				"uint_ptr", $iID, _
				"wstr", $sText)
	EndIf

	If @error Then
		ConsoleWrite("_AppendPopupMenu: DllCall error = " & @error & @CRLF)
		Return False
	EndIf

	If Not IsArray($aRet) Then
		ConsoleWrite("_AppendPopupMenu: retorno invalido" & @CRLF)
		Return False
	EndIf

	If $aRet[0] = 0 Then
		ConsoleWrite("_AppendPopupMenu: AppendMenuW retornou 0" & @CRLF)
		Return False
	EndIf

	Return True
EndFunc   ;==>_AppendPopupMenu

Func _DestroyPopupFail($hMenu, $iErr)
	DllCall("user32.dll", "bool", "DestroyMenu", "handle", $hMenu)
	Return SetError($iErr, 0, 0)
EndFunc   ;==>_DestroyPopupFail

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


	If $Mouse_Tag_Pos_X_New + 200 > $Window_Tag_Pos[0] + $Window_Tag_Pos[2] Then
		$Mouse_Tag_Pos_X_New_calc = ($Mouse_Tag_Pos_X_New + 200) - ($Window_Tag_Pos[0] + $Window_Tag_Pos[2])
		$Mouse_Tag_Pos_X_New = ($Mouse_Tag_Pos_X_New - $Mouse_Tag_Pos_X_New_calc) - 70
	Else
		$Mouse_Tag_Pos_X_New = $mousePosX - 20
	EndIf


	If $Mouse_Tag_Pos_Y_New + 150 > $Window_Tag_Pos[1] + $Window_Tag_Pos[3] Then
		$Mouse_Tag_Pos_Y_New_calc = ($Mouse_Tag_Pos_Y_New + 200) - ($Window_Tag_Pos[1] + $Window_Tag_Pos[3])
		$Mouse_Tag_Pos_Y_New = ($Mouse_Tag_Pos_Y_New - $Mouse_Tag_Pos_Y_New_calc) ; - 70
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

;~ 	#cs
	Global $Form_Tag = GUICreate("Add/Edit Tag", 249, 181, $Mouse_Tag_Pos_X_New, $Mouse_Tag_Pos_Y_New, BitOR($WS_BORDER, $WS_POPUP, $DS_SETFOREGROUND, $DS_MODALFRAME), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW), $Form_WorkDays)

	$Label_Tag = GUICtrlCreateLabel("Selected Date (YYYY/MM/DD): " & $CYear & "/" & $Month & "/" & $Day, 8, 10, 250, 15)
	$Button_Tag_Cancel = GUICtrlCreateButton("Cancel", 8, 150, 75, 25)
	$Edit_Tag = GUICtrlCreateEdit("", 8, 38, 233, 105, BitOR($ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL, $ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_NOHIDESEL))
	$Button_Tag_Save = GUICtrlCreateButton("Save", 165, 150, 75, 25, $BS_DEFPUSHBUTTON)
;~ 	#ce
	#cs
		Global $Form_Tag = GUICreate("Add/Edit Tag", 249, 181, $Mouse_Tag_Pos_X_New, $Mouse_Tag_Pos_Y_New, BitOR($WS_BORDER, $WS_POPUP, $DS_SETFOREGROUND, $DS_MODALFRAME), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW), $Form_WorkDays)
		$Label_Tag = GUICtrlCreateLabel("Selected Date (YYYY/MM/DD): " & $CYear & "/" & $Month & "/" & $Day, 8, 10, 233, 105)
		$Button_Tag_Cancel = GUICtrlCreateButton("Cancel", 8, 150, 75, 25)
		$Edit_Tag = GUICtrlCreateEdit("", 8, 38, 233, 105, BitOR($ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL, $ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_NOHIDESEL))
		$Button_Tag_Save = GUICtrlCreateButton("Save", 165, 150, 75, 25);, $BS_DEFPUSHBUTTON)
	#ce
	GUICtrlSetData($Edit_Tag, $RegReadTag) ;, 1)

	GUISetState(@SW_SHOW, $Form_Tag)


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
				$SelDate_slipt = StringSplit($DateToTag, "/")
				$holidayName = GUICtrlRead($Edit_Tag)
				$Register = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
				If $Register = "" Then $Register = "B"
				RegWrite($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3], "REG_SZ", StringLeft($Register, 1) & $holidayName)
				GUIDelete($Form_Tag)
				Return

		EndSwitch
	WEnd

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
;~ 	ConsoleWrite("##### ----->>>>> Remote" & @CRLF)
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
;~ 	ConsoleWrite("##### ----->>>>> Travel" & @CRLF)
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
;~ 	ConsoleWrite("##### ----->>>>> PTO" & @CRLF)
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
;~ 	ConsoleWrite("##### ----->>>>> Holiday" & @CRLF)
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
;~ 	ConsoleWrite("##### ----->>>>> Sick" & @CRLF)
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
;~ 	ConsoleWrite("##### ----->>>>> Weekend" & @CRLF)
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

;~ 		ConsoleWrite("$i: " & $i & @CRLF)

	Next



	Return

EndFunc   ;==>_CreateMenu

Func _CheckDateReturn($DateToCheck)

	$DateToCheck_split = StringSplit($DateToCheck, "/")

	$DateToCheck_Value = RegRead($DB & "\" & $DateToCheck_split[1] & "\" & $DateToCheck_split[2], $DateToCheck_split[3])

	$DateToCheck_Value = StringLeft($DateToCheck_Value, 1)

	Return $DateToCheck_Value

EndFunc   ;==>_CheckDateReturn

Func _WorkDayInAWeekend($DateToCheck, $NewStatus)

	$DateToCheck_split = StringSplit($DateToCheck, "/")

	$WeekDayNum = _DateToDayOfWeek($DateToCheck_split[1], $DateToCheck_split[2], $DateToCheck_split[3])

	If $WeekDayNum = "1" Or $WeekDayNum = "7" Then
		$DateToCheck_Value = RegRead($DB & "\" & $DateToCheck_split[1] & "\" & $DateToCheck_split[2], $DateToCheck_split[3])
		If $NewStatus <> "W" And StringLeft($DateToCheck_Value, 1) <> $NewStatus And StringLeft($DateToCheck_Value, 1) = "W" Then
			$iMsgBoxAnswer = MsgBox(262436, "The selected day is a weekend", "The selected date is a WEEKEND!" & @CRLF & @CRLF & "Are you sure that you want to turn a WEEKEND in a WORKING DAY?", 0, $Form_WorkDays)
			Select
				Case $iMsgBoxAnswer = 6         ;Yes
;~ 					ConsoleWrite("Aqui 1" & @CRLF)
					Return 0

				Case $iMsgBoxAnswer = 7         ;No
;~ 					ConsoleWrite("Aqui 2" & @CRLF)
					Return 1
			EndSelect
		Else
			Return 0
		EndIf
	Else
		Return 0
	EndIf


EndFunc   ;==>_WorkDayInAWeekend


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

	If $DateToCheck_Value <> "" And $DateToCheck_Value <> "B" And $DateToCheck_Value <> "W" And StringLeft($DateToCheck_Value, 1) <> $NewStatus Then
		If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(262436, "Replace current value", "You're about to replace the current status for the selected date. " & @CRLF & @CRLF & "Current Status: " & _Label(StringLeft($DateToCheck_Value, 1)) & @CRLF & "New Status: " & _Label($NewStatus) & @CRLF & @CRLF & "Do you want to continue?", 0, $Form_WorkDays)
		Select
			Case $iMsgBoxAnswer = 6         ;Yes
				$WorkDayInAWeekend = _WorkDayInAWeekend($DateToCheck, $NewStatus)
				If $WorkDayInAWeekend = "0" Then ;Yes
;~ 					ConsoleWrite("Aqui 3" & @CRLF)
					Return 0
				Else ;No
;~ 					ConsoleWrite("Aqui 4" & @CRLF)
					Return 1
				EndIf

			Case $iMsgBoxAnswer = 7         ;No
;~ 				ConsoleWrite("Aqui 5" & @CRLF)
				Return 1

		EndSelect

	Else
		$WorkDayInAWeekend = _WorkDayInAWeekend($DateToCheck, $NewStatus)
		If $WorkDayInAWeekend = "0" Then ;Yes
;~ 			ConsoleWrite("Aqui 6" & @CRLF)
			Return 0
		Else ;No
;~ 			ConsoleWrite("Aqui 7" & @CRLF)
			Return 1
		EndIf
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

	GUICtrlDelete($g_idLV)

	ConsoleWrite("$SelDate: " & $SelDate & @CRLF)

	Local $aDate = StringSplit($SelDate, "/")
	If @error Or $aDate[0] <> 3 Then Return SetError(1, 0, 0)

	Local $iDataYear = Number($aDate[1])
	Local $iDataMonth = Number($aDate[2])
	Local $iDataDay = Number($aDate[3])

	_ReadINI($iDataYear)

	If $iDataMonth < 1 Or $iDataMonth > 12 Then Return SetError(2, 0, 0)
	If $iDataDay < 1 Or $iDataDay > 31 Then Return SetError(3, 0, 0)

	Local $sDataMonth = StringFormat("%02d", $iDataMonth)
	Local $sDataDay = StringFormat("%02d", $iDataDay)

	Local $sDataRegister1 = RegRead($DB & "\" & $iDataYear & "\" & $sDataMonth, $sDataDay)
	If @error Then $sDataRegister1 = ""

	Local $sDataRegister = StringLeft($sDataRegister1, 1)
	Local $sTip = ""

	If StringLen($sDataRegister1) > 1 Then
		$sTip = StringTrimLeft($sDataRegister1, 1)
		GUICtrlSetData($Input_Tag, $sTip)
	Else
		GUICtrlSetData($Input_Tag, "")
	EndIf

	Local $sStatusName = "BLANK"
	Switch $sDataRegister
		Case "W"
			$sStatusName = "WEEKEND"
		Case "O"
			$sStatusName = "ON-SITE"
		Case "R"
			$sStatusName = "REMOTE"
		Case "T"
			$sStatusName = "TRAVEL"
		Case "P"
			$sStatusName = "PTO"
		Case "H"
			$sStatusName = "HOLIDAY"
		Case "S"
			$sStatusName = "SICK DAY"
		Case "B", "", "   "
			$sStatusName = "BLANK"
	EndSwitch

	Local $iWeekDayNum = _DateToDayOfWeek($iDataYear, $iDataMonth, $iDataDay)
	Local $sWeekDayName = _DateDayOfWeek($iWeekDayNum, 1)
	Local $iWeekDayNumber = _WeekNumberISO($iDataYear, $iDataMonth, $iDataDay)

	Local $sStatusComment
	If $sTip <> "" Then
		$sStatusComment = $iDataYear & "/" & $sDataMonth & "/" & $sDataDay & @CRLF & _
				$sWeekDayName & " (Week: " & $iWeekDayNumber & ") - " & $sStatusName & @CRLF & _
				"----" & @CRLF & "- " & StringReplace($sTip, @CRLF, @CRLF & "- ")
	Else
		$sStatusComment = $iDataYear & "/" & $sDataMonth & "/" & $sDataDay & @CRLF & _
				$sWeekDayName & " (Week: " & $iWeekDayNumber & ") - " & $sStatusName
	EndIf

	; Texto mostrado na c�lula
	Local $sDisplay = $sDataRegister
	If $sDisplay = "B" Then $sDisplay = "   "

	; Atualiza o subitem do ListView
	If $g_hLV <> 0 Then
		_GUICtrlListView_SetItemText($g_hLV, $iItem[$iDataMonth][0], $sDisplay, $iDataDay)
	EndIf

	; Atualiza os arrays usados pelo NM_CUSTOMDRAW
	$g_aCellColor[$iDataMonth - 1][$iDataDay] = _ColorFromDate($sDisplay)
	$g_aCellColorBK[$iDataMonth - 1][$iDataDay] = _ColorFromDateFont($sDisplay)
	$g_aCellStatus[$iDataMonth - 1][$iDataDay] = $sTip
	$g_aCellTip[$iDataMonth - 1][$iDataDay] = $sStatusComment

	; Atualiza a data selecionada
	GUICtrlSetData($Input_SelDate, $iDataYear & "/" & $sDataMonth & "/" & $sDataDay)
	_GUICtrlMonthCal_SetCurSel($Calendar, $iDataYear, $iDataMonth, $iDataDay)
	 _CheckQuarter()

	; Recalcula estat�sticas
	_ReadStatistics($iDataYear)

	; For�a redesenho do ListView
	_WinAPI_InvalidateRect($g_hLV, 0, True)
	_WinAPI_UpdateWindow($g_hLV)

	_CreateMenu()

	Return 1

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
	$iYear = $SelDate_slipt[1]
	GUICtrlSetData($Group_Q1, " Q1 - " & $SelDate_slipt[1])
	GUICtrlSetData($Group_Q2, " Q2 - " & $SelDate_slipt[1])
	GUICtrlSetData($Group_Q3, " Q3 - " & $SelDate_slipt[1])
	GUICtrlSetData($Group_Q4, " Q4 - " & $SelDate_slipt[1])

	GUICtrlSetState($Label_ratio_q1, $gui_hide)
	GUICtrlSetState($Label_ratio_q2, $gui_hide)
	GUICtrlSetState($Label_ratio_q3, $gui_hide)
	GUICtrlSetState($Label_ratio_q4, $gui_hide)

	GUICtrlSetState($Input_RaTio_q1, $gui_hide)
	GUICtrlSetState($Input_RaTio_q2, $gui_hide)
	GUICtrlSetState($Input_RaTio_q3, $gui_hide)
	GUICtrlSetState($Input_RaTio_q4, $gui_hide)

	If $SelDate_slipt[1] <> $Input_SelDate_slipt[1] Then
		_CriaINI($SelDate_slipt[1])
;~ 		_ClearScreen()
		$SelDate_slipt = StringSplit($SelDate, "/")
;~ 		_ReadINI($SelDate_slipt[1])

	EndIf
	_Update($SelDate)
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

	Return

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

	$Count_O = 0
	$Count_R = 0
	$Count_H = 0
	$Count_P = 0
	$Count_T = 0
	$Count_S = 0
	$Count_W = 0
	$Count_B = 0

	$Count_Q1_O = 0
	$Count_Q1_R = 0
	$Count_Q1_H = 0
	$Count_Q1_P = 0
	$Count_Q1_T = 0
	$Count_Q1_S = 0
	$Count_Q1_W = 0
	$Count_Q1_B = 0

	$Count_Q2_O = 0
	$Count_Q2_R = 0
	$Count_Q2_H = 0
	$Count_Q2_P = 0
	$Count_Q2_T = 0
	$Count_Q2_S = 0
	$Count_Q2_W = 0
	$Count_Q2_B = 0

	$Count_Q3_O = 0
	$Count_Q3_R = 0
	$Count_Q3_H = 0
	$Count_Q3_P = 0
	$Count_Q3_T = 0
	$Count_Q3_S = 0
	$Count_Q3_W = 0
	$Count_Q3_B = 0

	$Count_Q4_O = 0
	$Count_Q4_R = 0
	$Count_Q4_H = 0
	$Count_Q4_P = 0
	$Count_Q4_T = 0
	$Count_Q4_S = 0
	$Count_Q4_W = 0
	$Count_Q4_B = 0

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

				If $Status = "O" Then
					$Count_O += 1
				EndIf

				If $Status = "R" Then
					$Count_R += 1
				EndIf

				If $Status = "H" Then
					$Count_H += 1
				EndIf

				If $Status = "P" Then
					$Count_P += 1
				EndIf

				If $Status = "T" Then
					$Count_T += 1
				EndIf

				If $Status = "S" Then
					$Count_S += 1
				EndIf

				If $Status = "W" Then
					$Count_W += 1
				EndIf

				If $Status = "" Or $Status = "B" Then
					$Count_B += 1
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

					If $Status = "O" Then
						$Count_Q1_O += 1
					EndIf

					If $Status = "R" Then
						$Count_Q1_R += 1
					EndIf

					If $Status = "H" Then
						$Count_Q1_H += 1
					EndIf

					If $Status = "P" Then
						$Count_Q1_P += 1
					EndIf

					If $Status = "T" Then
						$Count_Q1_T += 1
					EndIf

					If $Status = "S" Then
						$Count_Q1_S += 1
					EndIf

					If $Status = "W" Then
						$Count_Q1_W += 1
					EndIf

					If $Status = "" Or $Status = "B" Then
						$Count_Q1_B += 1
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

					If $Status = "O" Then
						$Count_Q2_O += 1
					EndIf

					If $Status = "R" Then
						$Count_Q2_R += 1
					EndIf

					If $Status = "H" Then
						$Count_Q2_H += 1
					EndIf

					If $Status = "P" Then
						$Count_Q2_P += 1
					EndIf

					If $Status = "T" Then
						$Count_Q2_T += 1
					EndIf

					If $Status = "S" Then
						$Count_Q2_S += 1
					EndIf

					If $Status = "W" Then
						$Count_Q2_W += 1
					EndIf

					If $Status = "" Or $Status = "B" Then
						$Count_Q2_B += 1
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

					If $Status = "O" Then
						$Count_Q3_O += 1
					EndIf

					If $Status = "R" Then
						$Count_Q3_R += 1
					EndIf

					If $Status = "H" Then
						$Count_Q3_H += 1
					EndIf

					If $Status = "P" Then
						$Count_Q3_P += 1
					EndIf

					If $Status = "T" Then
						$Count_Q3_T += 1
					EndIf

					If $Status = "S" Then
						$Count_Q3_S += 1
					EndIf

					If $Status = "W" Then
						$Count_Q3_W += 1
					EndIf

					If $Status = "" Or $Status = "B" Then
						$Count_Q3_B += 1
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

					If $Status = "O" Then
						$Count_Q4_O += 1
					EndIf

					If $Status = "R" Then
						$Count_Q4_R += 1
					EndIf

					If $Status = "H" Then
						$Count_Q4_H += 1
					EndIf

					If $Status = "P" Then
						$Count_Q4_P += 1
					EndIf

					If $Status = "T" Then
						$Count_Q4_T += 1
					EndIf

					If $Status = "S" Then
						$Count_Q4_S += 1
					EndIf

					If $Status = "W" Then
						$Count_Q4_W += 1
					EndIf

					If $Status = "" Or $Status = "B" Then
						$Count_Q4_B += 1
					EndIf

				EndIf


				If $Status = "O" Then
					If $j = "01" Or $j = "02" Or $j = "03" Then
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q1 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q2 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q3 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q4 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q1 += 1
;~ 						EndIf

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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q2 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q3 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q4 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q1 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q2 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q3 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q4 += 1
;~ 						EndIf
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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q1 += 1
;~ 							ConsoleWrite("$Counta_WD_q1: " & $Counta_WD_q1 & @CRLF)
;~ 						EndIf

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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q2 += 1
;~ 						EndIf

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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q3 += 1
;~ 						EndIf

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
;~ 						If $WeekDayNum <> 1 And $WeekDayNum <> 7 Then
						$Counta_WD_q4 += 1
;~ 						EndIf

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




	GUICtrlSetData($Label_Q1_Sumary_Value_OnSite, $Count_Q1_O)
	GUICtrlSetData($Label_Q1_Sumary_Value_Holiday, $Count_Q1_H)
	GUICtrlSetData($Label_Q1_Sumary_Value_Travel, $Count_Q1_T)
	GUICtrlSetData($Label_Q1_Sumary_Value_Blank, $Count_Q1_B)
	GUICtrlSetData($Label_Q1_Sumary_Value_Remote, $Count_Q1_R)
	GUICtrlSetData($Label_Q1_Sumary_Value_PTO, $Count_Q1_P)
	GUICtrlSetData($Label_Q1_Sumary_Value_Sick, $Count_Q1_S)
	GUICtrlSetData($Label_Q1_Sumary_Value_Weekend, $Count_Q1_W)

	GUICtrlSetData($Label_Q2_Sumary_Value_OnSite, $Count_Q2_O)
	GUICtrlSetData($Label_Q2_Sumary_Value_Holiday, $Count_Q2_H)
	GUICtrlSetData($Label_Q2_Sumary_Value_Travel, $Count_Q2_T)
	GUICtrlSetData($Label_Q2_Sumary_Value_Blank, $Count_Q2_B)
	GUICtrlSetData($Label_Q2_Sumary_Value_Remote, $Count_Q2_R)
	GUICtrlSetData($Label_Q2_Sumary_Value_PTO, $Count_Q2_P)
	GUICtrlSetData($Label_Q2_Sumary_Value_Sick, $Count_Q2_S)
	GUICtrlSetData($Label_Q2_Sumary_Value_Weekend, $Count_Q2_W)

	GUICtrlSetData($Label_Q3_Sumary_Value_OnSite, $Count_Q3_O)
	GUICtrlSetData($Label_Q3_Sumary_Value_Holiday, $Count_Q3_H)
	GUICtrlSetData($Label_Q3_Sumary_Value_Travel, $Count_Q3_T)
	GUICtrlSetData($Label_Q3_Sumary_Value_Blank, $Count_Q3_B)
	GUICtrlSetData($Label_Q3_Sumary_Value_Remote, $Count_Q3_R)
	GUICtrlSetData($Label_Q3_Sumary_Value_PTO, $Count_Q3_P)
	GUICtrlSetData($Label_Q3_Sumary_Value_Sick, $Count_Q3_S)
	GUICtrlSetData($Label_Q3_Sumary_Value_Weekend, $Count_Q3_W)

	GUICtrlSetData($Label_Q4_Sumary_Value_OnSite, $Count_Q4_O)
	GUICtrlSetData($Label_Q4_Sumary_Value_Holiday, $Count_Q4_H)
	GUICtrlSetData($Label_Q4_Sumary_Value_Travel, $Count_Q4_T)
	GUICtrlSetData($Label_Q4_Sumary_Value_Blank, $Count_Q4_B)
	GUICtrlSetData($Label_Q4_Sumary_Value_Remote, $Count_Q4_R)
	GUICtrlSetData($Label_Q4_Sumary_Value_PTO, $Count_Q4_P)
	GUICtrlSetData($Label_Q4_Sumary_Value_Sick, $Count_Q4_S)
	GUICtrlSetData($Label_Q4_Sumary_Value_Weekend, $Count_Q4_W)



	_Chart()

	GUICtrlSetData($Label_YSumary_Value_OnSite, $Count_O & " (" & Round($Percentage_O * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_Remote, $Count_R & " (" & Round($Percentage_R * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_Holiday, $Count_H & " (" & Round($Percentage_H * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_PTO, $Count_P & " (" & Round($Percentage_P * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_Travel, $Count_T & " (" & Round($Percentage_T * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_Sick, $Count_S & " (" & Round($Percentage_S * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_Blank, $Count_B & " (" & Round($Percentage_B * 100, 0) & "%)")
	GUICtrlSetData($Label_YSumary_Value_Weekend, $Count_W & " (" & Round($Percentage_W * 100, 0) & "%)")

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

;~ 	ConsoleWrite("$Counta_WD_q1: " & $Counta_WD_q1 & @CRLF)
;~ 	MsgBox(2602144,"","$Counta_R_Onsite_q1: "& $Counta_R_Onsite_q1 & @CRLF & "$Counta_WD_q1: " & $Counta_WD_q1)

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

	#cs
		ConsoleWrite(@CRLF & _
				"Dias �teis em Q2 (total): " & $Counta_WD_q3 & @CRLF & _
				"Dias �teis em Q2 (to date): " & $Counta_WD_Quarter_Q3 & @CRLF & _
				"Dias on-site Q2 (total): " & $Counta_R_Onsite_q3 & @CRLF & _
				"Dias on-site Q2 (to date): " & $Counta_R_Onsite_Quarter_Q3 & @CRLF & _
				"Ratio Q2 (total): " & $Ratio_R_Q3 & @CRLF & _
				"Ratio Q2 (to date): " & $Ratio_Q3 & @CRLF)
	#ce

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

Func _Chart($Type = "")



	$Color_Graphic_Transparent = RegRead($DB, "Font_Graphic")
	$Color_Graphic_BK = RegRead($DB, "Color_Graphic")

	GUICtrlSetBkColor($Label_YSumary_OnSite, $Color_bk_OnSite)
	GUICtrlSetColor($Label_YSumary_OnSite, $Font_OnSite)

	GUICtrlSetBkColor($Label_YSumary_Remote, $Color_bk_Remote)
	GUICtrlSetColor($Label_YSumary_Remote, $Font_Remote)

	GUICtrlSetBkColor($Label_YSumary_Holiday, $Color_bk_holiday)
	GUICtrlSetColor($Label_YSumary_Holiday, $Font_Holiday)

	GUICtrlSetBkColor($Label_YSumary_PTO, $Color_bk_PTO)
	GUICtrlSetColor($Label_YSumary_PTO, $Font_PTO)

	GUICtrlSetBkColor($Label_YSumary_Travel, $Color_bk_Travel)
	GUICtrlSetColor($Label_YSumary_Travel, $Font_Travel)

	GUICtrlSetBkColor($Label_YSumary_Sick, $Color_bk_Sick)
	GUICtrlSetColor($Label_YSumary_Sick, $Font_Sick)

	GUICtrlSetBkColor($Label_YSumary_Blank, $Color_bk_Blank)
	GUICtrlSetColor($Label_YSumary_Blank, $Font_Blank)

	GUICtrlSetBkColor($Label_YSumary_Weekend, $Color_bk_Weekend)
	GUICtrlSetColor($Label_YSumary_Weekend, $Font_Weekend)


	GUICtrlSetBkColor($Label_Q1_Sumary_OnSite, $Color_bk_OnSite)
	GUICtrlSetColor($Label_Q1_Sumary_OnSite, $Font_OnSite)

	GUICtrlSetBkColor($Label_Q1_Sumary_Holiday, $Color_bk_holiday)
	GUICtrlSetColor($Label_Q1_Sumary_Holiday, $Font_Holiday)

	GUICtrlSetBkColor($Label_Q1_Sumary_Travel, $Color_bk_Travel)
	GUICtrlSetColor($Label_Q1_Sumary_Travel, $Font_Travel)

	GUICtrlSetBkColor($Label_Q1_Sumary_Blank, $Color_bk_Blank)
	GUICtrlSetColor($Label_Q1_Sumary_Blank, $Font_Blank)

	GUICtrlSetBkColor($Label_Q1_Sumary_Remote, $Color_bk_Remote)
	GUICtrlSetColor($Label_Q1_Sumary_Remote, $Font_Remote)

	GUICtrlSetBkColor($Label_Q1_Sumary_PTO, $Color_bk_PTO)
	GUICtrlSetColor($Label_Q1_Sumary_PTO, $Font_PTO)

	GUICtrlSetBkColor($Label_Q1_Sumary_Sick, $Color_bk_Sick)
	GUICtrlSetColor($Label_Q1_Sumary_Sick, $Font_Sick)

	GUICtrlSetBkColor($Label_Q1_Sumary_Weekend, $Color_bk_Weekend)
	GUICtrlSetColor($Label_Q1_Sumary_Weekend, $Font_Weekend)

	GUICtrlSetBkColor($Label_Q2_Sumary_OnSite, $Color_bk_OnSite)
	GUICtrlSetColor($Label_Q2_Sumary_OnSite, $Font_OnSite)

	GUICtrlSetBkColor($Label_Q2_Sumary_Holiday, $Color_bk_holiday)
	GUICtrlSetColor($Label_Q2_Sumary_Holiday, $Font_Holiday)

	GUICtrlSetBkColor($Label_Q2_Sumary_Travel, $Color_bk_Travel)
	GUICtrlSetColor($Label_Q2_Sumary_Travel, $Font_Travel)

	GUICtrlSetBkColor($Label_Q2_Sumary_Blank, $Color_bk_Blank)
	GUICtrlSetColor($Label_Q2_Sumary_Blank, $Font_Blank)

	GUICtrlSetBkColor($Label_Q2_Sumary_Remote, $Color_bk_Remote)
	GUICtrlSetColor($Label_Q2_Sumary_Remote, $Font_Remote)

	GUICtrlSetBkColor($Label_Q2_Sumary_PTO, $Color_bk_PTO)
	GUICtrlSetColor($Label_Q2_Sumary_PTO, $Font_PTO)

	GUICtrlSetBkColor($Label_Q2_Sumary_Sick, $Color_bk_Sick)
	GUICtrlSetColor($Label_Q2_Sumary_Sick, $Font_Sick)

	GUICtrlSetBkColor($Label_Q2_Sumary_Weekend, $Color_bk_Weekend)
	GUICtrlSetColor($Label_Q2_Sumary_Weekend, $Font_Weekend)

	GUICtrlSetBkColor($Label_Q3_Sumary_OnSite, $Color_bk_OnSite)
	GUICtrlSetColor($Label_Q3_Sumary_OnSite, $Font_OnSite)

	GUICtrlSetBkColor($Label_Q3_Sumary_Holiday, $Color_bk_holiday)
	GUICtrlSetColor($Label_Q3_Sumary_Holiday, $Font_Holiday)

	GUICtrlSetBkColor($Label_Q3_Sumary_Travel, $Color_bk_Travel)
	GUICtrlSetColor($Label_Q3_Sumary_Travel, $Font_Travel)

	GUICtrlSetBkColor($Label_Q3_Sumary_Blank, $Color_bk_Blank)
	GUICtrlSetColor($Label_Q3_Sumary_Blank, $Font_Blank)

	GUICtrlSetBkColor($Label_Q3_Sumary_Remote, $Color_bk_Remote)
	GUICtrlSetColor($Label_Q3_Sumary_Remote, $Font_Remote)

	GUICtrlSetBkColor($Label_Q3_Sumary_PTO, $Color_bk_PTO)
	GUICtrlSetColor($Label_Q3_Sumary_PTO, $Font_PTO)

	GUICtrlSetBkColor($Label_Q3_Sumary_Sick, $Color_bk_Sick)
	GUICtrlSetColor($Label_Q3_Sumary_Sick, $Font_Sick)

	GUICtrlSetBkColor($Label_Q3_Sumary_Weekend, $Color_bk_Weekend)
	GUICtrlSetColor($Label_Q3_Sumary_Weekend, $Font_Weekend)

	GUICtrlSetBkColor($Label_Q4_Sumary_OnSite, $Color_bk_OnSite)
	GUICtrlSetColor($Label_Q4_Sumary_OnSite, $Font_OnSite)

	GUICtrlSetBkColor($Label_Q4_Sumary_Holiday, $Color_bk_holiday)
	GUICtrlSetColor($Label_Q4_Sumary_Holiday, $Font_Holiday)

	GUICtrlSetBkColor($Label_Q4_Sumary_Travel, $Color_bk_Travel)
	GUICtrlSetColor($Label_Q4_Sumary_Travel, $Font_Travel)

	GUICtrlSetBkColor($Label_Q4_Sumary_Blank, $Color_bk_Blank)
	GUICtrlSetColor($Label_Q4_Sumary_Blank, $Font_Blank)

	GUICtrlSetBkColor($Label_Q4_Sumary_Remote, $Color_bk_Remote)
	GUICtrlSetColor($Label_Q4_Sumary_Remote, $Font_Remote)

	GUICtrlSetBkColor($Label_Q4_Sumary_PTO, $Color_bk_PTO)
	GUICtrlSetColor($Label_Q4_Sumary_PTO, $Font_PTO)

	GUICtrlSetBkColor($Label_Q4_Sumary_Sick, $Color_bk_Sick)
	GUICtrlSetColor($Label_Q4_Sumary_Sick, $Font_Sick)

	GUICtrlSetBkColor($Label_Q4_Sumary_Weekend, $Color_bk_Weekend)
	GUICtrlSetColor($Label_Q4_Sumary_Weekend, $Font_Weekend)

	If $Type = "" Then
		$Chart = ""
	EndIf
	If StringInStr($Chart, $Type) Then
		$Chart = StringReplace($Chart, $Type, "")
	Else
		$Chart = $Chart & $Type
	EndIf
;~ 	ConsoleWrite("$Type: " & $Type & " - $Chart: " & $Chart & @CRLF)


	#Region ; PIE Chart
;~ $Pass = 16;Total number of Passed
;~ $Fail = 3;Total Number of Failed
;~ $Warnings = 5;Total Number of Warnings

	;===== The following functions calculate Percentages and Degrees =====
;~ $Total = $Count_O + $Count_R + $Count_H + $Count_P + $Count_T + $Count_S + $Count_B + $Count_W;Get the total number of all "for Percentage calculations"
	$Total = 0

;~ If $Count_O > 0 and GUICtrlRead($Label_YSumary_OnSite) = $gui_checked Then
	If $Count_O > 0 Then
		$Total = $Total + $Count_O
	EndIf

;~ If $Count_R > 0 and GUICtrlRead($Label_YSumary_Remote) = $gui_checked Then
	If $Count_R > 0 Then
		$Total = $Total + $Count_R
	EndIf

;~ If $Count_H > 0 and GUICtrlRead($Label_YSumary_Holiday) = $gui_checked Then
	If $Count_H > 0 Then
		$Total = $Total + $Count_H
	EndIf

;~ If $Count_P > 0 and GUICtrlRead($Label_YSumary_PTO) = $gui_checked Then
	If $Count_P > 0 Then
		$Total = $Total + $Count_P
	EndIf

;~ If $Count_T > 0 and GUICtrlRead($Label_YSumary_Travel) = $gui_checked Then
	If $Count_T > 0 Then
		$Total = $Total + $Count_T
	EndIf

;~ If $Count_S > 0 and GUICtrlRead($Label_YSumary_Sick) = $gui_checked Then
	If $Count_S > 0 Then
		$Total = $Total + $Count_S
	EndIf

;~ If $Count_B > 0 and GUICtrlRead($Label_YSumary_Blank) = $gui_checked Then
	If $Count_B > 0 Then
		$Total = $Total + $Count_B
	EndIf

;~ If $Count_W > 0 and GUICtrlRead($Label_YSumary_Weekend) = $gui_checked Then
	If $Count_W > 0 Then
		$Total = $Total + $Count_W
	EndIf


;~ If $Count_O > 0 and GUICtrlRead($Label_YSumary_OnSite) = $gui_checked Then
;~ 	If $Count_O > 0 Then
	$Percentage_O = $Count_O / $Total     ; Get percentage
	$Degrees_O = $Percentage_O * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_R > 0 and GUICtrlRead($Label_YSumary_Remote) = $gui_checked Then
;~ 	If $Count_R > 0 Then
	$Percentage_R = $Count_R / $Total     ; Get percentage
	$Degrees_R = $Percentage_R * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_H > 0 and GUICtrlRead($Label_YSumary_Holiday) = $gui_checked Then
;~ 	If $Count_H > 0 Then
	$Percentage_H = $Count_H / $Total     ; Get percentage
	$Degrees_H = $Percentage_H * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_P > 0 and GUICtrlRead($Label_YSumary_PTO) = $gui_checked Then
;~ 	If $Count_P > 0 Then
	$Percentage_P = $Count_P / $Total     ; Get percentage
	$Degrees_P = $Percentage_P * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_T > 0 and GUICtrlRead($Label_YSumary_Travel) = $gui_checked Then
;~ 	If $Count_T > 0 Then
	$Percentage_T = $Count_T / $Total     ; Get percentage
	$Degrees_T = $Percentage_T * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_S > 0 and GUICtrlRead($Label_YSumary_Sick) = $gui_checked Then
;~ 	If $Count_S > 0 Then
	$Percentage_S = $Count_S / $Total     ; Get percentage
	$Degrees_S = $Percentage_S * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_B > 0 and GUICtrlRead($Label_YSumary_Blank) = $gui_checked Then
;~ 	If $Count_B > 0 Then
	$Percentage_B = $Count_B / $Total     ; Get percentage
	$Degrees_B = $Percentage_B * 360     ;Get the Degrees
;~ 	EndIf

;~ If $Count_W > 0 and GUICtrlRead($Label_YSumary_Weekend) = $gui_checked Then
;~ 	If $Count_W > 0 Then
	$Percentage_W = $Count_W / $Total     ; Get percentage
	$Degrees_W = $Percentage_W * 360     ;Get the Degrees
;~ 	EndIf

	;=== This section will create the Pie Chart ==========================

	GUICtrlDelete($Pie1)
	$Pie1 = GUICtrlCreateGraphic($Pie1_left, $Pie1_top, $Pie1_width, $Pie1_height) ;Create the main graphic area

	If $Count_O = 0 And $Chart = "O" Then
		$Chart = ""
	EndIf

	If $Count_R = 0 And $Chart = "R" Then
		$Chart = ""
	EndIf

	If $Count_H = 0 And $Chart = "H" Then
		$Chart = ""
	EndIf

	If $Count_P = 0 And $Chart = "P" Then
		$Chart = ""
	EndIf

	If $Count_T = 0 And $Chart = "T" Then
		$Chart = ""
	EndIf

	If $Count_S = 0 And $Chart = "S" Then
		$Chart = ""
	EndIf

	If $Count_B = 0 And $Chart = "B" Then
		$Chart = ""
	EndIf

	If $Count_W = 0 And $Chart = "W" Then
		$Chart = ""
	EndIf

	If $Count_O > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_OnSite = $Color_bk_OnSite
		Else
			$Color_bk_Graphic_OnSite = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_OnSite, $Color_bk_OnSite) ;Set the color of Passed to light blue
		If StringInStr($Chart, "O") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90, $Degrees_O) ;Set the Pie chart piece Starts at 90^ and sweeps for $PassP number of ^
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90, $Degrees_O) ;Set the Pie chart piece Starts at 90^ and sweeps for $PassP number of ^
			EndIf
		EndIf

	EndIf


	If $Count_R > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_Remote = $Color_bk_Remote
		Else
			$Color_bk_Graphic_Remote = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_Remote, $Color_bk_Remote) ;Set the color
		If StringInStr($Chart, "R") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O, $Degrees_R) ;Set the Pie chart Piece Starts at 90^ + total ^ of $PassP
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O, $Degrees_R) ;Set the Pie chart Piece Starts at 90^ + total ^ of $PassP
			EndIf
		EndIf
	EndIf


	If $Count_H > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_holiday = $Color_bk_holiday
		Else
			$Color_bk_Graphic_holiday = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_holiday, $Color_bk_holiday)
		If StringInStr($Chart, "H") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R, $Degrees_H)
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R, $Degrees_H)
			EndIf
		EndIf
	EndIf


	If $Count_P > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_PTO = $Color_bk_PTO
		Else
			$Color_bk_Graphic_PTO = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_PTO, $Color_bk_PTO)
		If StringInStr($Chart, "P") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H, $Degrees_P)
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H, $Degrees_P)
			EndIf
		EndIf
	EndIf


	If $Count_T > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_Travel = $Color_bk_Travel
		Else
			$Color_bk_Graphic_Travel = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_Travel, $Color_bk_Travel)
		If StringInStr($Chart, "T") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P, $Degrees_T)
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P, $Degrees_T)
			EndIf
		EndIf
	EndIf

	If $Count_S > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_Sick = $Color_bk_Sick
		Else
			$Color_bk_Graphic_Sick = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_Sick, $Color_bk_Sick)
		If StringInStr($Chart, "S") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P + $Degrees_T, $Degrees_S)
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P + $Degrees_T, $Degrees_S)
			EndIf
		EndIf
	EndIf


	If $Count_B > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_Blank = $Color_bk_Blank
		Else
			$Color_bk_Graphic_Blank = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_Blank, $Color_bk_Blank)
		If StringInStr($Chart, "B") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P + $Degrees_T + $Degrees_S, $Degrees_B)
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P + $Degrees_T + $Degrees_S, $Degrees_B)
			EndIf
		EndIf
	EndIf


	If $Count_W > 0 Then

		If $Color_Graphic_Transparent = "1" Then
			$Color_bk_Graphic_Weekend = $Color_bk_Weekend
		Else
			$Color_bk_Graphic_Weekend = $Color_Graphic_BK
		EndIf

		GUICtrlSetGraphic($Pie1, $GUI_GR_COLOR, $Color_bk_Graphic_Weekend, $Color_bk_Weekend)
		If StringInStr($Chart, "W") Then
			GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P + $Degrees_T + $Degrees_S + $Degrees_B, $Degrees_W)
		Else
			If $Chart = "" Then
				GUICtrlSetGraphic($Pie1, $GUI_GR_PIE, 120, 70, 60, 90 + $Degrees_O + $Degrees_R + $Degrees_H + $Degrees_P + $Degrees_T + $Degrees_S + $Degrees_B, $Degrees_W)
			EndIf
		EndIf
	EndIf


	#EndRegion ; PIE Chart

EndFunc   ;==>_Chart

Func _ReadINI($iYear, $Splash = 0)

	ConsoleWrite("$iYear: " & $iYear & @CRLF)

	_ReadStatistics($iYear)

	$g_idLV = GUICtrlCreateListView("", 7, 210, 1127, 365, BitOR($LVS_REPORT, $LVS_SINGLESEL))
	If $g_idLV = 0 Then Exit MsgBox(16, "Error", "Failed to create ListView.")

	$g_hLV = GUICtrlGetHandle($g_idLV)
	If $g_hLV = 0 Then Exit MsgBox(16, "Error", "Failed to get ListView handle.")

	_GUICtrlListView_SetExtendedListViewStyle($g_hLV, $LVS_EX_GRIDLINES)

	; Columns
	_GUICtrlListView_InsertColumn($g_hLV, 0, "", 40, $LVCFMT_LEFT)
;~ 	_GUICtrlListView_SetTextColor ( $g_hLV,0x0000FF)
;~ 	_GUICtrlListView_SetView ($g_hLV, 3)

	For $d = 1 To 31
		_GUICtrlListView_InsertColumn($g_hLV, $d, String($d), $Picker_Grid_Size_X_Read, $LVCFMT_CENTER)
	Next

	Local $hHeader = _GUICtrlListView_GetHeader($g_hLV)
	If $hHeader <> 0 Then
		If $g_hFontHeaderBold <> 0 Then _WinAPI_DeleteObject($g_hFontHeaderBold)
		$g_hFontHeaderBold = _WinAPI_CreateFont($g_iListViewFontHeight, 0, 0, 0, 700, False, False, False, _
				$DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $DEFAULT_QUALITY, 0, "Segoe UI")
		If $g_hFontHeaderBold <> 0 Then _WinAPI_SetFont($hHeader, $g_hFontHeaderBold, True)
	EndIf


	; Rows + initial colors
	For $m = 1 To 3
		$iItem[$m][0] = _GUICtrlListView_AddItem($g_hLV, $g_aMonths[$m - 1], -1)
		_ReadDays($m, $iYear)
	Next

;~ 	Local $iItem = _GUICtrlListView_AddItem($g_hLV, "")

	For $m = 4 To 6
		$iItem[$m][0] = _GUICtrlListView_AddItem($g_hLV, $g_aMonths[$m - 1], -1)
		_ReadDays($m, $iYear)
	Next

;~ 	Local $iItem = _GUICtrlListView_AddItem($g_hLV, "")

	For $m = 7 To 9
		$iItem[$m][0] = _GUICtrlListView_AddItem($g_hLV, $g_aMonths[$m - 1], -1)
		_ReadDays($m, $iYear)
	Next

;~ 	Local $iItem = _GUICtrlListView_AddItem($g_hLV, "")

	For $m = 10 To 12
		$iItem[$m][0] = _GUICtrlListView_AddItem($g_hLV, $g_aMonths[$m - 1], -1)
		_ReadDays($m, $iYear)
	Next



	$g_hFontNormal = _WinAPI_CreateFont($g_iListViewFontHeight, 0, 0, 0, 400, False, False, False, _
			$DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $DEFAULT_QUALITY, 0, "Segoe UI")

	$g_hFontBold = _WinAPI_CreateFont($g_iListViewFontHeight, 0, 0, 0, 700, False, False, False, _
			$DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $DEFAULT_QUALITY, 0, "Segoe UI")

	$g_hFontUnderline = _WinAPI_CreateFont($g_iListViewFontHeight, 0, 0, 0, 400, False, True, False, _
			$DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $DEFAULT_QUALITY, 0, "Segoe UI")

	$g_hFontBoldUnderline = _WinAPI_CreateFont($g_iListViewFontHeight, 0, 0, 0, 700, False, True, False, _
			$DEFAULT_CHARSET, $OUT_DEFAULT_PRECIS, $CLIP_DEFAULT_PRECIS, $DEFAULT_QUALITY, 0, "Segoe UI")

	If $g_hFontNormal = 0 Or $g_hFontBold = 0 Or $g_hFontUnderline = 0 Or $g_hFontBoldUnderline = 0 Then
		Exit MsgBox(16, "Erro", "Falha ao criar as fontes.")
	EndIf

EndFunc   ;==>_ReadINI

Func _ReadDays($m, $iYear)


	For $d = 1 To 31
;~ 		ConsoleWrite("$m/$d: " & $m & "/" & $d & @CRLF)

		If $d < 10 And Not StringInStr($d, "0") Then $d = "0" & $d
		If $m < 10 And Not StringInStr($m, "0") Then $m = "0" & $m

		$Status1 = RegRead($DB & "\" & $iYear & "\" & $m, $d)

		If @error Then

			$g_aCellColor[$m - 1][$d] = _ColorFromDate("x")
			GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

		Else


			Global $Status = StringLeft($Status1, 1)


			If $Status = "W" Then $StatusName = "WEEKEND"
			If $Status = "O" Then $StatusName = "ON-SITE"
			If $Status = "R" Then $StatusName = "REMOTE"
			If $Status = "T" Then $StatusName = "TRAVEL"
			If $Status = "P" Then $StatusName = "PTO"
			If $Status = "H" Then $StatusName = "HOLIDAY"
			If $Status = "S" Then $StatusName = "SICK DAY"

			If $Status = "B" Then $StatusName = "BLANK"
			If $Status = "" Then $StatusName = "BLANK"
			If $Status = "   " Then $StatusName = "BLANK"

			$WeekDayNum = _DateToDayOfWeek($iYear, $m, $d)
			$WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
			$WeekDayNumber = _WeekNumberISO($iYear, $m, $d)


			Global $Status_Comment_1 = StringTrimLeft($Status1, 1)

			If $Status_Comment_1 <> "" Then
				$Status_Comment = $iYear & "/" & $m & "/" & $d & @CRLF & $WeekDayName & " (Week: " & $WeekDayNumber & ") - " & $StatusName & @CRLF & "----" & @CRLF & "- " & StringReplace($Status_Comment_1, @CRLF, @CRLF & "- ")
			Else
				$Status_Comment = $iYear & "/" & $m & "/" & $d & @CRLF & $WeekDayName & " (Week: " & $WeekDayNumber & ") - " & $StatusName
			EndIf


			If $Status = "B" Then $Status = "   "


			_GUICtrlListView_AddSubItem($g_hLV, $iItem[$m][0], $Status, $d, 1)




;~ 		If $m > 2 Then $m  +=1
;~ 		If $m > 5 Then $m  +=1
;~ 		If $m > 8 Then $m  +=1

			$g_aCellColor[$m - 1][$d] = _ColorFromDate($Status)
			$g_aCellColorBK[$m - 1][$d] = _ColorFromDateFont($Status)

			$g_aCellStatus[$m - 1][$d] = $Status_Comment_1
			$g_aCellTip[$m - 1][$d] = $Status_Comment

			GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

		EndIf

	Next
EndFunc   ;==>_ReadDays

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam

	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	If @error Then Return $GUI_RUNDEFMSG

	If DllStructGetData($tNMHDR, "hWndFrom") <> $g_hLV Then Return $GUI_RUNDEFMSG

	Switch DllStructGetData($tNMHDR, "Code")
;~ 		#cs
		Case $NM_CLICK

			ConsoleWrite("NM_CLICK" & @CRLF)

			Local $aHit = _GUICtrlListView_SubItemHitTest($g_hLV)
			If @error Or Not IsArray($aHit) Then Return $GUI_RUNDEFMSG

			Local $iRow = $aHit[0]
			Local $iCol = $aHit[1]
			ConsoleWrite("CLICK row/col = " & $iRow & "/" & $iCol & @CRLF)

			If $iRow < 0 Or $iCol < 1 Or $iCol > 31 Then Return $GUI_RUNDEFMSG

			Local $iMonth = $iRow + 1
			Local $iDay = $iCol

			If $iDay > _DaysInMonth2($iYear, $iMonth) Then Return $GUI_RUNDEFMSG

			$mousePosX = MouseGetPos(0)
			$mousePosY = MouseGetPos(1)

			Local $sMonth = StringFormat("%02d", $iMonth)
			Local $sDay = StringFormat("%02d", $iDay)
			Local $sDate = $iYear & "/" & $sMonth & "/" & $sDay

			GUICtrlSetData($Input_SelDate, $sDate)
			_GUICtrlMonthCal_SetCurSel($Calendar, $iYear, $iMonth, $iDay)
			_WinAPI_InvalidateRect($g_hLV, 0, True)
			_WinAPI_UpdateWindow($g_hLV)
			_CheckQuarter()

			Local $sStatus = RegRead($DB & "\" & $iYear & "\" & $sMonth, $sDay)
			If @error Or $sStatus = "" Then
				GUICtrlSetData($Input_Tag, "")
			Else
				GUICtrlSetData($Input_Tag, StringTrimLeft($sStatus, 1))
			EndIf

			$g_iMenuDay = $iDay
			$g_iMenuMonth = $iMonth
			$g_iMenuYear = $iYear
			$g_bShowCellMenu = True
			Return 0
;~ 		#ce
			#cs
				Local $aHit = _GUICtrlListView_SubItemHitTest($g_hLV)
				If @error Or Not IsArray($aHit) Then Return $GUI_RUNDEFMSG

				Local $iRow = $aHit[0]
				Local $iCol = $aHit[1] ; 0 = Month, 1..31 = day
				ConsoleWrite("$iRow/$iCol: " & $iRow & "/" & $iCol & @CRLF)

				; Only day cells
				If $iRow < 0 Or $iCol < 1 Or $iCol > 31 Then Return $GUI_RUNDEFMSG

				Local $iMonth = $iRow + 1
				Local $iDay = $iCol

				; Ignore invalid days
				If $iDay > _DaysInMonth2($g_iYear, $iMonth) Then Return $GUI_RUNDEFMSG


				; Force redraw
				_WinAPI_RedrawWindow($g_hLV, 0, 0, BitOR($RDW_INVALIDATE, $RDW_UPDATENOW))

				$FullDate = GUICtrlRead($Input_SelDate)
				$FullDate_Split = StringSplit($FullDate, "/")
				$ClickedDate = $iYear & "/" & $iMonth & "/" & $iCol
				ConsoleWrite("$ClickedDate: " & $ClickedDate & @CRLF)
				_GUICtrlMonthCal_SetCurSel($Calendar, $iYear, $iMonth, $iCol)
				_ReadINI($iYear)    ;, $Splash = 0)


				Return $GUI_RUNDEFMSG
			#ce
		Case $NM_CUSTOMDRAW
			Local $tCD = DllStructCreate($tagNMLVCUSTOMDRAW, $lParam)
			If @error Then Return $GUI_RUNDEFMSG

			Local $iStage = DllStructGetData($tCD, "dwDrawStage")

			If $iStage = $CDDS_PREPAINT Then
				Return BitOR($CDRF_NOTIFYITEMDRAW, $CDRF_NOTIFYPOSTPAINT)
			EndIf

			If $iStage = $CDDS_ITEMPREPAINT Then
				Return $CDRF_NOTIFYSUBITEMDRAW
			EndIf

			If $iStage = $CDDS_POSTPAINT Then
				_DrawQuarterSeparators()
				_DrawTodayCellBorder()
				_DrawSelectedCellBorder()
				Return $CDRF_DODEFAULT
			EndIf

			If $iStage = BitOR($CDDS_ITEMPREPAINT, $CDDS_SUBITEM) Then
				Local $iItem = DllStructGetData($tCD, "dwItemSpec")
				Local $iSub = DllStructGetData($tCD, "iSubItem")
				Local $hDC = DllStructGetData($tCD, "hdc")

				If $iSub = 0 And $iItem >= 0 And $iItem <= 11 Then
					If $hDC <> 0 Then
						_WinAPI_SelectObject($hDC, $g_hFontBold)
						Return $CDRF_NEWFONT
					EndIf
				EndIf

				If $iSub >= 1 And $iSub <= 31 And $iItem >= 0 And $iItem <= 11 Then
					Local $clr = _DecColorToRGBHex($g_aCellColor[$iItem][$iSub])
					Local $clrbk = _DecColorToRGBHex($g_aCellColorBK[$iItem][$iSub])
					Local $sSelDate = GUICtrlRead($Input_SelDate)
					Local $aSel = StringSplit($sSelDate, "/")


					Local $Status_Comment2 = $g_aCellTip[$iItem][$iSub]
					Local $Status_2 = $g_aCellStatus[$iItem][$iSub]

					If $clr <> -1 Then
						DllStructSetData($tCD, "clrTextBk", $clr)
						DllStructSetData($tCD, "clrText", $clrbk)

						If $hDC <> 0 Then
							If $Status_2 = "" Then
								_WinAPI_SelectObject($hDC, $g_hFontNormal)
								Return $CDRF_NEWFONT
							Else
								_WinAPI_SelectObject($hDC, $g_hFontBoldUnderline)
								Return $CDRF_NEWFONT
							EndIf
						EndIf
					EndIf
				EndIf

				Return $CDRF_DODEFAULT
			EndIf

			If $iStage = $CDDS_POSTPAINT Then
				_DrawTodayCellBorder()
				Return $CDRF_DODEFAULT
			EndIf

		Case $NM_RCLICK
			ConsoleWrite("NM_RCLICK" & @CRLF)

			Local $aHit = _GUICtrlListView_SubItemHitTest($g_hLV)
			If @error Or Not IsArray($aHit) Then Return $GUI_RUNDEFMSG

			Local $iRow = $aHit[0]
			Local $iCol = $aHit[1]
			ConsoleWrite("RCLICK row/col = " & $iRow & "/" & $iCol & @CRLF)

			If $iRow < 0 Or $iCol < 1 Or $iCol > 31 Then Return $GUI_RUNDEFMSG

			Local $iMonth = $iRow + 1
			Local $iDay = $iCol

			If $iDay > _DaysInMonth2($iYear, $iMonth) Then Return $GUI_RUNDEFMSG

			$mousePosX = MouseGetPos(0)
			$mousePosY = MouseGetPos(1)

			Local $sMonth = StringFormat("%02d", $iMonth)
			Local $sDay = StringFormat("%02d", $iDay)
			Local $sDate = $iYear & "/" & $sMonth & "/" & $sDay

			GUICtrlSetData($Input_SelDate, $sDate)
			_GUICtrlMonthCal_SetCurSel($Calendar, $iYear, $iMonth, $iDay)
			_WinAPI_InvalidateRect($g_hLV, 0, True)
			_WinAPI_UpdateWindow($g_hLV)
			_CheckQuarter()

			Local $sStatus = RegRead($DB & "\" & $iYear & "\" & $sMonth, $sDay)
			If @error Or $sStatus = "" Then
				GUICtrlSetData($Input_Tag, "")
			Else
				GUICtrlSetData($Input_Tag, StringTrimLeft($sStatus, 1))
			EndIf

			$g_iMenuDay = $iDay
			$g_iMenuMonth = $iMonth
			$g_iMenuYear = $iYear
			$g_bShowCellMenu = True

			$g_bShowCellMenu = False
			_MenuContextual($g_iMenuDay, $g_iMenuMonth, $g_iMenuYear)

			Return 0

	EndSwitch



	Return $GUI_RUNDEFMSG

EndFunc   ;==>WM_NOTIFY

Func _DrawQuarterSeparators()
	If $g_hLV = 0 Then Return
	If $g_iQuarterBorderSize < 1 Then Return

	Local $aFirstRect = _GUICtrlListView_GetSubItemRect($g_hLV, 0, 0, 0)
	Local $aLastRect = _GUICtrlListView_GetSubItemRect($g_hLV, 11, 31, 0)
	If @error Or Not IsArray($aFirstRect) Or Not IsArray($aLastRect) Then Return

	Local $hDC = _WinAPI_GetDC($g_hLV)
	If $hDC = 0 Then Return

	Local $hBrush = _WinAPI_CreateSolidBrush(_DecColorToRGBHex($g_clrQuarterBorder))
	If $hBrush = 0 Then
		_WinAPI_ReleaseDC($g_hLV, $hDC)
		Return
	EndIf

	For $iRow = 2 To 8 Step 3
		Local $aLeftRect = _GUICtrlListView_GetSubItemRect($g_hLV, $iRow, 0, 0)
		Local $aRightRect = _GUICtrlListView_GetSubItemRect($g_hLV, $iRow, 31, 0)
		If @error Or Not IsArray($aLeftRect) Or Not IsArray($aRightRect) Then ContinueLoop

		Local $tRect = DllStructCreate($tagRECT)
		DllStructSetData($tRect, "Left", $aLeftRect[0])
		DllStructSetData($tRect, "Top", $aLeftRect[3] - Int($g_iQuarterBorderSize / 2))
		DllStructSetData($tRect, "Right", $aRightRect[2])
		DllStructSetData($tRect, "Bottom", DllStructGetData($tRect, "Top") + $g_iQuarterBorderSize)

		_WinAPI_FillRect($hDC, $tRect, $hBrush)
	Next

	_WinAPI_DeleteObject($hBrush)
	_WinAPI_ReleaseDC($g_hLV, $hDC)
EndFunc   ;==>_DrawQuarterSeparators

Func _DrawSelectedCellBorder()
	If $g_hLV = 0 Then Return

	Local $sSelDate = GUICtrlRead($Input_SelDate)
	If $sSelDate = "" Then Return

	Local $aSel = StringSplit($sSelDate, "/")
	If @error Or $aSel[0] <> 3 Then Return

	Local $iSelYear = Number($aSel[1])
	Local $iSelMonth = Number($aSel[2])
	Local $iSelDay = Number($aSel[3])

	If $iSelYear <> $iYear Then Return
	If $iSelMonth < 1 Or $iSelMonth > 12 Then Return
	If $iSelDay < 1 Or $iSelDay > 31 Then Return
	If $iSelDay > _DaysInMonth2($iSelYear, $iSelMonth) Then Return

	Local $iRow = $iSelMonth - 1
	Local $iCol = $iSelDay

	Local $aRect = _GUICtrlListView_GetSubItemRect($g_hLV, $iRow, $iCol, 0)
	If @error Or Not IsArray($aRect) Then Return

	Local $hDC = _WinAPI_GetDC($g_hLV)
	If $hDC = 0 Then Return

	Local $tRect = DllStructCreate($tagRECT)
	DllStructSetData($tRect, "Left", $aRect[0] + 1)
	DllStructSetData($tRect, "Top", $aRect[1] + 1)
	DllStructSetData($tRect, "Right", $aRect[2] - 1)
	DllStructSetData($tRect, "Bottom", $aRect[3] - 1)

	Local $hBrush = _WinAPI_CreateSolidBrush(_DecColorToRGBHex($Color_bk_Selected))
	If $hBrush <> 0 Then
		_WinAPI_FrameRect($hDC, $tRect, $hBrush)
		_WinAPI_DeleteObject($hBrush)
	EndIf

	_WinAPI_ReleaseDC($g_hLV, $hDC)
EndFunc   ;==>_DrawSelectedCellBorder

Func _DrawTodayCellBorder()
	If $g_hLV = 0 Then Return

	; s� desenha se o ano exibido for o ano atual
	If $iYear <> @YEAR Then Return

	Local $iRow = Number(@MON) - 1
	Local $iCol = Number(@MDAY)

	If $iRow < 0 Or $iRow > 11 Then Return
	If $iCol < 1 Or $iCol > 31 Then Return
	If $iCol > _DaysInMonth2($iYear, $iRow + 1) Then Return

	Local $aRect = _GUICtrlListView_GetSubItemRect($g_hLV, $iRow, $iCol, 0)
	If @error Or Not IsArray($aRect) Then Return

	Local $hDC = _WinAPI_GetDC($g_hLV)
	If $hDC = 0 Then Return

	Local $tRect = DllStructCreate($tagRECT)
	DllStructSetData($tRect, "Left", $aRect[0] + 1)
	DllStructSetData($tRect, "Top", $aRect[1] + 1)
	DllStructSetData($tRect, "Right", $aRect[2] - 1)
	DllStructSetData($tRect, "Bottom", $aRect[3] - 1)

	Local $hBrush = _WinAPI_CreateSolidBrush(_DecColorToRGBHex($Color_bk_Today))
	If $hBrush <> 0 Then
		; 1� borda
		_WinAPI_FrameRect($hDC, $tRect, $hBrush)

		; 2� borda, 1 px para dentro
		Local $tRect2 = DllStructCreate($tagRECT)
		DllStructSetData($tRect2, "Left", DllStructGetData($tRect, "Left") + 1)
		DllStructSetData($tRect2, "Top", DllStructGetData($tRect, "Top") + 1)
		DllStructSetData($tRect2, "Right", DllStructGetData($tRect, "Right") - 1)
		DllStructSetData($tRect2, "Bottom", DllStructGetData($tRect, "Bottom") - 1)

		_WinAPI_FrameRect($hDC, $tRect2, $hBrush)

		_WinAPI_DeleteObject($hBrush)
	EndIf

	_WinAPI_ReleaseDC($g_hLV, $hDC)
EndFunc   ;==>_DrawTodayCellBorder

Func _HideListViewCellTip()
	If $g_bTipVisible Then ToolTip("")
	$g_bTipVisible = False
	$g_iTipRow = -1
	$g_iTipCol = -1
	$g_sTipText = ""
EndFunc   ;==>_HideListViewCellTip

Func _UpdateListViewCellTip()
	Local $aCur = GUIGetCursorInfo($Form_WorkDays)
	If @error Or Not IsArray($aCur) Then
		_HideListViewCellTip()
		Return
	EndIf

	; [4] = control ID sob o mouse
	If $aCur[4] <> $g_idLV Then
		_HideListViewCellTip()
		Return
	EndIf

	Local $aHit = _GUICtrlListView_SubItemHitTest($g_hLV)
	If @error Or Not IsArray($aHit) Then
		_HideListViewCellTip()
		Return
	EndIf

	Local $iRow = $aHit[0]
	Local $iCol = $aHit[1]

	; s� dias v�lidos, n�o a coluna do m�s
	If $iRow < 0 Or $iRow > 11 Or $iCol < 1 Or $iCol > 31 Then
		_HideListViewCellTip()
		Return
	EndIf

	Local $iMonth = $iRow + 1
	Local $iDay = $iCol

	If $iDay > _DaysInMonth2($iYear, $iMonth) Then
		_HideListViewCellTip()
		Return
	EndIf

	Local $sTip = $g_aCellTip[$iRow][$iCol]

	; sem coment�rio = sem tooltip
;~ 	If $sTip = "" Then
;~ 		_HideListViewCellTip()
;~ 		Return
;~ 	EndIf

	; evita recriar o tooltip toda hora
;~ 	If $g_iTipRow = $iRow And $g_iTipCol = $iCol And $g_sTipText = $sTip Then Return
	If $g_iTipRow = $iRow And $g_iTipCol = $iCol Then Return

	Local $aMouse = MouseGetPos()
	ToolTip($sTip, $aMouse[0] + 16, $aMouse[1] + 20)

	$g_iTipRow = $iRow
	$g_iTipCol = $iCol
	$g_sTipText = $sTip
	$g_bTipVisible = True
EndFunc   ;==>_UpdateListViewCellTip

Func _DecColorToRGBHex($nColor)
	Local $r = BitAND($nColor, 0xFF)
	Local $g = BitAND(BitShift($nColor, 8), 0xFF)
	Local $b = BitAND(BitShift($nColor, 16), 0xFF)
	Return StringFormat("0x%02X%02X%02X", $r, $g, $b)
EndFunc   ;==>_DecColorToRGBHex

Func _DaysInMonth2($iY, $iM)
	If $iM = 2 Then
		If _IsLeapYear($iY) Then Return 29
		Return 28
	EndIf
	Switch $iM
		Case 4, 6, 9, 11
			Return 30
		Case Else
			Return 31
	EndSwitch
EndFunc   ;==>_DaysInMonth2

Func _IsLeapYear($iY)
	If Mod($iY, 400) = 0 Then Return True
	If Mod($iY, 100) = 0 Then Return False
	Return (Mod($iY, 4) = 0)
EndFunc   ;==>_IsLeapYear

Func _ColorFromDate($Status)
	Switch $Status

		; Weekend
		Case "W"
			Return $Color_bk_Weekend

			; On-Site
		Case "O"
			Return $Color_bk_OnSite

			; Remote
		Case "R"
			Return $Color_bk_Remote

			; Travel
		Case "T"
			Return $Color_bk_Travel

			; PTO
		Case "P"
			Return $Color_bk_PTO

			; Holiday
		Case "H"
			Return $Color_bk_holiday

			; Sick Day
		Case "S"
			Return $Color_bk_Sick

			; Blank / empty
		Case "", "   ", "B", " "
			Return $Color_bk_Blank

	EndSwitch
EndFunc   ;==>_ColorFromDate

Func _ColorFromDateFont($Status)
	Switch $Status

		; Weekend
		Case "W"
			Return $Font_Weekend

			; On-Site
		Case "O"
			Return $Font_OnSite

			; Remote
		Case "R"
			Return $Font_Remote

			; Travel
		Case "T"
			Return $Font_Travel

			; PTO
		Case "P"
			Return $Font_PTO

			; Holiday
		Case "H"
			Return $Font_Holiday

			; Sick Day
		Case "S"
			Return $Font_Sick

			; Blank / empty
		Case "", "   ", "B"
			Return $Font_Blank

	EndSwitch
EndFunc   ;==>_ColorFromDateFont

Func _CheckQuarter()

	$SelDate = GUICtrlRead($Calendar)


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
;~ ConsoleWrite("$SelDate_slipt[2]: " & $SelDate_slipt[2] & @CRLF)

	If $SelDate_slipt[2] = "01" Or $SelDate_slipt[2] = "02" Or $SelDate_slipt[2] = "03" Then
		GUICtrlSetData($Input_Quarter, "Q1")
	EndIf

	If $SelDate_slipt[2] = "04" Or $SelDate_slipt[2] = "05" Or $SelDate_slipt[2] = "06" Then
		GUICtrlSetData($Input_Quarter, "Q2")
;~ 		ConsoleWrite("Aqui22" & @CRLF)
	EndIf

	If $SelDate_slipt[2] = "07" Or $SelDate_slipt[2] = "08" Or $SelDate_slipt[2] = "09" Then
		GUICtrlSetData($Input_Quarter, "Q3")
	EndIf

	If $SelDate_slipt[2] = "10" Or $SelDate_slipt[2] = "11" Or $SelDate_slipt[2] = "12" Then
		GUICtrlSetData($Input_Quarter, "Q4")
	EndIf

	If $SelDate_slipt[1] = @YEAR Then

		If @MON = "01" Or @MON = "02" Or @MON = "03" Then

			GUICtrlSetState($Label_ratio_q1, $gui_show)
			GUICtrlSetState($Label_ratio_q2, $gui_hide)
			GUICtrlSetState($Label_ratio_q3, $gui_hide)
			GUICtrlSetState($Label_ratio_q4, $gui_hide)

			GUICtrlSetState($Input_RaTio_q1, $gui_show)
			GUICtrlSetState($Input_RaTio_q2, $gui_hide)
			GUICtrlSetState($Input_RaTio_q3, $gui_hide)
			GUICtrlSetState($Input_RaTio_q4, $gui_hide)

;~ 			GUICtrlSetData($Input_Quarter, "Q1")
			GUICtrlSetState($Group_Q1, $GUI_SHOW)
			GUICtrlSetData($Group_Q1, "|| Q1 - " & @YEAR & " ||")

		Else

			GUICtrlSetData($Group_Q1, "Q1 - " & $SelDate_slipt[1])

		EndIf



		If @MON = "04" Or @MON = "05" Or @MON = "06" Then

			GUICtrlSetState($Label_ratio_q1, $gui_hide)
			GUICtrlSetState($Label_ratio_q2, $gui_show)
			GUICtrlSetState($Label_ratio_q3, $gui_hide)
			GUICtrlSetState($Label_ratio_q4, $gui_hide)

			GUICtrlSetState($Input_RaTio_q1, $gui_hide)
			GUICtrlSetState($Input_RaTio_q2, $gui_show)
			GUICtrlSetState($Input_RaTio_q3, $gui_hide)
			GUICtrlSetState($Input_RaTio_q4, $gui_hide)

;~ 			GUICtrlSetData($Input_Quarter, "Q2")
			GUICtrlSetState($Group_Q2, $GUI_SHOW)
			GUICtrlSetData($Group_Q2, "|| Q2 - " & $SelDate_slipt[1] & " ||")

		Else

			GUICtrlSetData($Group_Q2, "Q2 - " & $SelDate_slipt[1])

		EndIf

		If @MON = "07" Or @MON = "08" Or @MON = "09" Then

			GUICtrlSetState($Label_ratio_q1, $gui_hide)
			GUICtrlSetState($Label_ratio_q2, $gui_hide)
			GUICtrlSetState($Label_ratio_q3, $gui_show)
			GUICtrlSetState($Label_ratio_q4, $gui_hide)

			GUICtrlSetState($Input_RaTio_q1, $gui_hide)
			GUICtrlSetState($Input_RaTio_q2, $gui_hide)
			GUICtrlSetState($Input_RaTio_q3, $gui_show)
			GUICtrlSetState($Input_RaTio_q4, $gui_hide)

;~ 			GUICtrlSetData($Input_Quarter, "Q3")
			GUICtrlSetState($Group_Q3, $GUI_SHOW)
			GUICtrlSetData($Group_Q3, "|| Q3 - " & $SelDate_slipt[1] & " ||")
		Else

			GUICtrlSetData($Group_Q3, "Q3 - " & $SelDate_slipt[1])

		EndIf

		If @MON = "10" Or @MON = "11" Or @MON = "12" Then

			GUICtrlSetState($Label_ratio_q1, $gui_hide)
			GUICtrlSetState($Label_ratio_q2, $gui_hide)
			GUICtrlSetState($Label_ratio_q3, $gui_hide)
			GUICtrlSetState($Label_ratio_q4, $gui_show)

			GUICtrlSetState($Input_RaTio_q1, $gui_hide)
			GUICtrlSetState($Input_RaTio_q2, $gui_hide)
			GUICtrlSetState($Input_RaTio_q3, $gui_hide)
			GUICtrlSetState($Input_RaTio_q4, $gui_show)

;~ 			GUICtrlSetData($Input_Quarter, "Q4")
			GUICtrlSetState($Group_Q4, $GUI_SHOW)
			GUICtrlSetData($Group_Q4, "|| Q4 - " & $SelDate_slipt[1] & " ||")
		Else

			GUICtrlSetData($Group_Q4, "Q4 - " & $SelDate_slipt[1])

		EndIf


	Else
		GUICtrlSetState($Group_Q1, $GUI_SHOW)
		GUICtrlSetState($Label_ratio_q1, $gui_hide)
		GUICtrlSetState($Label_ratio_q2, $gui_hide)
		GUICtrlSetState($Label_ratio_q3, $gui_hide)
		GUICtrlSetState($Label_ratio_q4, $gui_hide)

		GUICtrlSetState($Input_RaTio_q1, $gui_hide)
		GUICtrlSetState($Input_RaTio_q2, $gui_hide)
		GUICtrlSetState($Input_RaTio_q3, $gui_hide)
		GUICtrlSetState($Input_RaTio_q4, $gui_hide)
	EndIf




	Return


EndFunc   ;==>_CheckQuarter

Func _CriaINI($Year)

	Local $sJulDate1 = _DateToDayValue($Year, "12", "31")
	For $i = 0 To 365 Step 1

		Local $y, $m, $d
		$sJulDate = _DayValueToDate($sJulDate1 - $i, $y, $m, $d)
		If $y = $Year Then
			$Wday = _DateToDayOfWeek($Year, $m, $d)
			If $Wday = 1 Or $Wday = 7 Then
				If RegRead($DB & "\" & $Year & "\" & $m, $d) = "" Then
					RegWrite($DB & "\" & $Year & "\" & $m, $d, "REG_SZ", "W")
				EndIf
			Else
				RegWrite($DB & "\" & $Year & "\" & $m, $d, "REG_SZ", RegRead($DB & "\" & $Year & "\" & $m, $d))
			EndIf
		EndIf
	Next
;~ 	_CreateMenu()
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

	$Form_Colors = GUICreate('Colors', 230, 500, 300, 100, $DS_MODALFRAME, BitOR($WS_EX_TOPMOST, $WS_EX_MDICHILD), $Form_WorkDays)

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
	GUICtrlCreateLabel("Graphic line:", 10, 315)
	GUICtrlCreateLabel("Quarter line:", 10, 345)
	GUICtrlCreateLabel("Border size:", 10, 375)
	$Slider_Border_Size = GUICtrlCreateSlider(65, 370, 140, 20, BitOR($GUI_SS_DEFAULT_SLIDER, $TBS_FIXEDLENGTH))
	GUICtrlSetLimit($Slider_Border_Size, 5, 0)
	GUICtrlSetData($Slider_Border_Size, $g_iQuarterBorderSize)
	$Label_Border_Size = GUICtrlCreateLabel(GUICtrlRead($Slider_Border_Size), 205, 373)

	GUICtrlCreateLabel("Font size:", 5, 405)
	$Slider_Font_Size = GUICtrlCreateSlider(65, 400, 140, 20, BitOR($GUI_SS_DEFAULT_SLIDER, $TBS_FIXEDLENGTH))
	GUICtrlSetLimit($Slider_Font_Size, 25, 10)
	GUICtrlSetData($Slider_Font_Size, $g_iListViewFontHeight)
	$Label_Font_Size = GUICtrlCreateLabel(GUICtrlRead($Slider_Font_Size), 205, 403)




;~ 	Global $g_iQuarterBorderSize = RegRead($DB, "Quarter_Border_Size")
;~ If @error Then $g_iQuarterBorderSize = 2


	$debug_check = GUICtrlCreateCheckbox("Dev tools", 10, 375)
	If $Debug = 1 Then
		GUICtrlSetState($debug_check, $gui_checked)
	Else
		GUICtrlSetState($debug_check, $gui_unchecked)
	EndIf
	GUICtrlSetState($debug_check, $gui_hide)

	;   $CP_FLAG_CHOOSERBUTTON
	;   $CP_FLAG_TIP
	;   $CP_FLAG_MAGNIFICATION
	;   $CP_FLAG_ARROWSTYLE
	;   $CP_FLAG_HANDCURSOR (don't used)
	;   $CP_FLAG_MOUSEWHEEL

	$Picker_OnSite = _GUIColorPicker_Create('', 70, 10, 60, 23, $Color_bk_OnSite, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Remote = _GUIColorPicker_Create('', 70, 40, 60, 23, $Color_bk_Remote, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Holiday = _GUIColorPicker_Create('', 70, 70, 60, 23, $Color_bk_holiday, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_PTO = _GUIColorPicker_Create('', 70, 100, 60, 23, $Color_bk_PTO, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Travel = _GUIColorPicker_Create('', 70, 130, 60, 23, $Color_bk_Travel, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Sick = _GUIColorPicker_Create('', 70, 160, 60, 23, $Color_bk_Sick, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Blank = _GUIColorPicker_Create('', 70, 190, 60, 23, $Color_bk_Blank, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Weekend = _GUIColorPicker_Create('', 70, 220, 60, 23, $Color_bk_Weekend, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Today = _GUIColorPicker_Create('', 70, 250, 60, 23, $Color_bk_Today, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Selected = _GUIColorPicker_Create('', 70, 280, 60, 23, $Color_bk_Selected, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Graphic = _GUIColorPicker_Create('', 70, 310, 60, 23, $Color_bk_Graphic, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Quarter = _GUIColorPicker_Create('', 70, 340, 60, 23, $g_clrQuarterBorder, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')

	$Picker_Font_OnSite = GUICtrlCreateCheckbox("White Font", 135, 10)
	$Picker_Font_Remote = GUICtrlCreateCheckbox("White Font", 135, 40)
	$Picker_Font_Holiday = GUICtrlCreateCheckbox("White Font", 135, 70)
	$Picker_Font_PTO = GUICtrlCreateCheckbox("White Font", 135, 100)
	$Picker_Font_Travel = GUICtrlCreateCheckbox("White Font", 135, 130)
	$Picker_Font_Sick = GUICtrlCreateCheckbox("White Font", 135, 160)
	$Picker_Font_Blank = GUICtrlCreateCheckbox("White Font", 135, 190)
	$Picker_Font_Weekend = GUICtrlCreateCheckbox("White Font", 135, 220)
	$Picker_Font_Graphic = GUICtrlCreateCheckbox("No Line", 135, 311)

	GUICtrlSetState($Picker_Font_OnSite, $Picker_Font_OnSite_Read)
	GUICtrlSetState($Picker_Font_Remote, $Picker_Font_Remote_Read)
	GUICtrlSetState($Picker_Font_Holiday, $Picker_Font_Holiday_Read)
	GUICtrlSetState($Picker_Font_PTO, $Picker_Font_PTO_Read)
	GUICtrlSetState($Picker_Font_Travel, $Picker_Font_Travel_Read)
	GUICtrlSetState($Picker_Font_Sick, $Picker_Font_Sick_Read)
	GUICtrlSetState($Picker_Font_Blank, $Picker_Font_Blank_Read)
	GUICtrlSetState($Picker_Font_Weekend, $Picker_Font_Weekend_Read)
	GUICtrlSetState($Picker_Font_Graphic, $Picker_Font_Graphic_Read)

	If GUICtrlRead($Picker_Font_Graphic) = 1 Then
		GUICtrlSetState($Picker_Graphic, $GUI_DISABLE)
	Else
		GUICtrlSetState($Picker_Graphic, $GUI_ENABLE)
	EndIf

	$Original_Color_1 = $Color_bk_OnSite & $Color_bk_Remote & $Color_bk_holiday & $Color_bk_PTO & _
			$Color_bk_Travel & $Color_bk_Sick & $Color_bk_Blank & $Color_bk_Weekend & $Color_bk_Today & _
			$Color_bk_Selected & $g_clrQuarterBorder & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & _
			$Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read & $g_iQuarterBorderSize


	$Colors_Close = GUICtrlCreateButton("Close", 85, 430, 70, 30)

	GUISetState(@SW_SHOW, $Form_Colors)

	While 1
		$Msg = GUIGetMsg()
		Switch $Msg
			Case $debug_check
				RegWrite($DB, "Debug", "REG_SZ", GUICtrlRead($debug_check))
				$Debug = GUICtrlRead($debug_check)

			Case $Slider_Border_Size
				GUICtrlSetData($Label_Border_Size, GUICtrlRead($Slider_Border_Size))

			Case $Slider_Font_Size
				GUICtrlSetData($Label_Font_Size, GUICtrlRead($Slider_Font_Size))


			Case $Picker_Font_Graphic
				If GUICtrlRead($Picker_Font_Graphic) = 1 Then
					GUICtrlSetState($Picker_Graphic, $GUI_DISABLE)
				Else
					GUICtrlSetState($Picker_Graphic, $GUI_ENABLE)
				EndIf

;~ 			Case $P_Graphic
;~ 				$Picker_Color_Graphic = _GUIColorPicker_GetColor($Picker_Graphic)
;~ 				GUICtrlSetBkColor($Label_Graphic,$Picker_Color_Graphic)
;~ 				MsgBox(262144, "1", $Color_bk_Graphic & @CRLF & $Picker_Color_Graphic)

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
				$Picker_Color_Graphic = _GUIColorPicker_GetColor($Picker_Graphic)

				$Picker_Color_Quarter = _GUIColorPicker_GetColor($Picker_Quarter)
				$g_clrQuarterBorder = _GUIColorPicker_GetColor($Picker_Quarter)

				_GUIColorPicker_Release($Form_Colors)


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
				RegWrite($DB, "Color_Graphic", "REG_SZ", $Picker_Color_Graphic)
				RegWrite($DB, "Color_Quarter", "REG_SZ", $Picker_Color_Quarter)

				$g_iQuarterBorderSize = GUICtrlRead($Slider_Border_Size)
				RegWrite($DB, "Quarter_Border_Size", "REG_SZ", $g_iQuarterBorderSize)

				$g_iListViewFontHeight = GUICtrlRead($Slider_Font_Size)
				RegWrite($DB, "Font_Size", "REG_SZ", $g_iListViewFontHeight)





				$Picker_Font_OnSite_Read = GUICtrlRead($Picker_Font_OnSite)
				$Picker_Font_Remote_Read = GUICtrlRead($Picker_Font_Remote)
				$Picker_Font_Holiday_Read = GUICtrlRead($Picker_Font_Holiday)
				$Picker_Font_PTO_Read = GUICtrlRead($Picker_Font_PTO)
				$Picker_Font_Travel_Read = GUICtrlRead($Picker_Font_Travel)
				$Picker_Font_Sick_Read = GUICtrlRead($Picker_Font_Sick)
				$Picker_Font_Blank_Read = GUICtrlRead($Picker_Font_Blank)
				$Picker_Font_Weekend_Read = GUICtrlRead($Picker_Font_Weekend)
				$Picker_Font_Graphic_Read = GUICtrlRead($Picker_Font_Graphic)

				RegWrite($DB, "Font_OnSite", "REG_SZ", $Picker_Font_OnSite_Read)
				RegWrite($DB, "Font_Remote", "REG_SZ", $Picker_Font_Remote_Read)
				RegWrite($DB, "Font_holiday", "REG_SZ", $Picker_Font_Holiday_Read)
				RegWrite($DB, "Font_PTO", "REG_SZ", $Picker_Font_PTO_Read)
				RegWrite($DB, "Font_Travel", "REG_SZ", $Picker_Font_Travel_Read)
				RegWrite($DB, "Font_Sick", "REG_SZ", $Picker_Font_Sick_Read)
				RegWrite($DB, "Font_Blank", "REG_SZ", $Picker_Font_Blank_Read)
				RegWrite($DB, "Font_Weekend", "REG_SZ", $Picker_Font_Weekend_Read)
				RegWrite($DB, "Font_Graphic", "REG_SZ", $Picker_Font_Graphic_Read)

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

				$Color_bk_Graphic = $Picker_Color_Graphic
				$Color_Graphic_Transparent = $Picker_Font_Graphic_Read
;~ 				MsgBox(262144, "2", $Color_bk_Graphic & @CRLF & $Picker_Color_Graphic)

				GUICtrlSetColor($Button_OnSite, $Font_OnSite)
				GUICtrlSetColor($Button_Remote, $Font_Remote)
				GUICtrlSetColor($Button_holiday, $Font_Holiday)
				GUICtrlSetColor($Button_PTO, $Font_PTO)
				GUICtrlSetColor($Button_Travel, $Font_Travel)
				GUICtrlSetColor($Button_Sick, $Font_Sick)
				GUICtrlSetColor($Button_Blank, $Font_Blank)
				GUICtrlSetColor($Button_Weekend, $Font_Weekend)


				$Original_Color_2 = $Picker_Color_OnSite & $Picker_Color_Remote & $Picker_Color_Holiday & $Picker_Color_PTO & _
						$Picker_Color_Travel & $Picker_Color_Sick & $Picker_Color_Blank & $Picker_Color_Weekend & $Picker_Color_Today & _
						$Picker_Color_Selected & $Picker_Quarter & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & _
						$Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read & $g_iQuarterBorderSize

				GUIDelete($Form_Colors)

				WinActivate("Work Days") ;,"",@SW_SHOW )

				_Chart()

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

	Global $g_clrQuarterBorder = RegRead($DB, "Color_Quarter")
	If @error Then $g_clrQuarterBorder = 0xE0E0E0

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

	Return


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

			For $d = 1 To 10000

				If $d < 10 Then
					$D1 = "0" & $d
				Else
					$D1 = $d
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

Func _GetColorGradient($value) ;Defines the bk colors for the Ratio according to the value
	If $value = "-" Then
		Return "0x007ECD"
	Else
		; Limita o valor m�nimo
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

		; Calcula a raz�o de interpola��o entre os dois pontos
		Local $ratio = ($value - $v1) / ($v2 - $v1)

		; Interpola cada canal de cor
		Local $r = _Interpolate($r1, $r2, $ratio)
		Local $g = _Interpolate($g1, $g2, $ratio)
		Local $b = _Interpolate($b1, $b2, $ratio)

		; Retorna em formato hexadecimal
		Return "0x" & StringFormat("%02X%02X%02X", $r, $g, $b)
	EndIf
EndFunc   ;==>_GetColorGradient

Func _GetColorFromValue($iValue) ;Defines the bk colors for the Remaning days according to the value
;~
	; Limita o intervalo
	If $iValue < -30 Then $iValue = -30
	If $iValue > 31 Then $iValue = 31

	; Valores de 0 ou menores = verde
	If $iValue <= 0 Then
		Return 0x00FF00 ; RGB(0,255,0)
	EndIf

	; Valores de 1 a 31: gradiente de amarelo a vermelho
	; Verde varia de 255 (no valor 1) at� 0 (no valor 31)
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
	$About_Text = "Work Days is a user-friendly calendar-based application for managing and categorizing your workdays like On Site, Remote, and Holiday, throughout the year." & @CRLF & @CRLF & "Developed by Fabricio Zambroni - CURRENT VERSION: " & FileGetVersion(@ScriptFullPath)
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
