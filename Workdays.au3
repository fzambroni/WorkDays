#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Icon=xcalendar4.ico
#AutoIt3Wrapper_Res_Description=Work Day management
#AutoIt3Wrapper_Res_Fileversion=2.1.4.15
#AutoIt3Wrapper_Res_ProductVersion=2.1.0.0
#AutoIt3Wrapper_Res_ProductName=Work Days
#AutoIt3Wrapper_Res_CompanyName=Fabricio Zambroni
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2026 Fabricio Zambroni
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\splash.jpg
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\Help.html
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\Updater.exe
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\About.db
#AutoIt3Wrapper_Res_File_Add=E:\GitHub\WorkDays\Workdays_Outlook_Agent.exe
#AutoIt3Wrapper_Run_After=E:\GitHub\WorkDays\FileUpdate.exe
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

; Script Start - Add your code below here
#NoTrayIcon
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
#include <ScrollBarConstants.au3>
#include <GuiEdit.au3>
#include <TabConstants.au3>
#include <File.au3>
#include <String.au3>
#include <InetConstants.au3>
#include "Updater_lib2.au3"
#include "Workdays_Backup.au3"


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

Global Const $g_clrTodayBorder = 0xFF0000 ; vermelho
Global Const $g_clrSelectedBorder = 0x00AA00 ; verde
Global $g_clrInvalidDayBG = 0xF0F0F0 ; disabled cells for dates that do not exist in the month
Global $g_clrInvalidDayFG = 0xA0A0A0

; Note: WM_SETREDRAW, RDW_INVALIDATE, RDW_ALLCHILDREN, RDW_UPDATENOW
; are already declared by the included WindowsConstants.au3 / WinAPI.au3.
; Literal values are used inside _LockWindow/_CleanRepaint to avoid conflicts.
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

; Fast-selection state
Global $g_iLVYear = 0
Global $g_iSelDay = 0
Global $g_iSelMonth = 0

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
Global $g_hCal = 0   ; MonthCal HWND – hidden, kept for programmatic date access

; Custom calendar widget – replaces the MonthCal visually.
; Uses labels so GUICtrlSetBkColor works reliably without any GDI/NM_CUSTOMDRAW.
Global $g_ccYear = @YEAR
Global $g_ccMonth = Number(@MON)
Global $g_ccPrev = 0               ; "< " button
Global $g_ccNext = 0               ; " >" button
Global $g_ccTitle = 0              ; month+year button (opens picker)
Global $g_ccToday = 0              ; "Today" button
Global $g_ccDayCells[42]           ; 6 rows x 7 cols day-cell label IDs (content)
Global $g_ccFrames[42]             ; border frame labels behind each day cell
Global $g_ccMarkers[42]            ; 2px bar at bottom of cell = visual note marker
Global $g_ccWeekNums[6]            ; 6 week-number label IDs
Global $g_ccDayValues[42]          ; day number in each cell (0 = empty)
Global $g_ccCacheTitle = Chr(0)
Global $g_ccCacheWeekNums[6]
Global $g_ccCacheText[42]
Global $g_ccCacheDayBG[42]
Global $g_ccCacheDayFG[42]
Global $g_ccCacheFontW[42]
Global $g_ccCacheFrame[42]
Global $g_ccCacheMarker[42]
Global $g_ccCacheTip[42]

Global $iYear = @YEAR
Global $DB = "HKEY_CURRENT_USER\Software\WorkDays"

Global $HelpFile = @ScriptDir & "\Help.html"
Global $sSplashPath = @ScriptDir & "\splash.jpg"
Global $AboutFile = $sSplashPath
Global $AboutDBFile = @ScriptDir & "\About.db"
Global $g_sOutlookAgentDB = $DB & "\OutlookAgent"
Global $g_sOutlookAgentDir = @ScriptDir
Global $g_sOutlookAgentExe = $g_sOutlookAgentDir & "\Workdays_Outlook_Agent.exe"
Global $g_sOutlookAgentLog = $g_sOutlookAgentDir & "\Workdays_Outlook_Agent.log"
Global $g_sOutlookAgentState = $g_sOutlookAgentDir & "\Workdays_Outlook_Agent_State.ini"
Global $g_sOutlookAgentProcess = "Workdays_Outlook_Agent.exe"
Global $g_bOutlookAgentRefreshPending = False
Global $g_sOutlookAgentPendingSeq = ""
Global $g_sOutlookAgentPendingDate = ""
Global $g_sOutlookAgentPendingStatus = ""
Global $g_bOutlookAgentSyncBlockedPending = False
Global $g_sOutlookAgentSyncBlockedReason = ""
Global $g_sOutlookAgentSyncBlockedPlanFile = ""
Global $g_sOutlookAgentLastGuardStatus = ""
Global $g_hOutlookAgentSettingsWindow = 0
Global $g_bWorkDaysUpdaterAvailable = False
Global $ResetPosition = 0
Global $Progress_Splash, $Form_Splash, $Label_Percentage, $Splash, $Button_Close_Splash

;Chart Variables
Global $Total, $Count_O, $Count_R, $Count_H, $Count_P, $Count_T, $Count_S, $Count_B, $Count_W, $Percentage_O, $Degrees_O, $Percentage_R, $Degrees_R, $Percentage_H, $Degrees_H, $Percentage_P, $Degrees_P, $Percentage_T, $Degrees_T, $Percentage_S, $Degrees_S, $Percentage_B, $Degrees_B, $Percentage_W, $Degrees_W, $Chart, $Color_Graphic_Transparent = 1
Global $g_sMainGridFilter = "" ; Active Year Summary/chart filter applied to the main grid. Empty = show all categories.


;Colors Variables
Global $Color_bk_OnSite, $Color_bk_Remote, $Color_bk_holiday, $Color_bk_PTO, $Color_bk_Travel, $Color_bk_Sick, $Color_bk_Blank, $Color_bk_Weekend, $Color_HighlightDate

_CheckSingleInstance()


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
_PosScreen()


FileInstall("splash.jpg", $sSplashPath, 1)
_splash("on")


; ----------------------------------------------------------------------------------------------------------------------
; Updater - GitHub based
; ----------------------------------------------------------------------------------------------------------------------
Global $GitHubAppName = "WorkDays"
; Skip automatic update checks when running the .au3 directly from SciTE/dev mode.
If Not StringInStr(StringLower(@ScriptName), ".au3") Then
	_CheckGitHubUpdate()
	If Not ProcessExists("Workdays_Outlook_Agent.exe") Then
		FileInstall("Workdays_Outlook_Agent.exe", $g_sOutlookAgentExe, 1)
	EndIf
EndIf









Func _PosScreen()
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

EndFunc   ;==>_PosScreen

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
Global $DBpMenu_Delete_Year[100]
Global $DBpMenu_Delete_Date[15]

Global $DBpMenu_Report_simple_Year[100]
Global $DBpMenu_Report_detailed_Year[100]
Global $DBpMenu_Report_professional_Year[100]
Global $DBpMenu_Report_Year[100]
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

Global $Color_HighlightDate = RegRead($DB, "Color_HighlightDate")
If @error Then $Color_HighlightDate = 0xFF0000

Global $g_clrInvalidDayBG = RegRead($DB, "Color_InvalidDay")
If @error Then $g_clrInvalidDayBG = 0xF0F0F0

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
If @error Or Number($Picker_Grid_Size_X_Read) < 25 Then $Picker_Grid_Size_X_Read = 35
$Picker_Grid_Size_X_Read = Number($Picker_Grid_Size_X_Read)
If $Picker_Grid_Size_X_Read > 45 Then $Picker_Grid_Size_X_Read = 45

Global $Picker_Grid_Size_Y_Read = RegRead($DB, "Grid_Size_Y")
If @error Then $Picker_Grid_Size_Y_Read = 0xFF0000


$Form_WorkDays = GUICreate("Work Days", $Window_X, $Window_Y, $WinPos_X, $WinPos_Y)
If $Form_WorkDays = 0 Then Exit MsgBox(16, "Error", "Failed to create main window.")

$g_hGUI = $Form_WorkDays
If $g_hGUI = 0 Then Exit MsgBox(16, "Error", "Failed to store GUI handle.")

Global $DBpMenu_db = GUICtrlCreateMenu("File")
Global $BkpMenu_Exit = GUICtrlCreateMenuItem("&Exit", $DBpMenu_db)

;~ Global $DBpMenu_backup_Data = GUICtrlCreateMenu("Data")
;~ Global $DBpMenu_backup = GUICtrlCreateMenuItem("Create Backup", $DBpMenu_backup_Data)
;~ Global $BkpMenu_Batch = GUICtrlCreateMenuItem("Restore Backup", $DBpMenu_backup_Data)
;~ Global $DBpMenu_backup_2 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
;~ Global $BkpMenu_reset_all1 = GUICtrlCreateMenu("Data Management", $DBpMenu_backup_Data)
;~ Global $BkpMenu_reset_all = GUICtrlCreateMenuItem("Reset Entire Database", $BkpMenu_reset_all1)
;~ Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
;~ Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_backup_Data)
;~ Global $DBpMenu_backup_Data_Holidays = GUICtrlCreateMenuItem("Import Holidays File", $DBpMenu_backup_Data)

Global $DBpMenu_settings = GUICtrlCreateMenu("Settings")
Global $BkpMenu_settings_BKcolors = GUICtrlCreateMenuItem("Preferences", $DBpMenu_settings)
Global $BkpMenu_settings_OutlookAgent = GUICtrlCreateMenuItem("Outlook Agent (Experimental)", $DBpMenu_settings)
;~ Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_settings)

Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_settings)
Global $DBpMenu_backup_Data_Holidays = GUICtrlCreateMenuItem("Import Special days", $DBpMenu_settings)
Global $DBpMenu_backup_3 = GUICtrlCreateMenuItem("", $DBpMenu_settings)
Global $BkpMenu_reset_all1 = GUICtrlCreateMenu("Data Management", $DBpMenu_settings)
Global $BkpMenu_Backup = GUICtrlCreateMenu("Backup", $BkpMenu_reset_all1)
Global $DBpMenu_backup = GUICtrlCreateMenuItem("Create Backup", $BkpMenu_Backup)
Global $BkpMenu_Batch = GUICtrlCreateMenuItem("Restore Backup", $BkpMenu_Backup)
Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
Global $DBpMenu_backup_4 = GUICtrlCreateMenuItem("", $BkpMenu_reset_all1)
Global $BkpMenu_reset_all = GUICtrlCreateMenuItem("Reset Entire Database", $BkpMenu_reset_all1)
Global $DBpMenu_backup_5 = GUICtrlCreateMenuItem("", $BkpMenu_reset_all1)

Global $DBpMenu_Report = GUICtrlCreateMenu("Report")
Global $DBpMenu_Report_Simple = GUICtrlCreateMenu("Simple", $DBpMenu_Report)
Global $DBpMenu_Report_Detailed = GUICtrlCreateMenu("Detailed", $DBpMenu_Report)
Global $DBpMenu_Report_Professional = GUICtrlCreateMenu("Analytical", $DBpMenu_Report)

;~ Global $BkpMenu_settings_ResetScreen = GUICtrlCreateMenuItem("Reset Screen Position", $DBpMenu_settings)
Global $BkpMenu_help = GUICtrlCreateMenu("?")
Global $BkpMenu_help_help = GUICtrlCreateMenuItem("Help", $BkpMenu_help)
Global $BkpMenu_help_space = GUICtrlCreateMenuItem("", $BkpMenu_help)
Global $BkpMenu_help_About = GUICtrlCreateMenuItem("About", $BkpMenu_help)

#EndRegion GLOBAL

$Calendar = GUICtrlCreateMonthCal(@YEAR & "/" & @MON & "/" & @MDAY, 8, 8, 273, 201, $MCS_WEEKNUMBERS)
$g_hCal = GUICtrlGetHandle($Calendar)
; Hide the MonthCal – the custom colored calendar replaces it visually.
; All GUICtrlRead($Calendar) / GUICtrlSetData($Calendar,...) calls still work on the hidden control.
GUICtrlSetState($Calendar, $GUI_HIDE)

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
GUICtrlSetColor($Button_Update, 0xFFFFFF)
GUICtrlSetBkColor($Button_Update, 0xFF0000)
GUICtrlSetFont($Button_Update, 9, 900)
GUICtrlSetState($Button_Update, $GUI_HIDE)


;~ $Button_OnSite = GUICtrlCreateButton("&On Site", 296, 84, 75, 25)
$Button_OnSite = GUICtrlCreateButton("&On Site", 384, 84, 75, 25)
GUICtrlSetBkColor($Button_OnSite, $Color_bk_OnSite)
GUICtrlSetColor($Button_OnSite, $Font_OnSite)


;~ $Button_Remote = GUICtrlCreateButton("&Remote", 384, 84, 75, 25)
$Button_Remote = GUICtrlCreateButton("&Remote", 296, 84, 75, 25)
GUICtrlSetBkColor($Button_Remote, $Color_bk_Remote)
GUICtrlSetColor($Button_Remote, $Font_Remote)

;~ $Button_holiday = GUICtrlCreateButton("&Holiday", 296, 114, 75, 25)
$Button_holiday = GUICtrlCreateButton("&Holiday", 384, 144, 75, 25)
GUICtrlSetBkColor($Button_holiday, $Color_bk_holiday)
GUICtrlSetColor($Button_holiday, $Font_Holiday)

$Button_OutlookSync = GUICtrlCreateButton("Sync", 472, 144, 75, 25)
GUICtrlSetTip($Button_OutlookSync, "Force an immediate Outlook Agent sync.")
GUICtrlSetBkColor($Button_OutlookSync, 0xD9ECFF)
GUICtrlSetColor($Button_OutlookSync, 0x0B4F8A)
GUICtrlSetFont($Button_OutlookSync, 9, 700)

;~ $Button_PTO = GUICtrlCreateButton("&PTO", 384, 114, 75, 25)
$Button_PTO = GUICtrlCreateButton("&PTO", 296, 114, 75, 25)
GUICtrlSetBkColor($Button_PTO, $Color_bk_PTO)
GUICtrlSetColor($Button_PTO, $Font_PTO)

;~ $Button_Travel = GUICtrlCreateButton("&Travel", 296, 144, 75, 25)
$Button_Travel = GUICtrlCreateButton("&Travel", 384, 114, 75, 25)
GUICtrlSetBkColor($Button_Travel, $Color_bk_Travel)
GUICtrlSetColor($Button_Travel, $Font_Travel)

;~ $Button_Sick = GUICtrlCreateButton("&Sick", 384, 144, 75, 25)
$Button_Sick = GUICtrlCreateButton("&Sick", 296, 144, 75, 25)
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

; Build the custom colored calendar (replaces the hidden MonthCal visually)
_CustomCal_Create()

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
_ReadStatistics(@YEAR)   ; populate stats panels (normally done by _Update, but startup calls _ReadINI directly)
_CustomCal_Update()      ; paint custom calendar with initial month's category colors
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
; Initialize the selected-day fields with today's persisted data.
; Startup previously highlighted today's cell, but did not load the associated marker/tag
; into $Input_Tag until the user clicked a calendar day.
$SelDate = @YEAR & "/" & StringFormat("%02d", @MON) & "/" & StringFormat("%02d", @MDAY)
GUICtrlSetData($Calendar, $SelDate)
_GUICtrlMonthCal_SetCurSel($Calendar, @YEAR, Number(@MON), Number(@MDAY))
_RefreshSelectedDateUI($SelDate)

;~ $UpdatedVersion = FileGetVersion($UpdatePath & "\WorkDays.exe")

GUISetState(@SW_SHOW, $Form_WorkDays)

ConsoleWrite("Window is visible: " & _Monitor_IsVisibleWindow($Form_WorkDays) & @CRLF)

GUIDelete($Form_Splash)
If Not StringInStr(StringLower(@ScriptName), ".au3") Then
	FileDelete($sSplashPath)
EndIf


$currentVersion = FileGetVersion(@ScriptDir & "\WorkDays.exe")

;~ ConsoleWrite("$UpdatedVersion:" & $UpdatedVersion & @CRLF)
;~ ConsoleWrite("$currentVersion:" & $currentVersion & @CRLF)
#cs
If $UpdatedVersion > $currentVersion Then
	FileCopy($UpdatePath & "\WorkDays.exe", @ScriptDir & "\WorkDays.tmp", 9)
EndIf
#ce
If FileExists(@ScriptDir & "\WorkDays.tmp") Then
	$g_bWorkDaysUpdaterAvailable = True
	GUICtrlSetData($Button_Update, "UPDATE AVAILABLE - Click to execute")
	GUICtrlSetColor($Button_Update, 0xFFFFFF)
	GUICtrlSetBkColor($Button_Update, 0xFF0000)
	GUICtrlSetState($Button_Update, $GUI_SHOW)
Else
	$g_bWorkDaysUpdaterAvailable = False
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

	_OutlookAgent_CheckRefreshNotification()

	; ── Custom calendar: prev/next navigation and day clicks ─────────
	If $nMsg = $g_ccPrev Then
		_CustomCal_Navigate(-1)
	ElseIf $nMsg = $g_ccNext Then
		_CustomCal_Navigate(1)
	ElseIf $nMsg = $g_ccTitle Then
		_CustomCal_ShowPicker()
	ElseIf $nMsg = $g_ccToday Then
		; Navigate to today.
		; Important: do NOT update $g_ccYear/$g_ccMonth before _CalendarRead().
		; When the year is the same but the month is different, the fast path inside
		; _CalendarRead() relies on the *currently displayed* custom-calendar month to
		; decide whether it needs a full redraw. Pre-setting these globals here makes
		; the code think the calendar is already on the target month, so the UI may keep
		; showing the old month even though the selected date changed.
		Local $sTodayDate = @YEAR & "/" & StringFormat("%02d", @MON) & "/" & StringFormat("%02d", @MDAY)
		GUICtrlSetData($Calendar, $sTodayDate)
		_GUICtrlMonthCal_SetCurSel($Calendar, @YEAR, Number(@MON), Number(@MDAY))
		If @YEAR <> $iYear Then
			_CriaINI(@YEAR)
		EndIf
		_CalendarRead()
	Else
		For $__ci = 0 To 41
			If ($nMsg = $g_ccDayCells[$__ci] Or $nMsg = $g_ccFrames[$__ci]) And $g_ccDayValues[$__ci] > 0 Then
				Local $__sD = StringFormat("%02d", $g_ccDayValues[$__ci])
				Local $__sM = StringFormat("%02d", $g_ccMonth)
				GUICtrlSetData($Calendar, $g_ccYear & "/" & $__sM & "/" & $__sD)
				_GUICtrlMonthCal_SetCurSel($Calendar, $g_ccYear, $g_ccMonth, $g_ccDayValues[$__ci])
				_CalendarRead()
				ExitLoop
			EndIf
		Next
	EndIf

	If $g_bShowCellMenu Then
		$g_bShowCellMenu = False
		_MenuContextual($g_iMenuDay, $g_iMenuMonth, $g_iMenuYear)
	EndIf

	For $j = 1 To 99

		If $nMsg = $DBpMenu_Report_simple_Year[$j] And $DBpMenu_Report_simple_Year[$j] <> 0 Then
			$DBpMenu_Report_Date = GUICtrlRead($DBpMenu_Report_simple_Year[$j], 1)
			GenerateWorkdaysReportHTML($DBpMenu_Report_Date, 0)
		EndIf

		If $nMsg = $DBpMenu_Report_detailed_Year[$j] And $DBpMenu_Report_detailed_Year[$j] <> 0 Then
			$DBpMenu_Report_Date = GUICtrlRead($DBpMenu_Report_detailed_Year[$j], 1)
			GenerateWorkdaysReportHTML($DBpMenu_Report_Date, 1)
		EndIf

		If $nMsg = $DBpMenu_Report_professional_Year[$j] And $DBpMenu_Report_professional_Year[$j] <> 0 Then
			$DBpMenu_Report_Date = GUICtrlRead($DBpMenu_Report_professional_Year[$j], 1)
			GenerateWorkdaysProfessionalReportHTML($DBpMenu_Report_Date)
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
						; _Reload handles all state updates - no need to pre-set date fields
						_Reload()

						MsgBox(262208, "Delete Year", "Year Deleted with Success", 0, $Form_WorkDays)

					Else
						_Reload()
						MsgBox(262160, "Year Delete", "An error occurred while attempting to delete this value from the database.", 0, $Form_WorkDays)
					EndIf

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		EndIf

	Next

	; Day controls and context-menu items only exist for 12 months.
	; Keep this loop separated from the report/delete year loop above,
	; otherwise $Inputs[$i][$j] will be accessed with $j > 12.
	For $j = 1 To 12
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
;~ 				_Reload()
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
			If $g_bOutlookAgentSyncBlockedPending Then
				_OutlookAgent_ShowSyncBlockedGuard()
			ElseIf $g_bOutlookAgentRefreshPending Then
				_OutlookAgent_RefreshWorkDaysFromAgentChange()
			ElseIf $g_bWorkDaysUpdaterAvailable Or FileExists(@ScriptDir & "\WorkDays.tmp") Then
				$Updater_File = @TempDir & "\Updater.exe"
				FileInstall("Updater.exe", $Updater_File, 1)
				Sleep(500)
				Run(@TempDir & "\Updater.exe '" & @ScriptDir & "'")
;~ 				Run($Updater_File)
				Sleep(500)
				_HideListViewCellTip()
				Exit
			EndIf

		Case $Button_OutlookSync
			_OutlookAgent_RequestSyncNow()

		Case $Label_YSumary_Reset
			_Chart("", True)

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
			#cs
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
			#ce


		Case $BkpMenu_settings_OutlookAgent
			_OutlookAgent_SettingsWindow()

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

				$Color_HighlightDate = RegRead($DB, "Color_HighlightDate")
				If @error Then $Color_HighlightDate = 0xFF0000

				$g_clrInvalidDayBG = RegRead($DB, "Color_InvalidDay")
				If @error Then $g_clrInvalidDayBG = 0xF0F0F0

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
				If Not @error And $SelDate_slipt[0] = 3 Then
					_RefreshMainGridCellStyles(Number($SelDate_slipt[1]))
					$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
					$Status = StringTrimLeft($Status1, 1)
					_UpdateSelectionHighlight(Number($SelDate_slipt[3]), Number($SelDate_slipt[2]))
				EndIf

			EndIf

		Case $BkpMenu_help_help

			FileInstall("Help.html", $HelpFile, 1)

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
					Run(@ScriptFullPath)
					Exit
;~ 					_CalendarRead()
					; _CalendarRead -> _Update -> _ReadStatistics -> _Chart already called

				Case $iMsgBoxAnswer = 7 ;No

			EndSelect

		Case $DBpMenu_backup_Data_Holidays
			If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(262452, "Special Days Import", "**WARNING** Importing data will overwrite any existing records for the selected dates. Do you want to proceed?", 0, $Form_WorkDays)
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					_ImportHolidays()
					Run(@ScriptFullPath)
					Exit
;~ 					_Reload()
					; _Reload already calls _Chart() internally


				Case $iMsgBoxAnswer = 7 ;No

			EndSelect


		Case $Calendar
			_CalendarRead()
;~ 			_Chart()

		Case $DBpMenu_backup
			_CreateBackup()

		Case $BkpMenu_reset_all
			_ResetDatabase()
			Run(@ScriptFullPath)
			Exit
;~ 			$SelDate = GUICtrlRead($Calendar)
;~ 			$SelDate_slipt = StringSplit($SelDate, "/")
;~ 			_CriaINI(@YEAR)
;~ 			_Reload()
			; _Reload already calls _Chart() internally

		Case $Button_Reload
			_Reload()
			; _Reload already calls _Chart() internally - no need to call again

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

Func _About()

	#cs
	For reference only:
	###### $Form_WorkDays = GUICreate("Work Days", 1140, 620, -1, -1)
	###### $Form_About = GUICreate("About", 655, 617, 280, -40, $WS_SYSMENU,$WS_EX_MDICHILD,$Form_WorkDays)
	#ce
;~ $AboutFile

;~ 	Global $AboutFile = @TempDir & "\about.jpg"

	FileInstall("About.db", $AboutDBFile, 1)
	FileInstall("splash.jpg", $AboutFile, 1)
	$About = "ERROR REading about file"

	; Open the file for reading and store the handle to a variable.
	Local $hAboutFileOpen = FileOpen($AboutDBFile, $FO_READ)
	If $hAboutFileOpen <> -1 Then
		; Read the contents of the file using the handle returned by FileOpen.
		Local $About = FileRead($hAboutFileOpen)
	EndIf

	; Close the handle returned by FileOpen.
	FileClose($hAboutFileOpen)

	$Form_About = GUICreate("About", 655, 617, 280, -40, $WS_SYSMENU, $WS_EX_MDICHILD, $Form_WorkDays)
	$Pic_About = GUICtrlCreatePic($AboutFile, 5, 5, 640, 360)
	$About_Text = "Work Days is a user-friendly calendar-based application for managing and categorizing your workdays like On Site, Remote, and Holiday, throughout the year." & @CRLF & @CRLF & "Developed by Fabricio Zambroni - CURRENT VERSION: " & FileGetVersion(@ScriptFullPath)
	$Text_About = GUICtrlCreateEdit($About_Text, 5, 293, 640, 90, BitOR($ES_MULTILINE, $ES_READONLY), -1)
	GUICtrlSetFont($Text_About, 12)
	GUICtrlSetColor($Text_About, 0x2211FF)
	$Edit_About = GUICtrlCreateEdit($About, 5, 396, 640, 180, BitOR($ES_MULTILINE, $ES_READONLY, $ES_AUTOVSCROLL, $WS_VSCROLL), -1)

	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($Form_About)
				Return
		EndSwitch
	WEnd

EndFunc   ;==>_About


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

Func _CheckSingleInstance()
	Local $aList = ProcessList(@ScriptName)
	Local $n = 0
	For $i = 1 To $aList[0][0]
		If $aList[$i][0] = @ScriptName Then
			$n += 1
			If $n > 1 Then
				MsgBox($MB_OK + $MB_ICONHAND + 262144, "Multiple Instances", "You cannot run multiple instances of this application.", 30)
;~                 _FileWriteLog($g_sLogPath, "App closed – duplicate instance.")
				Exit
			EndIf
		EndIf
	Next
EndFunc   ;==>_CheckSingleInstance

Func _BKColorPallet()

	; Create custom (4 x 5) color palette
	Dim $aPalette[20] = _
			[0xFFFFFF, 0x000000, 0xC0C0C0, 0x808080, _
			0xFF0000, 0x800000, 0xFFFF00, 0x808000, _
			0x00FF00, 0x008000, 0x00FFFF, 0x008080, _
			0x0000FF, 0x000080, 0xFF00FF, 0x800080, _
			0xC0DCC0, 0xA6CAF0, 0xFFFBF0, 0xA0A0A4]

	$Form_Colors = GUICreate('Colors', 230, 640, 300, 30, $DS_MODALFRAME, BitOR($WS_EX_TOPMOST, $WS_EX_MDICHILD), $Form_WorkDays)

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
	GUICtrlCreateLabel("Highlight:", 10, 315)
	GUICtrlCreateLabel("Invalid date:", 10, 345)
	GUICtrlCreateLabel("Graphic line:", 10, 375)
	GUICtrlCreateLabel("Quarter line:", 10, 405)
	GUICtrlCreateLabel("Border size:", 10, 435)
	$Slider_Border_Size = GUICtrlCreateSlider(65, 430, 140, 20, BitOR($GUI_SS_DEFAULT_SLIDER, $TBS_FIXEDLENGTH))
	GUICtrlSetLimit($Slider_Border_Size, 5, 0)
	GUICtrlSetData($Slider_Border_Size, $g_iQuarterBorderSize)
	$Label_Border_Size = GUICtrlCreateLabel(GUICtrlRead($Slider_Border_Size), 205, 433)

	GUICtrlCreateLabel("Font size:", 10, 465)
	$Slider_Font_Size = GUICtrlCreateSlider(65, 460, 140, 20, BitOR($GUI_SS_DEFAULT_SLIDER, $TBS_FIXEDLENGTH))
	GUICtrlSetLimit($Slider_Font_Size, 25, 10)
	GUICtrlSetData($Slider_Font_Size, $g_iListViewFontHeight)
	$Label_Font_Size = GUICtrlCreateLabel(GUICtrlRead($Slider_Font_Size), 205, 463)

	GUICtrlCreateLabel("Cell size:", 10, 495)
	$Slider_Cell_Size = GUICtrlCreateSlider(65, 490, 140, 20, BitOR($GUI_SS_DEFAULT_SLIDER, $TBS_FIXEDLENGTH))
	GUICtrlSetLimit($Slider_Cell_Size, 45, 25)
	GUICtrlSetData($Slider_Cell_Size, $Picker_Grid_Size_X_Read)
	$Label_Cell_Size = GUICtrlCreateLabel(GUICtrlRead($Slider_Cell_Size), 205, 493)

	$Checkbox_ResetScreen = GUICtrlCreateCheckbox("Reset Screen Position", 10, 520)




;~ 	Global $g_iQuarterBorderSize = RegRead($DB, "Quarter_Border_Size")
;~ If @error Then $g_iQuarterBorderSize = 2



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
	$Picker_HighlightDate = _GUIColorPicker_Create('', 70, 310, 60, 23, $Color_HighlightDate, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_InvalidDay = _GUIColorPicker_Create('', 70, 340, 60, 23, $g_clrInvalidDayBG, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Graphic = _GUIColorPicker_Create('', 70, 370, 60, 23, $Color_bk_Graphic, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')
	$Picker_Quarter = _GUIColorPicker_Create('', 70, 400, 60, 23, $g_clrQuarterBorder, BitOR($CP_FLAG_CHOOSERBUTTON, $CP_FLAG_ARROWSTYLE, $CP_FLAG_MOUSEWHEEL), $aPalette, 4, 5, 0, '', 'More...')

	$Picker_Font_OnSite = GUICtrlCreateCheckbox("White Font", 135, 10)
	$Picker_Font_Remote = GUICtrlCreateCheckbox("White Font", 135, 40)
	$Picker_Font_Holiday = GUICtrlCreateCheckbox("White Font", 135, 70)
	$Picker_Font_PTO = GUICtrlCreateCheckbox("White Font", 135, 100)
	$Picker_Font_Travel = GUICtrlCreateCheckbox("White Font", 135, 130)
	$Picker_Font_Sick = GUICtrlCreateCheckbox("White Font", 135, 160)
	$Picker_Font_Blank = GUICtrlCreateCheckbox("White Font", 135, 190)
	$Picker_Font_Weekend = GUICtrlCreateCheckbox("White Font", 135, 220)
	$Picker_Font_Graphic = GUICtrlCreateCheckbox("No Line", 135, 371)

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
			$Color_bk_Selected & $Color_HighlightDate & $g_clrInvalidDayBG & $Color_bk_Graphic & $g_clrQuarterBorder & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & _
			$Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read & $g_iQuarterBorderSize & $g_iListViewFontHeight & $Picker_Grid_Size_X_Read


	$Colors_Close = GUICtrlCreateButton("Close", 85, 570, 70, 30)

	GUISetState(@SW_SHOW, $Form_Colors)

	While 1
		$Msg = GUIGetMsg()
		Switch $Msg

			Case $Slider_Border_Size
				GUICtrlSetData($Label_Border_Size, GUICtrlRead($Slider_Border_Size))

			Case $Slider_Font_Size
				GUICtrlSetData($Label_Font_Size, GUICtrlRead($Slider_Font_Size))

			Case $Slider_Cell_Size
				GUICtrlSetData($Label_Cell_Size, GUICtrlRead($Slider_Cell_Size))


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
				$Picker_Color_HighlightDate = _GUIColorPicker_GetColor($Picker_HighlightDate)
				$Picker_Color_InvalidDay = _GUIColorPicker_GetColor($Picker_InvalidDay)
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
				RegWrite($DB, "Color_HighlightDate", "REG_SZ", $Picker_Color_HighlightDate)
				RegWrite($DB, "Color_InvalidDay", "REG_SZ", $Picker_Color_InvalidDay)
				RegWrite($DB, "Color_Graphic", "REG_SZ", $Picker_Color_Graphic)
				RegWrite($DB, "Color_Quarter", "REG_SZ", $Picker_Color_Quarter)

				$g_iQuarterBorderSize = GUICtrlRead($Slider_Border_Size)
				RegWrite($DB, "Quarter_Border_Size", "REG_SZ", $g_iQuarterBorderSize)

				$g_iListViewFontHeight = GUICtrlRead($Slider_Font_Size)
				RegWrite($DB, "Font_Size", "REG_SZ", $g_iListViewFontHeight)

				$Picker_Grid_Size_X_Read = GUICtrlRead($Slider_Cell_Size)
				If Number($Picker_Grid_Size_X_Read) < 20 Then $Picker_Grid_Size_X_Read = 20
				If Number($Picker_Grid_Size_X_Read) > 60 Then $Picker_Grid_Size_X_Read = 60
				RegWrite($DB, "Grid_Size_X", "REG_SZ", $Picker_Grid_Size_X_Read)





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

				$Color_HighlightDate = $Picker_Color_HighlightDate
				$g_clrInvalidDayBG = $Picker_Color_InvalidDay
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
						$Picker_Color_Selected & $Picker_Color_HighlightDate & $Picker_Color_InvalidDay & $Picker_Color_Graphic & $Picker_Color_Quarter & $Picker_Font_OnSite_Read & $Picker_Font_Remote_Read & $Picker_Font_Holiday_Read & _
						$Picker_Font_PTO_Read & $Picker_Font_Travel_Read & $Picker_Font_Sick_Read & $Picker_Font_Blank_Read & $Picker_Font_Weekend_Read & $g_iQuarterBorderSize & $g_iListViewFontHeight & $Picker_Grid_Size_X_Read

				$Checkbox_ResetScreenStatus = GUICtrlRead($Checkbox_ResetScreen)

				GUIDelete($Form_Colors)

				WinActivate("Work Days") ;,"",@SW_SHOW )
				_ApplyMainGridCellSize($Picker_Grid_Size_X_Read)

				_Chart()

				If $Checkbox_ResetScreenStatus = $GUI_CHECKED Then
					RegWrite($DB, "WinPosX", "REG_SZ", "")
					RegWrite($DB, "WinPosY", "REG_SZ", "")
					$ResetPosition = 1
					If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
					$iMsgBoxAnswer = MsgBox(262208, "Workdays", "Window position restored to the default value.", 30)
					Select
						Case $iMsgBoxAnswer = -1 ;Timeout

						Case Else        ;OK

					EndSelect
				EndIf

				If $Original_Color_1 = $Original_Color_2 Then
					Return 0
				Else
					Return 1
				EndIf

		EndSwitch
	WEnd


EndFunc   ;==>_BKColorPallet


Func _DateHighlightRegName($CYear, $Month, $Day)
	Return "Highlight_" & StringFormat("%04d_%02d_%02d", Number($CYear), Number($Month), Number($Day))
EndFunc   ;==>_DateHighlightRegName


Func _IsHighlightedDate($CYear, $Month, $Day)
	Local $sHighlighted = RegRead($DB, _DateHighlightRegName($CYear, $Month, $Day))
	If @error Then Return False
	Return ($sHighlighted = "1")
EndFunc   ;==>_IsHighlightedDate


Func _GetDateFontColor($CYear, $Month, $Day, $Status)
	If _IsHighlightedDate($CYear, $Month, $Day) Then Return $Color_HighlightDate
	Return _ColorFromDateFont($Status)
EndFunc   ;==>_GetDateFontColor


Func _GetDateDisplayText($CYear, $Month, $Day, $Status)
	Local $sDisplay = $Status
	If $sDisplay = "B" Then $sDisplay = "   "

	; Highlighted blank dates need a visible character. Otherwise the font color
	; changes correctly, but the user cannot see it because the cell has no text.
	If _IsHighlightedDate($CYear, $Month, $Day) Then
		If $sDisplay = "" Or $sDisplay = " " Or $sDisplay = "   " Then Return "X"
	EndIf

	Return $sDisplay
EndFunc   ;==>_GetDateDisplayText


Func _Button_HighlightDate($Month = "-1", $Day = "-1", $CYear = "-1")
	Local $SelDate
	If $Month = "-1" Then
		$SelDate = GUICtrlRead($Calendar)
	Else
		$SelDate = $CYear & "/" & $Month & "/" & $Day
	EndIf

	Local $aDate = StringSplit($SelDate, "/")
	If @error Or $aDate[0] <> 3 Then Return 0
	If Not _IsValidCalendarDay($aDate[1], $aDate[2], $aDate[3]) Then Return 0

	Local $sRegName = _DateHighlightRegName($aDate[1], $aDate[2], $aDate[3])
	If _IsHighlightedDate($aDate[1], $aDate[2], $aDate[3]) Then
		RegDelete($DB, $sRegName)
	Else
		RegWrite($DB, $sRegName, "REG_SZ", "1")
	EndIf

	_Update($SelDate)
	If $g_hLV <> 0 Then _CleanRepaint($g_hLV)
	_CustomCal_Update()

	Return 1
EndFunc   ;==>_Button_HighlightDate


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
		$Mouse_Tag_Pos_Y_New = ($Mouse_Tag_Pos_Y_New - $Mouse_Tag_Pos_Y_New_calc)
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

	Local $Form_Tag = GUICreate("Add/Edit Tag", 249, 181, $Mouse_Tag_Pos_X_New, $Mouse_Tag_Pos_Y_New, BitOR($WS_BORDER, $WS_POPUP, $DS_SETFOREGROUND, $DS_MODALFRAME), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW), $Form_WorkDays)
	Local $Label_Tag = GUICtrlCreateLabel("Selected Date (YYYY/MM/DD): " & $CYear & "/" & $Month & "/" & $Day, 8, 10, 250, 15)
	Local $Button_Tag_Cancel = GUICtrlCreateButton("Cancel", 8, 150, 75, 25)
	Local $Edit_Tag = GUICtrlCreateEdit("", 8, 38, 233, 105, BitOR($ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL, $ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_NOHIDESEL))
	Local $Button_Tag_Save = GUICtrlCreateButton("Save", 165, 150, 75, 25, $BS_DEFPUSHBUTTON)
	#forceref $Label_Tag

	GUICtrlSetData($Edit_Tag, $RegReadTag)
	GUISetState(@SW_SHOW, $Form_Tag)
	WinActivate($Form_Tag)

	While 1
		Local $aMsg = GUIGetMsg(1)
		If Not IsArray($aMsg) Then ContinueLoop

		Switch $aMsg[1]
			Case $Form_Tag
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE, $Button_Tag_Cancel
						GUIDelete($Form_Tag)
						WinActivate($Form_WorkDays)
						Return 0

					Case $Button_Tag_Save
						Local $DateToTag = $CYear & "/" & $Month & "/" & $Day
						Local $aDateToTag = StringSplit($DateToTag, "/")
						Local $sHolidayName = GUICtrlRead($Edit_Tag)
						Local $Register = RegRead($DB & "\" & $aDateToTag[1] & "\" & $aDateToTag[2], $aDateToTag[3])
						If $Register = "" Then $Register = "B"
						RegWrite($DB & "\" & $aDateToTag[1] & "\" & $aDateToTag[2], $aDateToTag[3], "REG_SZ", StringLeft($Register, 1) & $sHolidayName)
						GUIDelete($Form_Tag)
						WinActivate($Form_WorkDays)
						Return 1
				EndSwitch
		EndSwitch
	WEnd

EndFunc   ;==>_Button_Tag


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


Func _UpdateSelectionHighlight($iDay, $iMonth)
	If $g_iSelDay > 0 And $g_iSelMonth > 0 Then
		If $g_iSelDay <= 31 And $g_iSelMonth <= 12 Then
			If $SelectLabel[$g_iSelDay][$g_iSelMonth] <> 0 Then GUICtrlSetState($SelectLabel[$g_iSelDay][$g_iSelMonth], $gui_hide)
		EndIf
	EndIf

	If $iDay > 0 And $iMonth > 0 Then
		If $iDay <= 31 And $iMonth <= 12 Then
			If $SelectLabel[$iDay][$iMonth] <> 0 Then GUICtrlSetState($SelectLabel[$iDay][$iMonth], $gui_show)
		EndIf
	EndIf

	$g_iSelDay = $iDay
	$g_iSelMonth = $iMonth
EndFunc   ;==>_UpdateSelectionHighlight


Func _SetInputQuarterFast($iMonth)
	Switch Number($iMonth)
		Case 1 To 3
			GUICtrlSetData($Input_Quarter, "Q1")
		Case 4 To 6
			GUICtrlSetData($Input_Quarter, "Q2")
		Case 7 To 9
			GUICtrlSetData($Input_Quarter, "Q3")
		Case 10 To 12
			GUICtrlSetData($Input_Quarter, "Q4")
	EndSwitch
EndFunc   ;==>_SetInputQuarterFast


Func _CustomCal_GetCellIndex($iYear, $iMonth, $iDay)
	If $g_ccPrev = 0 Then Return -1
	If Number($iYear) <> Number($g_ccYear) Or Number($iMonth) <> Number($g_ccMonth) Then Return -1
	If $iDay < 1 Or $iDay > _DaysInMonth2($iYear, $iMonth) Then Return -1

	Local Const $iFirstDow = 6
	Local $iDay1MCM = Mod(_DateToDayOfWeek($iYear, $iMonth, 1) - 2 + 7, 7)
	Local $iDay1Col = Mod($iDay1MCM - $iFirstDow + 7, 7)
	Local $iOff = $iDay1Col + $iDay - 1
	Local $iRow = Int($iOff / 7)
	If $iRow < 0 Or $iRow > 5 Then Return -1
	Return $iRow * 7 + Mod($iOff, 7)
EndFunc   ;==>_CustomCal_GetCellIndex


Func _CustomCal_ApplyDayVisual($iYear, $iMonth, $iDay, $bSelected)
	Local $iIdx = _CustomCal_GetCellIndex($iYear, $iMonth, $iDay)
	If $iIdx < 0 Then Return 0

	Local $iBG = 0xFFFFFF
	Local $iFG = 0x000000
	If Number($iYear) = Number($g_iLVYear) Then
		$iBG = $g_aCellColor[$iMonth - 1][$iDay]
		$iFG = $g_aCellColorBK[$iMonth - 1][$iDay]
	EndIf

	Local $bHasNote = (Number($iYear) = Number($g_iLVYear) And $g_aCellStatus[$iMonth - 1][$iDay] <> "")
	Local $iFW = ($bHasNote) ? 700 : 400
	Local $iFrameColor = $iBG

	If Number($iYear) = @YEAR And Number($iMonth) = Number(@MON) And Number($iDay) = Number(@MDAY) Then
		$iFW = 700
		$iFrameColor = $Color_bk_Today
	EndIf

	If $bSelected Then
		$iFW = 700
		$iFrameColor = $Color_bk_Selected
	EndIf

	Local $iMarkerColor = ($bHasNote) ? $iFG : $iBG
	GUICtrlSetFont($g_ccDayCells[$iIdx], 9, $iFW, 0)
	GUICtrlSetBkColor($g_ccFrames[$iIdx], $iFrameColor)
	GUICtrlSetBkColor($g_ccMarkers[$iIdx], $iMarkerColor)
	_CustomCal_RedrawCtrl($g_ccFrames[$iIdx])
	_CustomCal_RedrawCtrl($g_ccDayCells[$iIdx])
	_CustomCal_RedrawCtrl($g_ccMarkers[$iIdx])
	Return 1
EndFunc   ;==>_CustomCal_ApplyDayVisual


Func _CustomCal_SelectDateFast($iOldYear, $iOldMonth, $iOldDay, $iNewYear, $iNewMonth, $iNewDay)
	; If the calendar must show a different month/year, fall back to the full redraw.
	If Number($g_ccYear) <> Number($iNewYear) Or Number($g_ccMonth) <> Number($iNewMonth) Then
		$g_ccYear = $iNewYear
		$g_ccMonth = $iNewMonth
		_CustomCal_Update()
		Return 1
	EndIf

	If $iOldYear = $iNewYear And $iOldMonth = $iNewMonth And $iOldDay = $iNewDay Then
		Return 1
	EndIf

	_CustomCal_ApplyDayVisual($iOldYear, $iOldMonth, $iOldDay, False)
	_CustomCal_ApplyDayVisual($iNewYear, $iNewMonth, $iNewDay, True)
	Return 1
EndFunc   ;==>_CustomCal_SelectDateFast


Func _RefreshSelectedDateUI($sSelDate)
	Local $aOld = StringSplit(GUICtrlRead($Input_SelDate), "/")
	Local $iOldYear = 0, $iOldMonth = 0, $iOldDay = 0
	If Not @error And IsArray($aOld) And $aOld[0] = 3 Then
		$iOldYear = Number($aOld[1])
		$iOldMonth = Number($aOld[2])
		$iOldDay = Number($aOld[3])
	EndIf

	Local $aDate = StringSplit($sSelDate, "/")
	If @error Or $aDate[0] <> 3 Then Return SetError(1, 0, 0)

	Local $iDataYear = Number($aDate[1])
	Local $iDataMonth = Number($aDate[2])
	Local $iDataDay = Number($aDate[3])
	If Not _IsValidCalendarDay($iDataYear, $iDataMonth, $iDataDay) Then Return SetError(2, 0, 0)
	Local $sDataMonth = StringFormat("%02d", $iDataMonth)
	Local $sDataDay = StringFormat("%02d", $iDataDay)

	Local $sDataRegister1 = RegRead($DB & "\" & $iDataYear & "\" & $sDataMonth, $sDataDay)
	If @error Then $sDataRegister1 = ""

	Local $sTip = ""
	If StringLen($sDataRegister1) > 1 Then
		$sTip = StringTrimLeft($sDataRegister1, 1)
	EndIf

	GUICtrlSetData($Input_SelDate, $iDataYear & "/" & $sDataMonth & "/" & $sDataDay)
	GUICtrlSetData($Input_Tag, $sTip)
	_GUICtrlMonthCal_SetCurSel($Calendar, $iDataYear, $iDataMonth, $iDataDay)

	_UpdateSelectionHighlight($iDataDay, $iDataMonth)
	_SetInputQuarterFast($iDataMonth)
	_CustomCal_SelectDateFast($iOldYear, $iOldMonth, $iOldDay, $iDataYear, $iDataMonth, $iDataDay)

	If $g_hLV <> 0 Then _WinAPI_InvalidateRect($g_hLV, 0, False)

	Return 1
EndFunc   ;==>_RefreshSelectedDateUI


Func _CalendarRead($i = 0, $j = 0)

	Local $SelDate = GUICtrlRead($Calendar)
	Local $SelDateYear = GUICtrlRead($Input_SelDate)
	Local $SelDate_slipt = StringSplit($SelDate, "/")
	Local $Input_SelDate_slipt = StringSplit($SelDateYear, "/")
	If @error Or $SelDate_slipt[0] <> 3 Then Return

	Local $iNewYear = Number($SelDate_slipt[1])
	Local $iNewMonth = Number($SelDate_slipt[2])
	Local $iNewDay = Number($SelDate_slipt[3])
	Local $iOldYear = 0
	If IsArray($Input_SelDate_slipt) And $Input_SelDate_slipt[0] = 3 Then $iOldYear = Number($Input_SelDate_slipt[1])

	; Fast path: same displayed year -> only refresh selection-dependent UI.
	If $g_iLVYear = $iNewYear And $iOldYear = $iNewYear Then
		_RefreshSelectedDateUI($SelDate)
		Return
	EndIf

	; Slow path: year changed -> rebuild the ListView once.
	_LockWindow($Form_WorkDays, False)

	$iYear = $iNewYear
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

	If $iNewYear <> $g_iLVYear Then
		_CriaINI($iNewYear)
	EndIf

	$g_ccYear = $iNewYear
	$g_ccMonth = $iNewMonth
	_Update($SelDate)
	_UpdateSelectionHighlight($iNewDay, $iNewMonth)

	Return

EndFunc   ;==>_CalendarRead


Func _Chart($Type = "", $bReset = False)



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

	If $bReset Then
		$Chart = ""
	ElseIf $Type <> "" Then
		If StringInStr($Chart, $Type) Then
			$Chart = StringReplace($Chart, $Type, "")
		Else
			$Chart = $Chart & $Type
		EndIf
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

	If $Count_O = 0 And StringInStr($Chart, "O") Then $Chart = StringReplace($Chart, "O", "")
	If $Count_R = 0 And StringInStr($Chart, "R") Then $Chart = StringReplace($Chart, "R", "")
	If $Count_H = 0 And StringInStr($Chart, "H") Then $Chart = StringReplace($Chart, "H", "")
	If $Count_P = 0 And StringInStr($Chart, "P") Then $Chart = StringReplace($Chart, "P", "")
	If $Count_T = 0 And StringInStr($Chart, "T") Then $Chart = StringReplace($Chart, "T", "")
	If $Count_S = 0 And StringInStr($Chart, "S") Then $Chart = StringReplace($Chart, "S", "")
	If $Count_B = 0 And StringInStr($Chart, "B") Then $Chart = StringReplace($Chart, "B", "")
	If $Count_W = 0 And StringInStr($Chart, "W") Then $Chart = StringReplace($Chart, "W", "")

	_ApplyMainGridCategoryFilter($Chart)

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


Func _CheckDate($DateToCheck, $NewStatus)

	$DateToCheck_split = StringSplit($DateToCheck, "/")
	If @error Or $DateToCheck_split[0] <> 3 Then Return 1
	If Not _IsValidCalendarDay($DateToCheck_split[1], $DateToCheck_split[2], $DateToCheck_split[3]) Then Return 1

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


Func _CheckDateReturn($DateToCheck)

	$DateToCheck_split = StringSplit($DateToCheck, "/")
	If @error Or $DateToCheck_split[0] <> 3 Then Return ""
	If Not _IsValidCalendarDay($DateToCheck_split[1], $DateToCheck_split[2], $DateToCheck_split[3]) Then Return ""

	$DateToCheck_Value = RegRead($DB & "\" & $DateToCheck_split[1] & "\" & $DateToCheck_split[2], $DateToCheck_split[3])

	$DateToCheck_Value = StringLeft($DateToCheck_Value, 1)

	Return $DateToCheck_Value

EndFunc   ;==>_CheckDateReturn


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

	Return $Color_bk_Blank
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

	Return $Font_Blank
EndFunc   ;==>_ColorFromDateFont


Func _CreateBackup($DBBKP = "")
	Local $sFilePath = ""

	If $DBBKP = "" Then
		DirCreate(@ScriptDir & "\Backup")
		$sFilePath = FileSaveDialog("Save backup file", @ScriptDir & "\Backup", "All (*.*)", 18, "Backup_" & @YEAR & "_" & @MON & "_" & @MDAY & ".bkp", $Form_WorkDays)
		If @error Then Return
	Else
		$sFilePath = $DBBKP
	EndIf

	Local $sCreated = _WD_Backup_Create($DB, $sFilePath, @ScriptDir & "\Backup", "Backup")
	If @error Or $sCreated = "" Then
		If $DBBKP = "" Then MsgBox(BitOR($MB_ICONERROR, $MB_TOPMOST), "Backup", "Unable to create the backup file." & @CRLF & @CRLF & "Target: " & $sFilePath, 0, $Form_WorkDays)
		Return SetError(1, 0, 0)
	EndIf

	If $DBBKP = "" Then
		MsgBox(BitOR(64, $MB_TOPMOST), "Sucess", "Backup saved: " & $sCreated, 0, $Form_WorkDays)
	EndIf

	Return $sCreated
EndFunc   ;==>_CreateBackup


Func _CreateMenu()

	GUICtrlDelete($DBpMenu_Report_Simple)
	GUICtrlDelete($DBpMenu_Report_Detailed)
	GUICtrlDelete($DBpMenu_Report_Professional)
	GUICtrlDelete($DBpMenu_Delete)

	Global $DBpMenu_Delete = GUICtrlCreateMenu("Delete Specific year", $BkpMenu_reset_all1)
	Global $DBpMenu_Report_Simple = GUICtrlCreateMenu("Simple", $DBpMenu_Report)
	Global $DBpMenu_Report_Detailed = GUICtrlCreateMenu("Detailed", $DBpMenu_Report)
	Global $DBpMenu_Report_Professional = GUICtrlCreateMenu("Analytical", $DBpMenu_Report)

	Local $sSubKey = ""
	For $i = 1 To 99

		$sSubKey = RegEnumKey($DB, $i)
		If @error Then ExitLoop

		If $sSubKey <> "OutlookAgent" Then
			$DBpMenu_Delete_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Delete)
			$DBpMenu_Report_simple_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Report_Simple)
			$DBpMenu_Report_detailed_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Report_Detailed)
			$DBpMenu_Report_professional_Year[$i] = GUICtrlCreateMenuItem($sSubKey, $DBpMenu_Report_Professional)
		EndIf

;~ 		ConsoleWrite("$i: " & $i & @CRLF)

	Next



	Return

EndFunc   ;==>_CreateMenu


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


Func _IsValidCalendarDay($iY, $iM, $iD)
	$iY = Number($iY)
	$iM = Number($iM)
	$iD = Number($iD)

	If $iY < 1 Then Return False
	If $iM < 1 Or $iM > 12 Then Return False
	If $iD < 1 Then Return False
	If $iD > _DaysInMonth2($iY, $iM) Then Return False

	Return True
EndFunc   ;==>_IsValidCalendarDay


Func _SetInvalidMainGridCell($iMonth, $iDay)
	$iMonth = Number($iMonth)
	$iDay = Number($iDay)

	If $iMonth < 1 Or $iMonth > 12 Then Return 0
	If $iDay < 1 Or $iDay > 31 Then Return 0

	If $g_hLV <> 0 And $iItem[$iMonth][0] <> 0 Then _GUICtrlListView_SetItemText($g_hLV, $iItem[$iMonth][0], "", $iDay)

	$g_aCellColor[$iMonth - 1][$iDay] = $g_clrInvalidDayBG
	$g_aCellColorBK[$iMonth - 1][$iDay] = $g_clrInvalidDayFG
	$g_aCellStatus[$iMonth - 1][$iDay] = ""
	$g_aCellTip[$iMonth - 1][$iDay] = ""

	Return 1
EndFunc   ;==>_SetInvalidMainGridCell


Func _NormalizeMainGridFilterStatus($sStatus)
	$sStatus = StringLeft($sStatus, 1)
	If $sStatus = "" Or $sStatus = "B" Or $sStatus = " " Then Return "B"
	Return $sStatus
EndFunc   ;==>_NormalizeMainGridFilterStatus


Func _MainGridStatusName($sStatus)
	Switch _NormalizeMainGridFilterStatus($sStatus)
		Case "W"
			Return "WEEKEND"
		Case "O"
			Return "ON-SITE"
		Case "R"
			Return "REMOTE"
		Case "T"
			Return "TRAVEL"
		Case "P"
			Return "PTO"
		Case "H"
			Return "HOLIDAY"
		Case "S"
			Return "SICK DAY"
		Case Else
			Return "BLANK"
	EndSwitch
EndFunc   ;==>_MainGridStatusName


Func _ApplyMainGridCategoryFilter($sFilter = "")
	$g_sMainGridFilter = $sFilter
	If $g_hLV = 0 Then Return 0

	Local $iTargetYear = $iYear

	For $iMonth = 1 To 12
		Local $sMonth = StringFormat("%02d", $iMonth)
		Local $iDaysInMonth = _DaysInMonth2($iTargetYear, $iMonth)

		For $iDay = 1 To 31
			If $iDay > $iDaysInMonth Then
				_SetInvalidMainGridCell($iMonth, $iDay)
				ContinueLoop
			EndIf

			Local $sDay = StringFormat("%02d", $iDay)
			Local $sRawValue = RegRead($DB & "\" & $iTargetYear & "\" & $sMonth, $sDay)
			If @error Then $sRawValue = ""

			Local $sStatus = StringLeft($sRawValue, 1)
			Local $sFilterStatus = _NormalizeMainGridFilterStatus($sStatus)

			If $sFilter <> "" And Not StringInStr($sFilter, $sFilterStatus) Then
				_GUICtrlListView_SetItemText($g_hLV, $iItem[$iMonth][0], "", $iDay)
				$g_aCellColor[$iMonth - 1][$iDay] = 0xFFFFFF
				$g_aCellColorBK[$iMonth - 1][$iDay] = 0x000000
				$g_aCellStatus[$iMonth - 1][$iDay] = ""
				$g_aCellTip[$iMonth - 1][$iDay] = ""
				ContinueLoop
			EndIf

			Local $sDisplay = _GetDateDisplayText($iTargetYear, $iMonth, $iDay, $sStatus)
			If $sFilter <> "" And $sFilterStatus = "B" Then $sDisplay = "B"
			_GUICtrlListView_SetItemText($g_hLV, $iItem[$iMonth][0], $sDisplay, $iDay)

			Local $sComment = ""
			If StringLen($sRawValue) > 1 Then $sComment = StringTrimLeft($sRawValue, 1)

			Local $iWeekDayNum = _DateToDayOfWeek($iTargetYear, $iMonth, $iDay)
			Local $sWeekDayName = _DateDayOfWeek($iWeekDayNum, 1)
			Local $iWeekDayNumber = _WeekNumberISO($iTargetYear, $iMonth, $iDay)
			Local $sStatusName = _MainGridStatusName($sStatus)
			Local $sTip = ""

			If $sComment <> "" Then
				$sTip = $iTargetYear & "/" & $sMonth & "/" & $sDay & @CRLF & _
						$sWeekDayName & " (Week: " & $iWeekDayNumber & ") - " & $sStatusName & @CRLF & _
						"----" & @CRLF & "- " & StringReplace($sComment, @CRLF, @CRLF & "- ")
			Else
				$sTip = $iTargetYear & "/" & $sMonth & "/" & $sDay & @CRLF & _
						$sWeekDayName & " (Week: " & $iWeekDayNumber & ") - " & $sStatusName
			EndIf

			$g_aCellColor[$iMonth - 1][$iDay] = _ColorFromDate($sStatus)
			$g_aCellColorBK[$iMonth - 1][$iDay] = _GetDateFontColor($iTargetYear, $iMonth, $iDay, $sStatus)
			$g_aCellStatus[$iMonth - 1][$iDay] = $sComment
			$g_aCellTip[$iMonth - 1][$iDay] = $sTip
		Next
	Next

	_HideListViewCellTip()
	If $g_hLV <> 0 Then _CleanRepaint($g_hLV)
	_CustomCal_Update()

	Return 1
EndFunc   ;==>_ApplyMainGridCategoryFilter


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


Func _DecColorToRGBHex($nColor)
	Local $r = BitAND($nColor, 0xFF)
	Local $g = BitAND(BitShift($nColor, 8), 0xFF)
	Local $b = BitAND(BitShift($nColor, 16), 0xFF)
	Return StringFormat("0x%02X%02X%02X", $r, $g, $b)
EndFunc   ;==>_DecColorToRGBHex


Func _DestroyPopupFail($hMenu, $iErr)
	DllCall("user32.dll", "bool", "DestroyMenu", "handle", $hMenu)
	Return SetError($iErr, 0, 0)
EndFunc   ;==>_DestroyPopupFail


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


; ════════════════════════════════════════════════════════════════════════════
; Custom Calendar – creates a colored label-grid that replaces the MonthCal
; visually.  GUICtrlSetBkColor on labels is 100 % reliable – no GDI/custom-
; draw dependency at all.
;
; Layout  (fits exactly in the original MonthCal slot: x=8, y=8, 273x201)
;   Row 0  (h=20) : [ < ]  [ Month Year (centred) ]  [ > ]
;   Row 1  (h=18) : Wk | Mon Tue Wed Thu Fri Sat Sun
;   Rows 2-7 (h=27 each, 6 rows) : week-number + 7 day cells
;   8 columns x 34 px = 272 px  (≈ 273)
; ════════════════════════════════════════════════════════════════════════════
Func _CustomCal_Create()
	Local Const $X = 8      ; left edge (same as hidden MonthCal)
	Local Const $CW = 34    ; column width  (8 cols x 34 = 272)
	Local Const $TH = 20    ; title row height
	Local Const $DNH = 16   ; day-names row height  (reduced: 18→16 to fit Today btn)
	Local Const $RH = 24    ; week-row height        (reduced: 27→24 to fit Today btn)
	; Layout summary:
	;   Title     y=8..28   (h=20)
	;   Day names y=28..44  (h=16)
	;   Grid      y=44..188 (6 rows × 24 = 144)
	;   Today btn y=190..207 (h=17)  ← safely above ListView which starts at y=210

	Local $Y0 = 8            ; title row top
	Local $Y1 = $Y0 + $TH    ; day-names row top  (= 28)
	Local $Y2 = $Y1 + $DNH   ; first week row top (= 46)

	; -- title row: [<]  [Month Year – click to pick]  [>] ------------
	$g_ccPrev = GUICtrlCreateButton("<", $X, $Y0, $CW, $TH)
	GUICtrlSetFont($g_ccPrev, 9, 700)
	GUICtrlSetTip($g_ccPrev, "Previous month")

	; Title is now a Button so it reacts to clicks for the month picker
	$g_ccTitle = GUICtrlCreateButton("", $X + $CW, $Y0, 6 * $CW, $TH)
	GUICtrlSetFont($g_ccTitle, 9, 700)
	GUICtrlSetBkColor($g_ccTitle, 0xDDE5F0)
	GUICtrlSetTip($g_ccTitle, "Click to jump to a specific month / year")

	$g_ccNext = GUICtrlCreateButton(">", $X + 7 * $CW, $Y0, $CW, $TH)
	GUICtrlSetFont($g_ccNext, 9, 700)
	GUICtrlSetTip($g_ccNext, "Next month")

	; -- day-names row  (first day = Sunday, hardcoded per user request)
	; MCM notation: 6 = Sunday, so columns go Su Mo Tu We Th Fr Sa
	Local Const $iFirstDow = 6   ; 6 = Sunday in MCM (0=Mon..6=Sun)
	GUICtrlCreateLabel("Wk", $X, $Y1, $CW, $DNH, $SS_CENTER)
	GUICtrlSetFont(-1, 7, 700)
	GUICtrlSetBkColor(-1, 0xC8D2E4)
	GUICtrlSetColor(-1, 0x404050)

	; Abbreviated names in MCM order (0=Mon..6=Sun)
	Local $aAbbr[7] = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
	For $k = 0 To 6
		Local $iDI = Mod($iFirstDow + $k, 7)   ; 0=Mon..6=Sun
		; highlight weekend cols: iDI 5=Sat, 6=Sun
		Local $iBGDN = ($iDI >= 5) ? 0xB8C4DC : 0xC8D2E4
		GUICtrlCreateLabel($aAbbr[$iDI], $X + ($k + 1) * $CW, $Y1, $CW, $DNH, $SS_CENTER)
		GUICtrlSetFont(-1, 8, 700)
		GUICtrlSetBkColor(-1, $iBGDN)
		GUICtrlSetColor(-1, 0x303050)
	Next

	; -- day cells + week-number cells (6 rows x 8 cols) --------------
	; Each day slot has two stacked labels:
	;   $g_ccFrames[i]   – full cell size, created FIRST (behind)
	;   $g_ccDayCells[i] – 1px inset on every side, created SECOND (on top)
	; The 1px rim of the frame label is always visible, and its background
	; is set to $Color_bk_Today or $Color_bk_Selected when appropriate.
	For $r = 0 To 5
		$g_ccWeekNums[$r] = GUICtrlCreateLabel("", $X, $Y2 + $r * $RH, $CW, $RH, $SS_CENTER)
		GUICtrlSetFont($g_ccWeekNums[$r], 7, 400)
		GUICtrlSetBkColor($g_ccWeekNums[$r], 0xC8D2E4)
		GUICtrlSetColor($g_ccWeekNums[$r], 0x505068)
		$g_ccCacheWeekNums[$r] = Chr(0)

		For $c = 0 To 6
			Local $iIdx = $r * 7 + $c
			Local $iCX = $X + ($c + 1) * $CW
			Local $iCY = $Y2 + $r * $RH

			; Frame label – created first so content label draws on top of it
			$g_ccFrames[$iIdx] = GUICtrlCreateLabel("", $iCX, $iCY, $CW, $RH)
			GUICtrlSetBkColor($g_ccFrames[$iIdx], 0xEEEEEE)

			; Content label – 1px inset so frame rim shows as colored border
			$g_ccDayCells[$iIdx] = GUICtrlCreateLabel("", _
					$iCX + 1, $iCY + 1, $CW - 2, $RH - 2, $SS_CENTER)
			GUICtrlSetFont($g_ccDayCells[$iIdx], 9, 400)
			GUICtrlSetBkColor($g_ccDayCells[$iIdx], 0xEEEEEE)
			$g_ccDayValues[$iIdx] = 0

			; Marker bar – 2px tall, at the very bottom of the content area.
			; Shown (colored) when the day has a note; hidden (matches bg) otherwise.
;~ 			$g_ccMarkers[$iIdx] = GUICtrlCreateLabel("", $iCX + 2, $iCY + $RH - 3, $CW - 4, 1)
			$g_ccMarkers[$iIdx] = GUICtrlCreateLabel("", $iCX + 10, $iCY + $RH - 7, $CW - 20, 1)
			GUICtrlSetBkColor($g_ccMarkers[$iIdx], 0xEEEEEE)
			$g_ccCacheText[$iIdx] = Chr(0)
			$g_ccCacheDayBG[$iIdx] = -1
			$g_ccCacheDayFG[$iIdx] = -1
			$g_ccCacheFontW[$iIdx] = -1
			$g_ccCacheFrame[$iIdx] = -1
			$g_ccCacheMarker[$iIdx] = -1
			$g_ccCacheTip[$iIdx] = Chr(0)
		Next
	Next

	; -- "Today" button at the bottom (full width, 17px tall) ---------
	; Y2 + 6*RH + 2gap = 44 + 144 + 2 = 190  →  ends at 207  (ListView is at 210)
	Local $iYToday = $Y2 + 6 * $RH + 2
;~ 	$g_ccToday = GUICtrlCreateButton("Today  ( " & @YEAR & "/" & @MON & "/" & @MDAY & " )", $X, $iYToday, 8 * $CW, 17)
	$g_ccToday = GUICtrlCreateButton("Today  ( " & @YEAR & "/" & @MON & "/" & @MDAY & " )", $X, $iYToday, 8 * $CW, 20)
	GUICtrlSetFont($g_ccToday, 8, 700)
	GUICtrlSetBkColor($g_ccToday, 0xDDE5F0)
	GUICtrlSetTip($g_ccToday, "Go to today: " & @YEAR & "/" & @MON & "/" & @MDAY)
EndFunc   ;==>_CustomCal_Create


Func _CustomCal_Update()
	If $g_ccPrev = 0 Then Return   ; not yet created

	Local $iDYear = $g_ccYear
	Local $iDMonth = $g_ccMonth
	Local $iDays = _DaysInMonth2($iDYear, $iDMonth)

	; -- title --------------------------------------------------------
	Local $aMon[12] = ["January", "February", "March", "April", "May", "June", _
			"July", "August", "September", "October", "November", "December"]
	Local $sTitle = $aMon[$iDMonth - 1] & "  " & $iDYear
	If $g_ccCacheTitle <> $sTitle Then
		GUICtrlSetData($g_ccTitle, $sTitle)
		$g_ccCacheTitle = $sTitle
	EndIf

	; -- first day of week: Sunday (hardcoded per user request)
	; MCM notation: 6 = Sunday (0=Mon..6=Sun)
	Local Const $iFirstDow = 6

	; -- column of day 1 (0-based, relative to first-day-of-week col) -
	; AutoIt _DateToDayOfWeek: 1=Sun..7=Sat -> MCM 0-based (0=Mon..6=Sun):
	;   (autoit - 2 + 7) mod 7
	Local $iDay1MCM = Mod(_DateToDayOfWeek($iDYear, $iDMonth, 1) - 2 + 7, 7)
	Local $iDay1Col = Mod($iDay1MCM - $iFirstDow + 7, 7)

	; -- selected date ------------------------------------------------
	Local $iSelY = 0, $iSelM = 0, $iSelD = 0
	Local $aSel = StringSplit(GUICtrlRead($Input_SelDate), "/")
	If Not @error And $aSel[0] = 3 Then
		$iSelY = Number($aSel[1])
		$iSelM = Number($aSel[2])
		$iSelD = Number($aSel[3])
	EndIf

	; Build desired state for each visible slot, then update only controls that changed.
	Local $aDayVal[42], $aText[42], $aDayBG[42], $aDayFG[42], $aFontW[42], $aFrame[42], $aMarker[42], $aTip[42]
	For $i = 0 To 41
		$aDayVal[$i] = 0
		$aText[$i] = ""
		$aDayBG[$i] = 0xEEEEEE
		$aDayFG[$i] = 0xBBBBBB
		$aFontW[$i] = 400
		$aFrame[$i] = 0xEEEEEE
		$aMarker[$i] = 0xEEEEEE
		$aTip[$i] = ""
	Next

	For $d = 1 To $iDays
		Local $iOff = $iDay1Col + $d - 1
		Local $iRow = Int($iOff / 7)
		Local $iCol = Mod($iOff, 7)
		If $iRow > 5 Then ContinueLoop

		Local $iIdx = $iRow * 7 + $iCol
		$aDayVal[$iIdx] = $d
		$aText[$iIdx] = String($d)

		Local $iBG = 0xFFFFFF
		Local $iFG = 0x000000
		If $iDYear = $iYear Then
			$iBG = $g_aCellColor[$iDMonth - 1][$d]
			$iFG = $g_aCellColorBK[$iDMonth - 1][$d]
		EndIf

		Local $bHasNote = ($iDYear = $iYear And $g_aCellStatus[$iDMonth - 1][$d] <> "")
		Local $iFW = 400
		If $bHasNote Then $iFW = 700

		Local $iFrameColor = $iBG
		If $iDYear = @YEAR And $iDMonth = Number(@MON) And $d = Number(@MDAY) Then
			$iFW = 700
			$iFrameColor = $Color_bk_Today
		EndIf

		If $iDYear = $iSelY And $iDMonth = $iSelM And $d = $iSelD Then
			$iFW = 700
			$iFrameColor = $Color_bk_Selected
		EndIf

		Local $iMarkerColor = ($bHasNote) ? $iFG : $iBG
		Local $sTip = ""
		If $iDYear = $iYear Then $sTip = $g_aCellTip[$iDMonth - 1][$d]

		$aDayBG[$iIdx] = $iBG
		$aDayFG[$iIdx] = $iFG
		$aFontW[$iIdx] = $iFW
		$aFrame[$iIdx] = $iFrameColor
		$aMarker[$iIdx] = $iMarkerColor
		$aTip[$iIdx] = $sTip
	Next

	; Apply only deltas. Avoid full-form lock/redraw here because that forces a heavy repaint
	; of the ListView and is slower than updating just the calendar controls.
	For $i = 0 To 41
		If $g_ccDayValues[$i] <> $aDayVal[$i] Or $g_ccCacheText[$i] <> $aText[$i] Then
			GUICtrlSetData($g_ccDayCells[$i], $aText[$i])
			$g_ccDayValues[$i] = $aDayVal[$i]
			$g_ccCacheText[$i] = $aText[$i]
		EndIf

		If $g_ccCacheDayBG[$i] <> $aDayBG[$i] Then
			GUICtrlSetBkColor($g_ccDayCells[$i], $aDayBG[$i])
			$g_ccCacheDayBG[$i] = $aDayBG[$i]
		EndIf

		If $g_ccCacheDayFG[$i] <> $aDayFG[$i] Then
			GUICtrlSetColor($g_ccDayCells[$i], $aDayFG[$i])
			$g_ccCacheDayFG[$i] = $aDayFG[$i]
		EndIf

		If $g_ccCacheFontW[$i] <> $aFontW[$i] Then
			GUICtrlSetFont($g_ccDayCells[$i], 9, $aFontW[$i], 0)
			$g_ccCacheFontW[$i] = $aFontW[$i]
		EndIf

		If $g_ccCacheFrame[$i] <> $aFrame[$i] Then
			GUICtrlSetBkColor($g_ccFrames[$i], $aFrame[$i])
			$g_ccCacheFrame[$i] = $aFrame[$i]
		EndIf

		If $g_ccCacheMarker[$i] <> $aMarker[$i] Then
			GUICtrlSetBkColor($g_ccMarkers[$i], $aMarker[$i])
			$g_ccCacheMarker[$i] = $aMarker[$i]
		EndIf

		; Native tooltips on all hit-testable parts of the small calendar.
		; This is safer than a manual hover tooltip here and avoids the click/freeze issue.
		If $g_ccCacheTip[$i] <> $aTip[$i] Then
			GUICtrlSetTip($g_ccDayCells[$i], $aTip[$i])
			GUICtrlSetTip($g_ccFrames[$i], $aTip[$i])
			GUICtrlSetTip($g_ccMarkers[$i], $aTip[$i])
			$g_ccCacheTip[$i] = $aTip[$i]
		EndIf
	Next

	; -- week numbers (left column) -----------------------------------
	For $r = 0 To 5
		Local $sWeek = ""
		Local $iFirst = $r * 7 - $iDay1Col + 1
		Local $iLast = ($r + 1) * 7 - $iDay1Col
		If Not ($iLast < 1 Or $iFirst > $iDays) Then
			Local $iWD = ($iFirst < 1) ? 1 : $iFirst
			$sWeek = _WeekNumberISO($iDYear, $iDMonth, $iWD)
		EndIf

		If $g_ccCacheWeekNums[$r] <> $sWeek Then
			GUICtrlSetData($g_ccWeekNums[$r], $sWeek)
			$g_ccCacheWeekNums[$r] = $sWeek
		EndIf
	Next

	; Background-color changes on AutoIt labels sometimes need an explicit repaint.
	; Repaint only the custom calendar controls to keep month navigation responsive.
	_CustomCal_Repaint()
EndFunc   ;==>_CustomCal_Update


Func _CustomCal_RedrawCtrl($idCtrl)
	If $idCtrl = 0 Then Return
	Local $hCtrl = GUICtrlGetHandle($idCtrl)
	If $hCtrl = 0 Then Return
	DllCall("user32.dll", "bool", "RedrawWindow", _
			"hwnd", $hCtrl, _
			"ptr", 0, _
			"handle", 0, _
			"uint", 0x0101) ; RDW_INVALIDATE | RDW_UPDATENOW
EndFunc   ;==>_CustomCal_RedrawCtrl


Func _CustomCal_Repaint()
	If $g_ccPrev = 0 Then Return
	_CustomCal_RedrawCtrl($g_ccPrev)
	_CustomCal_RedrawCtrl($g_ccTitle)
	_CustomCal_RedrawCtrl($g_ccNext)
	_CustomCal_RedrawCtrl($g_ccToday)
	For $r = 0 To 5
		_CustomCal_RedrawCtrl($g_ccWeekNums[$r])
	Next
	For $i = 0 To 41
		_CustomCal_RedrawCtrl($g_ccFrames[$i])
		_CustomCal_RedrawCtrl($g_ccDayCells[$i])
		_CustomCal_RedrawCtrl($g_ccMarkers[$i])
	Next
EndFunc   ;==>_CustomCal_Repaint


Func _CustomCal_Navigate($iDelta)
	; Always navigate by selecting the first day of the target month.
	; Important: calculate the target date without pre-setting $g_ccYear/$g_ccMonth.
	; _CalendarRead() / _CustomCal_SelectDateFast() use the current displayed month
	; to decide whether a full custom-calendar repaint is needed. If we update the
	; globals here first, the UI can think it is already on the new month/year and
	; skip the redraw, which is what causes colors to appear only after selecting a day.
	Local $iNewYear = Number($g_ccYear)
	Local $iNewMonth = Number($g_ccMonth) + Number($iDelta)

	If $iNewMonth > 12 Then
		$iNewMonth = 1
		$iNewYear += 1
	EndIf
	If $iNewMonth < 1 Then
		$iNewMonth = 12
		$iNewYear -= 1
	EndIf

	Local $sNewDate = $iNewYear & "/" & StringFormat("%02d", $iNewMonth) & "/01"

	; Sync the hidden MonthCal and then let the normal selection pipeline update
	; Input_SelDate, quarter, ListView year data, selected highlight, tips, and colors.
	GUICtrlSetData($Calendar, $sNewDate)
	_GUICtrlMonthCal_SetCurSel($Calendar, $iNewYear, $iNewMonth, 1)
	_CalendarRead()
EndFunc   ;==>_CustomCal_Navigate


; Opens a small popup MonthCal so the user can jump to any month/year quickly.
; Positioned just below the title button.  When the user selects a date, the
; custom calendar navigates to that month and the popup closes automatically.
Func _CustomCal_ShowPicker()
	; Position popup under the title button
	; Title button: screen coords = window pos + (8 + 34, 8) = (42, 8) client
	Local $aWinPos = WinGetPos($Form_WorkDays)
	If @error Or Not IsArray($aWinPos) Then Return

	; Approximate title button screen position
	; Client area starts ~30px from window top (title bar), 4px from left (border)
	Local $iBorderX = 4
	Local $iTitleBarH = 30
	Local $iTitleBtnX = $aWinPos[0] + $iBorderX + 8 + 34          ; = X + CW
	Local $iTitleBtnY = $aWinPos[1] + $iTitleBarH + 8 + 20 + 2    ; below title row

	Local $hPicker = GUICreate("", 220, 165, $iTitleBtnX, $iTitleBtnY, _
			BitOR($WS_POPUP, $WS_BORDER), $WS_EX_TOPMOST, $Form_WorkDays)

	; MonthCal in the popup – no week numbers to keep it compact
	Local $sPickerDate = $g_ccYear & "/" & StringFormat("%02d", $g_ccMonth) & "/01"
	Local $hPickerCal = GUICtrlCreateMonthCal($sPickerDate, 2, 2, 215, 160)

	GUISetState(@SW_SHOW, $hPicker)
	GUISwitch($Form_WorkDays)   ; keep focus on main form

	; Sub-loop: wait for a date click or Escape / close
	While 1
		Local $nPick = GUIGetMsg(1)   ; 1 = check all GUIs
		If Not IsArray($nPick) Then ContinueLoop

		If $nPick[1] = $hPicker Then
			Select
				Case $nPick[0] = $GUI_EVENT_CLOSE
					GUIDelete($hPicker)
					Return
				Case $nPick[0] = $hPickerCal
					; User clicked a day - extract month/year only
					Local $sPicked = GUICtrlRead($hPickerCal)
					Local $aPicked = StringSplit($sPicked, "/")
					If Not @error And $aPicked[0] >= 2 Then
						$g_ccYear = Number($aPicked[1])
						$g_ccMonth = Number($aPicked[2])
						Local $sNewDate = $g_ccYear & "/" & StringFormat("%02d", $g_ccMonth) & "/01"
						GUICtrlSetData($Calendar, $sNewDate)
					EndIf
					GUIDelete($hPicker)

					If $g_ccYear <> $iYear Then
						; Year changed: explicitly initialise the registry for the
						; new year BEFORE _CalendarRead runs.  This is critical for
						; years that have never been opened before – without it,
						; _CriaINI would only be called if _CalendarRead's own
						; year-comparison fires, but that comparison uses $Input_SelDate
						; which may already reflect the new year in some code paths.
						_CriaINI($g_ccYear)

						; Do NOT pre-set $Input_SelDate here.  _CalendarRead compares
						; $Calendar (new year) vs $Input_SelDate (old year) to decide
						; whether to call _CriaINI.  Pre-setting it would make both
						; equal and skip the check.  Let _Update set it correctly.
						_CalendarRead()
					Else
						; Same year, just repaint the custom calendar widget.
						_CustomCal_Update()
					EndIf
					Return
			EndSelect
		EndIf

		; Close picker if user clicks outside it
		If _IsPressed("1B") Then   ; Escape
			GUIDelete($hPicker)
			Return
		EndIf
	WEnd
EndFunc   ;==>_CustomCal_ShowPicker


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


Func _HideListViewCellTip()
	If $g_bTipVisible Then ToolTip("")
	$g_bTipVisible = False
	$g_iTipRow = -1
	$g_iTipCol = -1
	$g_sTipText = ""
EndFunc   ;==>_HideListViewCellTip


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


Func _Interpolate($v1, $v2, $ratio)
	Return Round($v1 + ($v2 - $v1) * $ratio)
EndFunc   ;==>_Interpolate


Func _IsLeapYear($iY)
	If Mod($iY, 400) = 0 Then Return True
	If Mod($iY, 100) = 0 Then Return False
	Return (Mod($iY, 4) = 0)
EndFunc   ;==>_IsLeapYear


Func _Label($LabelName)

	If $LabelName = "" Then Return "Blank"
	If $LabelName = "O" Then Return "On Site""On Site"
	If $LabelName = "R" Then Return "Remote"
	If $LabelName = "H" Then Return "Holiday"
	If $LabelName = "P" Then Return "PTO"
	If $LabelName = "T" Then Return "Travel"
	If $LabelName = "S" Then Return "Sick"
	If $LabelName = "B" Then Return "Blank"
	If $LabelName = "W" Then Return "Weekend"

EndFunc   ;==>_Label


Func _Log($Inputs_log)
;~ 	$Log = FileOpen(@ScriptDir & "\Log.txt",9)
;~ 	FileWriteLine($Log,$Inputs_log)
;~ 	FileClose($Log)
	ConsoleWrite($Inputs_log & @CRLF)

EndFunc   ;==>_Log


Func _MenuContextual($U, $V, $SelYear)
	; U = day
	; V = month
	; Custom popup menu instead of the native Windows popup menu.
	; This allows each action to use the configured background/font colors.

	Local $sDay = StringFormat("%02d", Number($U))
	Local $sMonth = StringFormat("%02d", Number($V))

	Local Const $IDM_HIGHLIGHT = 1001
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

	Local Const $iMenuW = 190
	Local Const $iLineH = 24
	Local Const $iSepH = 6
	Local Const $iMenuH = 256

	Local $iX = $mousePosX
	Local $iY = $mousePosY

	; Keep the popup inside the work area of the monitor where the user clicked.
	; @DesktopWidth/@DesktopHeight only describe the primary desktop in some setups,
	; so clamping against them can push the menu back to another monitor.
	_ClampPopupToCurrentMonitor($iX, $iY, $iMenuW, $iMenuH, $hOwner)

	Local $hPopup = GUICreate("", $iMenuW, $iMenuH, $iX, $iY, BitOR($WS_POPUP, $WS_BORDER), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW), $hOwner)
	If $hPopup = 0 Then Return SetError(1, 0, 0)

	GUISetBkColor(0xF4F4F4, $hPopup)

	Local $iYPos = 2
	Local $idDate = GUICtrlCreateLabel("  Date: " & $sDay & "/" & $sMonth & "/" & $SelYear, 1, $iYPos, $iMenuW - 2, $iLineH, $SS_CENTERIMAGE)
	GUICtrlSetBkColor($idDate, 0xF0F0F0)
	GUICtrlSetColor($idDate, 0x666666)
	GUICtrlSetFont($idDate, 9, 700, 0, "Segoe UI")
	$iYPos += $iLineH

	_CreateColoredContextMenuSeparator($iYPos, $iMenuW)
	$iYPos += $iSepH

	Local $idHighlight = _CreateColoredContextMenuItem("Highlight date", $iYPos, $iMenuW, $iLineH, 0xFFFFFF, $Color_HighlightDate, True)
	$iYPos += $iLineH

	Local $idTag = _CreateColoredContextMenuItem("Add/Edit Marker", $iYPos, $iMenuW, $iLineH, 0xFFFFFF, 0x000000, False)
	$iYPos += $iLineH

	_CreateColoredContextMenuSeparator($iYPos, $iMenuW)
	$iYPos += $iSepH

	Local $idOnSite = _CreateColoredContextMenuItem("On-Site", $iYPos, $iMenuW, $iLineH, $Color_bk_OnSite, $Font_OnSite, True)
	$iYPos += $iLineH

	Local $idRemote = _CreateColoredContextMenuItem("Remote", $iYPos, $iMenuW, $iLineH, $Color_bk_Remote, $Font_Remote, True)
	$iYPos += $iLineH

	Local $idHoliday = _CreateColoredContextMenuItem("Holiday", $iYPos, $iMenuW, $iLineH, $Color_bk_holiday, $Font_Holiday, True)
	$iYPos += $iLineH

	Local $idPTO = _CreateColoredContextMenuItem("PTO", $iYPos, $iMenuW, $iLineH, $Color_bk_PTO, $Font_PTO, True)
	$iYPos += $iLineH

	Local $idTravel = _CreateColoredContextMenuItem("Travel", $iYPos, $iMenuW, $iLineH, $Color_bk_Travel, $Font_Travel, True)
	$iYPos += $iLineH

	Local $idSick = _CreateColoredContextMenuItem("Sick", $iYPos, $iMenuW, $iLineH, $Color_bk_Sick, $Font_Sick, True)
	$iYPos += $iLineH

	; The Blank button writes Weekend automatically when the selected date is Saturday/Sunday.
	; Show the resulting configured color in the menu for that specific date.
	Local $iBlankMenuBk = $Color_bk_Blank
	Local $iBlankMenuFont = $Font_Blank
	Local $iWeekDay = _DateToDayOfWeek(Number($SelYear), Number($sMonth), Number($sDay))
	If $iWeekDay = 1 Or $iWeekDay = 7 Then
		$iBlankMenuBk = $Color_bk_Weekend
		$iBlankMenuFont = $Font_Weekend
	EndIf

	Local $idBlank = _CreateColoredContextMenuItem("Blank / Weekends", $iYPos, $iMenuW, $iLineH, $iBlankMenuBk, $iBlankMenuFont, True)

	GUISetState(@SW_SHOW, $hPopup)
	WinActivate($hPopup)

	Local $iSelected = 0
	Local $bMouseReleased = False

	While 1
		Local $aMsg = GUIGetMsg(1)
		If IsArray($aMsg) Then
			If $aMsg[1] = $hPopup Then
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						$iSelected = 0
						ExitLoop
					Case $idHighlight
						$iSelected = $IDM_HIGHLIGHT
						ExitLoop
					Case $idTag
						$iSelected = $IDM_TAG
						ExitLoop
					Case $idOnSite
						$iSelected = $IDM_ONSITE
						ExitLoop
					Case $idRemote
						$iSelected = $IDM_REMOTE
						ExitLoop
					Case $idHoliday
						$iSelected = $IDM_HOLIDAY
						ExitLoop
					Case $idPTO
						$iSelected = $IDM_PTO
						ExitLoop
					Case $idTravel
						$iSelected = $IDM_TRAVEL
						ExitLoop
					Case $idSick
						$iSelected = $IDM_SICK
						ExitLoop
					Case $idBlank
						$iSelected = $IDM_BLANK
						ExitLoop
				EndSwitch
			ElseIf $bMouseReleased And $aMsg[0] <> 0 Then
				; Any click/control event on another window should dismiss the popup,
				; matching normal context-menu behavior.
				ExitLoop
			EndIf
		EndIf

		; Wait for the original right-click to be released before closing on outside clicks.
		If Not _IsPressed("01") And Not _IsPressed("02") Then $bMouseReleased = True

		If $bMouseReleased Then
			If _IsPressed("1B") Then ExitLoop ; ESC

			If _IsPressed("01") Or _IsPressed("02") Then
				If Not _MouseInsideRect($iX, $iY, $iMenuW, $iMenuH) Then ExitLoop
			EndIf

			; If the user clicks somewhere else, Windows moves focus away from the
			; popup. This catches fast clicks that can happen between polling ticks.
			If Not WinActive($hPopup) Then
				Local $hForeground = _GetForegroundWindowHandle()
				If $hForeground <> 0 And $hForeground <> $hPopup And Not _MouseInsideRect($iX, $iY, $iMenuW, $iMenuH) Then ExitLoop
			EndIf
		EndIf

		Sleep(10)
	WEnd

	GUIDelete($hPopup)
	WinActivate($hOwner)

	Switch $iSelected
		Case 0
			Return 0

		Case $IDM_HIGHLIGHT
			_Button_HighlightDate($sMonth, $sDay, $SelYear)

		Case $IDM_TAG
			If _Button_Tag($sMonth, $sDay, $SelYear) Then
				; Tag edits only affect the selected day tooltip/marker visuals.
				; A full _Reload() is unnecessary here and previously could freeze the UI.
				_Update($SelYear & "/" & $sMonth & "/" & $sDay)
			EndIf

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


Func _CreateColoredContextMenuItem($sText, $iY, $iW, $iH, $iBkColor, $iFontColor, $bBold = False)
	Local $idItem = GUICtrlCreateLabel("  " & $sText, 1, $iY, $iW - 2, $iH, BitOR($SS_NOTIFY, $SS_CENTERIMAGE))
	GUICtrlSetBkColor($idItem, $iBkColor)
	GUICtrlSetColor($idItem, $iFontColor)
	If $bBold Then
		GUICtrlSetFont($idItem, 9, 700, 0, "Segoe UI")
	Else
		GUICtrlSetFont($idItem, 9, 400, 0, "Segoe UI")
	EndIf
	Return $idItem
EndFunc   ;==>_CreateColoredContextMenuItem


Func _CreateColoredContextMenuSeparator($iY, $iW)
	Local $idSep = GUICtrlCreateLabel("", 4, $iY + 2, $iW - 8, 1)
	GUICtrlSetBkColor($idSep, 0xD0D0D0)
	Return $idSep
EndFunc   ;==>_CreateColoredContextMenuSeparator


Func _ClampPopupToCurrentMonitor(ByRef $iX, ByRef $iY, $iW, $iH, $hOwner = 0)
	Local $iLeft = 0
	Local $iTop = 0
	Local $iRight = @DesktopWidth
	Local $iBottom = @DesktopHeight

	; Prefer the monitor under the mouse cursor; fall back to the app window monitor.
	Local $iMonitor = _Monitor_GetFromPoint($iX, $iY)
	Local $bMonitorFound = (Not @error And $iMonitor > 0)

	If Not $bMonitorFound And $hOwner <> 0 Then
		Local $iOwnerMonitor = _Monitor_GetFromWindow($hOwner)
		If Not @error And $iOwnerMonitor > 0 Then
			$iMonitor = $iOwnerMonitor
			$bMonitorFound = True
		EndIf
	EndIf

	If $bMonitorFound Then
		Local $iWorkLeft = 0, $iWorkTop = 0, $iWorkRight = 0, $iWorkBottom = 0
		If _Monitor_GetWorkArea($iMonitor, $iWorkLeft, $iWorkTop, $iWorkRight, $iWorkBottom) Then
			$iLeft = $iWorkLeft
			$iTop = $iWorkTop
			$iRight = $iWorkRight
			$iBottom = $iWorkBottom
		EndIf
	Else
		; Last-resort fallback for unusual monitor layouts.
		Local $aVirtual = _Monitor_GetVirtualBounds()
		If IsArray($aVirtual) Then
			$iLeft = $aVirtual[0]
			$iTop = $aVirtual[1]
			$iRight = $aVirtual[0] + $aVirtual[2]
			$iBottom = $aVirtual[1] + $aVirtual[3]
		EndIf
	EndIf

	If $iX + $iW > $iRight Then $iX = $iRight - $iW - 4
	If $iY + $iH > $iBottom Then $iY = $iBottom - $iH - 4
	If $iX < $iLeft Then $iX = $iLeft + 4
	If $iY < $iTop Then $iY = $iTop + 4
EndFunc   ;==>_ClampPopupToCurrentMonitor


Func _GetForegroundWindowHandle()
	Local $aRet = DllCall("user32.dll", "hwnd", "GetForegroundWindow")
	If @error Or Not IsArray($aRet) Then Return 0
	Return $aRet[0]
EndFunc   ;==>_GetForegroundWindowHandle


Func _MouseInsideRect($iX, $iY, $iW, $iH)
	Local $aPos = MouseGetPos()
	If @error Or Not IsArray($aPos) Then Return False
	If $aPos[0] < $iX Then Return False
	If $aPos[0] > ($iX + $iW) Then Return False
	If $aPos[1] < $iY Then Return False
	If $aPos[1] > ($iY + $iH) Then Return False
	Return True
EndFunc   ;==>_MouseInsideRect


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
	If @error Then $Color_bk_Today = 0xFF0000

	Global $Color_HighlightDate = RegRead($DB, "Color_HighlightDate")
	If @error Then $Color_HighlightDate = 0xFF0000

	Global $g_clrInvalidDayBG = RegRead($DB, "Color_InvalidDay")
	If @error Then $g_clrInvalidDayBG = 0xF0F0F0

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


; Applies the configured main ListView day-cell width.
Func _ApplyMainGridCellSize($iCellSize = -1)
	If $iCellSize = -1 Then $iCellSize = $Picker_Grid_Size_X_Read

	$iCellSize = Number($iCellSize)
	If $iCellSize < 20 Then $iCellSize = 20
	If $iCellSize > 60 Then $iCellSize = 60

	$Picker_Grid_Size_X_Read = $iCellSize

	If $g_hLV = 0 Then Return 0

	_LockWindow($g_hLV, False)
	For $iCol = 1 To 31
		_GUICtrlListView_SetColumnWidth($g_hLV, $iCol, $iCellSize)
	Next
	_LockWindow($g_hLV, True)
	_CleanRepaint($g_hLV)

	Return 1
EndFunc   ;==>_ApplyMainGridCellSize


; Refreshes the in-memory color/font arrays used by NM_CUSTOMDRAW.
; This is needed after changing category colors in Options because the ListView
; does not store the colors directly; it paints each cell from $g_aCellColor and
; $g_aCellColorBK. _Update() only refreshes the selected cell, so without this
; routine the main grid keeps the old category colors until the year is rebuilt.
Func _RefreshMainGridCellStyles($iTargetYear = -1)
	If $iTargetYear = -1 Then $iTargetYear = $iYear

	For $m = 1 To 12
		Local $sMonth = StringFormat("%02d", $m)
		Local $iDaysInMonth = _DaysInMonth2($iTargetYear, $m)

		For $d = 1 To 31
			If $d <= $iDaysInMonth Then
				Local $sDay = StringFormat("%02d", $d)
				Local $sStatus1 = RegRead($DB & "\" & $iTargetYear & "\" & $sMonth, $sDay)
				If @error Then $sStatus1 = ""

				Local $sStatus = StringLeft($sStatus1, 1)
				Local $sDisplay = _GetDateDisplayText($iTargetYear, $m, $d, $sStatus)

				If $g_hLV <> 0 And Number($iTargetYear) = Number($g_iLVYear) Then _GUICtrlListView_SetItemText($g_hLV, $iItem[$m][0], $sDisplay, $d)
				$g_aCellColor[$m - 1][$d] = _ColorFromDate($sStatus)
				$g_aCellColorBK[$m - 1][$d] = _GetDateFontColor($iTargetYear, $m, $d, $sStatus)
			Else
				; Keep unused days visually disabled instead of letting stale colors remain.
				_SetInvalidMainGridCell($m, $d)
			EndIf
		Next
	Next

	; Force the ListView/custom draw and the small calendar to repaint using the
	; refreshed arrays. If a Year Summary filter is active, keep that filtered
	; view instead of returning the grid to the full-year view.
	If $g_sMainGridFilter <> "" Then
		_ApplyMainGridCategoryFilter($g_sMainGridFilter)
		Return
	EndIf

	If $g_hLV <> 0 Then _CleanRepaint($g_hLV)
	_CustomCal_Update()
EndFunc   ;==>_RefreshMainGridCellStyles


Func _ReadDays($m, $iYear)

	Local $iMonth = Number($m)
	Local $sMonth = StringFormat("%02d", $iMonth)
	Local $iDaysInMonth = _DaysInMonth2($iYear, $iMonth)

	For $d = 1 To 31
		Local $iDay = Number($d)

		; The main grid always has 31 day columns, but not every month has 31 days.
		; Dates that do not exist must be rendered as disabled cells and must not
		; inherit the normal blank-day style.
		If $iDay > $iDaysInMonth Then
			_SetInvalidMainGridCell($iMonth, $iDay)
			ContinueLoop
		EndIf

		Local $sDay = StringFormat("%02d", $iDay)
		Local $Status1 = RegRead($DB & "\" & $iYear & "\" & $sMonth, $sDay)

		If @error Then
			Local $sDisplayBlank = _GetDateDisplayText($iYear, $iMonth, $iDay, "")
			If $sDisplayBlank <> "" Then _GUICtrlListView_AddSubItem($g_hLV, $iItem[$iMonth][0], $sDisplayBlank, $iDay, 1)

			$g_aCellColor[$iMonth - 1][$iDay] = _ColorFromDate("")
			$g_aCellColorBK[$iMonth - 1][$iDay] = _GetDateFontColor($iYear, $iMonth, $iDay, "")
			$g_aCellStatus[$iMonth - 1][$iDay] = ""
			$g_aCellTip[$iMonth - 1][$iDay] = ""
			ContinueLoop
		EndIf

		Local $Status = StringLeft($Status1, 1)
		Local $StatusName = "BLANK"

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
			Case "B", "", "   "
				$StatusName = "BLANK"
		EndSwitch

		Local $WeekDayNum = _DateToDayOfWeek($iYear, $iMonth, $iDay)
		Local $WeekDayName = _DateDayOfWeek($WeekDayNum, 1)
		Local $WeekDayNumber = _WeekNumberISO($iYear, $iMonth, $iDay)
		Local $Status_Comment_1 = StringTrimLeft($Status1, 1)
		Local $Status_Comment

		If $Status_Comment_1 <> "" Then
			$Status_Comment = $iYear & "/" & $sMonth & "/" & $sDay & @CRLF & $WeekDayName & " (Week: " & $WeekDayNumber & ") - " & $StatusName & @CRLF & "----" & @CRLF & "- " & StringReplace($Status_Comment_1, @CRLF, @CRLF & "- ")
		Else
			$Status_Comment = $iYear & "/" & $sMonth & "/" & $sDay & @CRLF & $WeekDayName & " (Week: " & $WeekDayNumber & ") - " & $StatusName
		EndIf

		Local $sDisplayStatus = _GetDateDisplayText($iYear, $iMonth, $iDay, $Status)
		_GUICtrlListView_AddSubItem($g_hLV, $iItem[$iMonth][0], $sDisplayStatus, $iDay, 1)

		$g_aCellColor[$iMonth - 1][$iDay] = _ColorFromDate($Status)
		$g_aCellColorBK[$iMonth - 1][$iDay] = _GetDateFontColor($iYear, $iMonth, $iDay, $Status)
		$g_aCellStatus[$iMonth - 1][$iDay] = $Status_Comment_1
		$g_aCellTip[$iMonth - 1][$iDay] = $Status_Comment
	Next
EndFunc   ;==>_ReadDays


Func _ReadINI($iYear, $Splash = 0)

	ConsoleWrite("$iYear: " & $iYear & @CRLF)

	; NOTE: _ReadStatistics is intentionally NOT called here.
	; _Update() calls _ReadINI() and then calls _ReadStatistics() once
	; at its end, avoiding a redundant double-call.
	; For the startup path (_ReadINI called directly), _ReadStatistics
	; is called by _CheckQuarter() which follows immediately after.

	$g_iLVYear = Number($iYear)
	$g_idLV = GUICtrlCreateListView("", 7, 210, 1127, 365, BitOR($LVS_REPORT, $LVS_SINGLESEL))
	If $g_idLV = 0 Then Exit MsgBox(16, "Error", "Failed to create ListView.")

	$g_hLV = GUICtrlGetHandle($g_idLV)
	If $g_hLV = 0 Then Exit MsgBox(16, "Error", "Failed to get ListView handle.")

	_GUICtrlListView_SetExtendedListViewStyle($g_hLV, $LVS_EX_GRIDLINES)

	; Columns
	_GUICtrlListView_InsertColumn($g_hLV, 0, "", 40, $LVCFMT_LEFT)

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

	For $m = 4 To 6
		$iItem[$m][0] = _GUICtrlListView_AddItem($g_hLV, $g_aMonths[$m - 1], -1)
		_ReadDays($m, $iYear)
	Next

	For $m = 7 To 9
		$iItem[$m][0] = _GUICtrlListView_AddItem($g_hLV, $g_aMonths[$m - 1], -1)
		_ReadDays($m, $iYear)
	Next

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

	; Register WM_NOTIFY handler exactly ONCE after the ListView is fully built.
	; Previously this was called inside _ReadDays' per-cell loop (up to 372 times).
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

EndFunc   ;==>_ReadINI


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
	$c = 0
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

		$c += 1
		If $c > 2 Then
			$c = 0
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


Func _Reload()

	; Do not disable redraw unconditionally here.
	; _Update() already handles WM_SETREDRAW safely when a real ListView rebuild is needed.
	; Locking the window here caused the UI to remain visually frozen on same-year refreshes
	; such as Add/Edit Marker -> OK/Cancel.
	_ReadColors()

	$SelDate = GUICtrlRead($Calendar)
	$SelDate_slipt = StringSplit($SelDate, "/")
	If @error Or $SelDate_slipt[0] <> 3 Then Return

	$Status1 = RegRead($DB & "\" & $SelDate_slipt[1] & "\" & $SelDate_slipt[2], $SelDate_slipt[3])
	$Status = StringTrimLeft($Status1, 1)
	_UpdateSelectionHighlight(Number($SelDate_slipt[3]), Number($SelDate_slipt[2]))

	_Chart()
	_Update($SelDate)
	_CreateMenu()

	Return

EndFunc   ;==>_Reload


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
					; Parse the backup line using the first "=" only.
					; This prevents settings values that contain backslashes (for example UNC paths)
					; from being incorrectly treated as calendar keys.
					Local $iEqualPos = StringInStr($HolidaysLine, "=")
					If $iEqualPos = 0 Then
						If StringStripWS($HolidaysLine, 3) <> "" Then $HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
						ContinueLoop
					EndIf

					Local $sBackupKey = StringLeft($HolidaysLine, $iEqualPos - 1)
					Local $sBackupValue = StringMid($HolidaysLine, $iEqualPos + 1)
					$sBackupValue = StringReplace($sBackupValue, " /n", @CRLF)

					If StringStripWS($sBackupKey, 3) = "" Then
						$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
						ContinueLoop
					EndIf

					If StringInStr($sBackupKey, "\") Then
						$HolidaysLine_key = StringSplit($sBackupKey, "\")
						If @error Or $HolidaysLine_key[0] <> 3 Then
							$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
							ContinueLoop
						EndIf

						If $HolidaysLine_key[1] = "" Or $HolidaysLine_key[2] = "" Or $HolidaysLine_key[3] = "" Then
							$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
							ContinueLoop
						EndIf

						$RegError = RegWrite($DB & "\" & $HolidaysLine_key[1] & "\" & $HolidaysLine_key[2], $HolidaysLine_key[3], "REG_SZ", $sBackupValue)
					Else
						$RegError = RegWrite($DB, $sBackupKey, "REG_SZ", $sBackupValue)
					EndIf

					If @error Then
						$HolidaysError = $HolidaysError & "Error to import line: " & $HolidaysLine & @CRLF
					Else
						$ImportCount += 1
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


Func _splash($Mode = "on")

	If $Mode = "on" Then

		$splashWin_X = 640
		$splashWin_Y = 360

		If $WinPos_X = -1 And $WinPos_Y = -1 Then
			Global $Form_Splash = GUICreate("", $splashWin_X, $splashWin_Y, -1, -1, $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW, $WS_EX_LAYERED))
		Else
			Global $Form_Splash = GUICreate("", $splashWin_X, $splashWin_Y, $WinPos_X + Round(($Window_X - $splashWin_X) - (($Window_X - $splashWin_X) / 2), 0), $WinPos_Y + Round(($Window_Y - $splashWin_Y) - (($Window_Y - $splashWin_Y) / 2), 0), $WS_POPUP, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW, $WS_EX_LAYERED))
		EndIf

		Global $Pic_Splash = GUICtrlCreatePic($sSplashPath, 5, 5, 630, 350)

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


Func _Update($SelDate)

	ConsoleWrite("$SelDate: " & $SelDate & @CRLF)

	Local $aDate = StringSplit($SelDate, "/")
	If @error Or $aDate[0] <> 3 Then Return SetError(1, 0, 0)

	Local $iDataYear = Number($aDate[1])
	Local $iDataMonth = Number($aDate[2])
	Local $iDataDay = Number($aDate[3])

	If $iDataMonth < 1 Or $iDataMonth > 12 Then Return SetError(2, 0, 0)
	If Not _IsValidCalendarDay($iDataYear, $iDataMonth, $iDataDay) Then Return SetError(3, 0, 0)

	Local $bRebuildList = ($g_idLV = 0 Or $g_hLV = 0 Or $g_iLVYear <> $iDataYear)

	; Freeze only when we need a structural rebuild.
	If $bRebuildList Then
		_LockWindow($Form_WorkDays, False)
		If $g_idLV <> 0 Then GUICtrlDelete($g_idLV)
		_ReadINI($iDataYear)
	EndIf

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

	Local $sDisplay = _GetDateDisplayText($iDataYear, $iDataMonth, $iDataDay, $sDataRegister)

	If $g_hLV <> 0 Then
		_GUICtrlListView_SetItemText($g_hLV, $iItem[$iDataMonth][0], $sDisplay, $iDataDay)
	EndIf

	$g_aCellColor[$iDataMonth - 1][$iDataDay] = _ColorFromDate($sDisplay)
	$g_aCellColorBK[$iDataMonth - 1][$iDataDay] = _GetDateFontColor($iDataYear, $iDataMonth, $iDataDay, $sDisplay)
	$g_aCellStatus[$iDataMonth - 1][$iDataDay] = $sTip
	$g_aCellTip[$iDataMonth - 1][$iDataDay] = $sStatusComment

	GUICtrlSetData($Input_SelDate, $iDataYear & "/" & $sDataMonth & "/" & $sDataDay)
	_GUICtrlMonthCal_SetCurSel($Calendar, $iDataYear, $iDataMonth, $iDataDay)
	_UpdateSelectionHighlight($iDataDay, $iDataMonth)
	_CheckQuarter()

	_ReadStatistics($iDataYear)

	If $bRebuildList Then
		_CreateMenu()
		_LockWindow($Form_WorkDays, True)
		_CleanRepaint($Form_WorkDays)
	EndIf

	$g_ccYear = $iDataYear
	$g_ccMonth = $iDataMonth
	_CustomCal_Update()

	Return 1

EndFunc   ;==>_Update


Func _UpdateListViewCellTip()
	Local $aCur = GUIGetCursorInfo($Form_WorkDays)
	If @error Or Not IsArray($aCur) Then
		_HideListViewCellTip()
		Return
	EndIf

	; Only the main ListView uses the manual tooltip.
	; The small calendar uses native GUICtrlSetTip on its controls to avoid hover/click instability.
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
	If $sTip = "" Then
		_HideListViewCellTip()
		Return
	EndIf

	If $g_iTipRow = $iRow And $g_iTipCol = $iCol And $g_sTipText = $sTip Then Return

	Local $aMouse = MouseGetPos()
	ToolTip($sTip, $aMouse[0] + 16, $aMouse[1] + 20)

	$g_iTipRow = $iRow
	$g_iTipCol = $iCol
	$g_sTipText = $sTip
	$g_bTipVisible = True
EndFunc   ;==>_UpdateListViewCellTip


Func _WorkDayInAWeekend($DateToCheck, $NewStatus)

	$DateToCheck_split = StringSplit($DateToCheck, "/")
	If @error Or $DateToCheck_split[0] <> 3 Then Return 1
	If Not _IsValidCalendarDay($DateToCheck_split[1], $DateToCheck_split[2], $DateToCheck_split[3]) Then Return 1

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


; ─────────────────────────────────────────────────────────────────────────────
; _LockWindow / _CleanRepaint
; Suppress all intermediate redraws while rebuilding the ListView, then force
; a single clean repaint at the end.  Uses DllCall directly so no specific
; WinAPI UDF version is required.  Literal values are used to avoid any
; redeclaration conflicts with the standard AutoIt includes.
; ─────────────────────────────────────────────────────────────────────────────
Func _LockWindow($hWnd, $bEnable)
	; WM_SETREDRAW = 0x000B
	; $bEnable: False = stop painting, True = resume painting
	DllCall("user32.dll", "lresult", "SendMessageW", _
			"hwnd", $hWnd, _
			"uint", 0x000B, _
			"wparam", $bEnable, _
			"lparam", 0)
EndFunc   ;==>_LockWindow

Func _CleanRepaint($hWnd)
	; RedrawWindow flags: RDW_INVALIDATE(0x0001) | RDW_ALLCHILDREN(0x0080) | RDW_UPDATENOW(0x0100)
	DllCall("user32.dll", "bool", "RedrawWindow", _
			"hwnd", $hWnd, _
			"ptr", 0, _
			"handle", 0, _
			"uint", 0x0181)
EndFunc   ;==>_CleanRepaint


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

			Local $sMonth = StringFormat("%02d", $iMonth)
			Local $sDay = StringFormat("%02d", $iDay)
			Local $sDate = $iYear & "/" & $sMonth & "/" & $sDay

			; Left click must only select the date.
			; The context menu is queued exclusively by NM_RCLICK.
			_RefreshSelectedDateUI($sDate)
			Return 0

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
					; Visually disable calendar dates that do not exist for that month
					; (for example Apr 31 or Feb 29 in non-leap years).
					If $iSub > _DaysInMonth2($iYear, $iItem + 1) Then
						DllStructSetData($tCD, "clrTextBk", _DecColorToRGBHex($g_clrInvalidDayBG))
						DllStructSetData($tCD, "clrText", _DecColorToRGBHex($g_clrInvalidDayFG))
						If $hDC <> 0 Then _WinAPI_SelectObject($hDC, $g_hFontNormal)
						Return $CDRF_NEWFONT
					EndIf

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

			_RefreshSelectedDateUI($sDate)

			$g_iMenuDay = $iDay
			$g_iMenuMonth = $iMonth
			$g_iMenuYear = $iYear
			$g_bShowCellMenu = True

			Return 0

	EndSwitch



	Return $GUI_RUNDEFMSG

EndFunc   ;==>WM_NOTIFY


; ======================================================================================================================
; WorkDays Outlook Agent integration
; Settings are owned by the main WorkDays application. The background agent only reads these registry values.
; ======================================================================================================================
Func _OA_SetCheck($idCtrl, $sValue)
	If String($sValue) = "1" Then
		GUICtrlSetState($idCtrl, $GUI_CHECKED)
	Else
		GUICtrlSetState($idCtrl, $GUI_UNCHECKED)
	EndIf
EndFunc   ;==>_OA_SetCheck

Func _OA_CheckTo01($iState)
	If BitAND($iState, $GUI_CHECKED) = $GUI_CHECKED Then Return "1"
	Return "0"
EndFunc   ;==>_OA_CheckTo01

Func _OA_SettingName($sSection, $sKey)
	Return $sSection & "_" & $sKey
EndFunc   ;==>_OA_SettingName

Func _OA_Read($sSection, $sKey, $sDefault)
	Local $sValue = RegRead($g_sOutlookAgentDB, _OA_SettingName($sSection, $sKey))
	If @error Or $sValue = "" Then Return $sDefault
	Return String($sValue)
EndFunc   ;==>_OA_Read

Func _OA_Write($sSection, $sKey, $sValue)
	Return RegWrite($g_sOutlookAgentDB, _OA_SettingName($sSection, $sKey), "REG_SZ", String($sValue))
EndFunc   ;==>_OA_Write

Func _OA_EnsureDefault($sSection, $sKey, $sDefault)
	RegRead($g_sOutlookAgentDB, _OA_SettingName($sSection, $sKey))
	If @error Then _OA_Write($sSection, $sKey, $sDefault)
EndFunc   ;==>_OA_EnsureDefault

Func _OutlookAgent_EnsureDefaults()
	_OA_EnsureDefault("Sync", "IntervalMinutes", "15")
	_OA_EnsureDefault("Sync", "PastDays", "60")
	_OA_EnsureDefault("Sync", "FutureDays", "370")
	_OA_EnsureDefault("Sync", "OutlookWinsOnConflict", "1")
	_OA_EnsureDefault("Sync", "DeleteInOutlookClearsWorkDays", "0")
	_OA_EnsureDefault("Sync", "SyncBlank", "0")
	_OA_EnsureDefault("Sync", "SyncWeekend", "0")
	_OA_EnsureDefault("Sync", "SyncTaggedBlankOrWeekend", "1")
	_OA_EnsureDefault("Sync", "RunAtWindowsStartup", "0")

	_OA_EnsureDefault("Outlook", "SubjectPrefix", "WorkDays -")
	_OA_EnsureDefault("Outlook", "CategoryPrefix", "WorkDays -")
	_OA_EnsureDefault("Outlook", "ReminderSet", "0")
	_OA_EnsureDefault("Outlook", "ManagedOnly", "0")
	_OA_EnsureDefault("Outlook", "DateOrder", "Auto")

	_OA_EnsureDefault("Markers", "ShowMarkerTagInSubject", "1")
	_OA_EnsureDefault("Markers", "MarkerSubjectSuffix", " [Marker]")
	_OA_EnsureDefault("Markers", "UseSeparateMarkerCategory", "1")
	_OA_EnsureDefault("Markers", "MarkerCategoryName", "WorkDays - Marker")
	_OA_EnsureDefault("Markers", "ReminderWhenMarkerExists", "1")
	_OA_EnsureDefault("Markers", "ReminderMinutesBeforeStart", "540")

	_OA_EnsureDefault("Safety", "EnableOutlookCleanup", "1")
	_OA_EnsureDefault("Safety", "CleanupPastYears", "10")
	_OA_EnsureDefault("Safety", "CleanupFutureYears", "10")
	_OA_EnsureDefault("Safety", "CleanupPrefixOnlyItems", "0")
	_OA_EnsureDefault("Safety", "PauseAfterOutlookCleanup", "1")
	_OA_EnsureDefault("Safety", "CleanupConfirmationPhrase", "CLEAN WORKDAYS OUTLOOK")
	_OA_EnsureDefault("Safety", "CreateBackupBeforeOutlookChanges", "1")
	_OA_EnsureDefault("Safety", "BlockMassChanges", "1")
	_OA_EnsureDefault("Safety", "MaxWorkDaysChangesPerSync", "20")
	_OA_EnsureDefault("Safety", "MaxChangePercentPerSync", "15")
	_OA_EnsureDefault("Safety", "MaxClearsPerSync", "0")
	_OA_EnsureDefault("Safety", "BlockIncompleteOutlookRead", "1")
	_OA_EnsureDefault("Safety", "IncompleteReadMinOutlookItems", "3")
	_OA_EnsureDefault("Safety", "IncompleteReadMinRatioPercent", "20")
	_OA_EnsureDefault("Safety", "RequireVisibleOutlookSession", "1")
	_OA_EnsureDefault("Safety", "StartupOutlookRetrySeconds", "15")
	_OA_EnsureDefault("Safety", "OutlookStartupGraceSeconds", "30")
	_OA_EnsureDefault("Safety", "OutlookReadyStableChecks", "3")
	_OA_EnsureDefault("Safety", "IncompleteReadRetries", "3")
	_OA_EnsureDefault("Safety", "IncompleteReadRetryDelayMs", "15000")

	_OA_EnsureDefault("Advanced", "LogLevel", "Normal")
	_OA_EnsureDefault("Logging", "VerboseMode", "0")
EndFunc   ;==>_OutlookAgent_EnsureDefaults

Func _OutlookAgent_IsInstalled()
	If FileExists($g_sOutlookAgentExe) Then Return True
	Return False
EndFunc   ;==>_OutlookAgent_IsInstalled

Func _OutlookAgent_IsRunning()
	Return ProcessExists($g_sOutlookAgentProcess) <> 0
EndFunc   ;==>_OutlookAgent_IsRunning

Func _OutlookAgent_StatusText()
	Local $sInstalled = "Not installed"
	If _OutlookAgent_IsInstalled() Then $sInstalled = "Installed"

	Local $sRunning = "Stopped"
	If _OutlookAgent_IsRunning() Then $sRunning = "Running"

	Return "Status: " & $sInstalled & " / " & $sRunning & @CRLF & "Location: " & $g_sOutlookAgentExe
EndFunc   ;==>_OutlookAgent_StatusText

Func _OutlookAgent_RunCommand()
	Return '"' & $g_sOutlookAgentExe & '"'
EndFunc   ;==>_OutlookAgent_RunCommand

Func _OutlookAgent_UpdateStartupRunKey()
	If _OA_Read("Sync", "RunAtWindowsStartup", "0") = "1" And FileExists($g_sOutlookAgentExe) Then
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "WorkDays Outlook Agent", "REG_SZ", _OutlookAgent_RunCommand())
	Else
		RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "WorkDays Outlook Agent")
	EndIf
EndFunc   ;==>_OutlookAgent_UpdateStartupRunKey

Func _OutlookAgent_Install()
	DirCreate($g_sOutlookAgentDir)
	_OutlookAgent_EnsureDefaults()

	If _OutlookAgent_IsRunning() Then
		Local $iStop = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent is currently running." & @CRLF & @CRLF & "WorkDays needs to close it before installing/updating the embedded agent." & @CRLF & @CRLF & "Continue?", 0, $Form_WorkDays)
		If $iStop <> $IDYES Then Return 0
		_OutlookAgent_Stop(False)
	EndIf

	; The agent executable is embedded when WorkDays is compiled. Keep this source path aligned with the build folder.
	FileInstall("Workdays_Outlook_Agent.exe", $g_sOutlookAgentExe, 1)
	If @error Or Not FileExists($g_sOutlookAgentExe) Then
		MsgBox($MB_ICONERROR + $MB_TOPMOST, "WorkDays Outlook Agent", "The embedded Outlook Agent could not be installed." & @CRLF & @CRLF & "When compiling WorkDays, confirm this file exists:" & @CRLF & @ScriptDir & "\Workdays_Outlook_Agent.exe", 0, $Form_WorkDays)
		Return 0
	EndIf

	RegWrite($g_sOutlookAgentDB, "Installed", "REG_SZ", "1")
	RegWrite($g_sOutlookAgentDB, "AgentPath", "REG_SZ", $g_sOutlookAgentExe)
	RegWrite($g_sOutlookAgentDB, "InstalledByWorkDays", "REG_SZ", "1")
	RegWrite($g_sOutlookAgentDB, "InstalledOn", "REG_SZ", StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))
	_OutlookAgent_UpdateStartupRunKey()

	MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "WorkDays Outlook Agent", "The embedded Outlook Agent was installed/updated successfully.", 0, $Form_WorkDays)
	Return 1
EndFunc   ;==>_OutlookAgent_Install

Func _OutlookAgent_Start()
	If Not _OutlookAgent_IsInstalled() Then
		Local $iInstall = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent is not installed yet." & @CRLF & @CRLF & "Install it now?", 0, $Form_WorkDays)
		If $iInstall <> $IDYES Then Return 0
		If Not _OutlookAgent_Install() Then Return 0
	EndIf

	If _OutlookAgent_IsRunning() Then Return 1
	Run(_OutlookAgent_RunCommand(), $g_sOutlookAgentDir)
	Sleep(800)
	If Not _OutlookAgent_IsRunning() Then
		MsgBox(BitOR($MB_ICONWARNING, $MB_TOPMOST), "WorkDays Outlook Agent", "WorkDays tried to start the Outlook Agent, but it does not appear to be running." & @CRLF & @CRLF & "Check the log file or Outlook installation.", 0, $Form_WorkDays)
		Return 0
	EndIf
	Return 1
EndFunc   ;==>_OutlookAgent_Start

Func _OutlookAgent_Stop($bShowMessage = True)
	If Not _OutlookAgent_IsRunning() Then Return 1
	ProcessClose($g_sOutlookAgentProcess)
	ProcessWaitClose($g_sOutlookAgentProcess, 5)
	If _OutlookAgent_IsRunning() Then
		If $bShowMessage Then MsgBox(BitOR($MB_ICONWARNING, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent is still running. You may need to close it manually from the tray.", 0, $Form_WorkDays)
		Return 0
	EndIf
	If $bShowMessage Then MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent was stopped.", 0, $Form_WorkDays)
	Return 1
EndFunc   ;==>_OutlookAgent_Stop

Func _OutlookAgent_OpenLog()
	If FileExists($g_sOutlookAgentLog) Then
		ShellExecute($g_sOutlookAgentLog)
	Else
		MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent log file has not been created yet.", 0, $Form_WorkDays)
	EndIf
EndFunc   ;==>_OutlookAgent_OpenLog

Func _OutlookAgent_CleanOutlookFromWorkDays()
	_OutlookAgent_EnsureDefaults()
	If _OA_Read("Safety", "EnableOutlookCleanup", "1") <> "1" Then
		MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "Outlook cleanup is disabled in the Agent settings.", 0, $Form_WorkDays)
		Return 0
	EndIf

	If Not _OutlookAgent_IsInstalled() Then
		MsgBox(BitOR($MB_ICONWARNING, $MB_TOPMOST), "WorkDays Outlook Agent", "Install the Outlook Agent before running Outlook cleanup.", 0, $Form_WorkDays)
		Return 0
	EndIf

	Local $sMsg = "This will remove WorkDays calendar items from Outlook only." & @CRLF & @CRLF & _
		"Your WorkDays data will remain saved in the WorkDays application." & @CRLF & _
		"The agent will be stopped before cleanup to avoid recreating items during the operation." & @CRLF & _
		"The synchronization state file will also be deleted." & @CRLF & @CRLF & _
		"Continue?"
	If MsgBox(BitOR($MB_ICONWARNING, $MB_YESNO, $MB_DEFBUTTON2, $MB_TOPMOST), "WorkDays Outlook Agent", $sMsg, 0, $Form_WorkDays) <> $IDYES Then Return 0

	If _OutlookAgent_IsRunning() Then _OutlookAgent_Stop(False)

	Local $iRC = RunWait(_OutlookAgent_RunCommand() & " /cleanoutlook", $g_sOutlookAgentDir)
	If $iRC <> 0 Then
		MsgBox(BitOR($MB_ICONERROR, $MB_TOPMOST), "WorkDays Outlook Agent", "Outlook cleanup failed. Open the log for details.", 0, $Form_WorkDays)
		Return 0
	EndIf

	; The agent deletes the state file as part of a successful cleanup.
	; Repeat the deletion here as a defensive fallback in case an older installed agent was used.
	If FileExists($g_sOutlookAgentState) Then
		If Not FileDelete($g_sOutlookAgentState) Then
			MsgBox(BitOR($MB_ICONERROR, $MB_TOPMOST), "WorkDays Outlook Agent", "Outlook items were cleaned, but the synchronization state file could not be deleted:" & @CRLF & @CRLF & $g_sOutlookAgentState & @CRLF & @CRLF & "Close any process using the file and delete it manually before restarting the agent.", 0, $Form_WorkDays)
			Return 0
		EndIf
	EndIf

	If _OA_Read("Safety", "PauseAfterOutlookCleanup", "1") = "1" Then
		MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "Outlook cleanup completed." & @CRLF & @CRLF & "The synchronization state was reset, and the agent was left stopped so the calendar stays clean until you start it again.", 0, $Form_WorkDays)
	Else
		_OutlookAgent_Start()
		MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "Outlook cleanup completed, the synchronization state was reset, and the agent was restarted.", 0, $Form_WorkDays)
	EndIf

	Return 1
EndFunc   ;==>_OutlookAgent_CleanOutlookFromWorkDays

Func _OutlookAgent_Uninstall()
	Local $sMsg = "This will remove the Outlook Agent executable installed by WorkDays." & @CRLF & @CRLF & _
		"Your WorkDays data and Agent settings will remain saved." & @CRLF & _
		"Outlook calendar items will not be deleted. Use Clean Outlook WorkDays Items first if you want a clean calendar." & @CRLF & @CRLF & _
		"Continue?"
	If MsgBox(BitOR($MB_ICONWARNING, $MB_YESNO, $MB_DEFBUTTON2, $MB_TOPMOST), "WorkDays Outlook Agent", $sMsg, 0, $Form_WorkDays) <> $IDYES Then Return 0

	If _OutlookAgent_IsRunning() Then _OutlookAgent_Stop(False)
	RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "WorkDays Outlook Agent")
	_OA_Write("Sync", "RunAtWindowsStartup", "0")
	If FileExists($g_sOutlookAgentExe) Then FileDelete($g_sOutlookAgentExe)
	RegWrite($g_sOutlookAgentDB, "Installed", "REG_SZ", "0")
	MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent was uninstalled from this Windows profile.", 0, $Form_WorkDays)
	Return 1
EndFunc   ;==>_OutlookAgent_Uninstall

Func _OutlookAgent_SaveSettings($idInterval, $idPast, $idFuture, $idConflictOutlook, $idDeleteOutlookClears, $idSyncBlank, $idSyncWeekend, $idSyncTaggedBlankWeekend, $idStartup, $idSubjectPrefix, $idCategoryPrefix, $idReminderSet, $idManagedOnly, $idShowMarkerSubject, $idMarkerSuffix, $idSeparateMarkerCategory, $idMarkerCategoryName, $idReminderMarker, $idReminderMinutes, $idCleanupEnabled, $idCleanupPastYears, $idCleanupFutureYears, $idCleanupPrefixOnly, $idPauseAfterCleanup, $idVerboseMode, $idBackupBeforeOutlookChanges, $idBlockMassChanges, $idMaxChanges, $idMaxPercent, $idMaxClears, $idBlockIncompleteRead)
	Local $iInterval = Number(GUICtrlRead($idInterval))
	If $iInterval < 1 Then $iInterval = 15
	Local $iPast = Number(GUICtrlRead($idPast))
	If $iPast < 0 Then $iPast = 60
	Local $iFuture = Number(GUICtrlRead($idFuture))
	If $iFuture < 1 Then $iFuture = 370
	Local $iReminderMinutes = Number(GUICtrlRead($idReminderMinutes))
	If $iReminderMinutes < 0 Then $iReminderMinutes = 0
	Local $iCleanupPast = Number(GUICtrlRead($idCleanupPastYears))
	If $iCleanupPast < 0 Then $iCleanupPast = 10
	Local $iCleanupFuture = Number(GUICtrlRead($idCleanupFutureYears))
	If $iCleanupFuture < 0 Then $iCleanupFuture = 10
	Local $iMaxChanges = Number(GUICtrlRead($idMaxChanges))
	If $iMaxChanges < 1 Then $iMaxChanges = 20
	Local $iMaxPercent = Number(GUICtrlRead($idMaxPercent))
	If $iMaxPercent < 1 Then $iMaxPercent = 15
	Local $iMaxClears = Number(GUICtrlRead($idMaxClears))
	If $iMaxClears < 0 Then $iMaxClears = 0

	_OA_Write("Sync", "IntervalMinutes", $iInterval)
	_OA_Write("Sync", "PastDays", $iPast)
	_OA_Write("Sync", "FutureDays", $iFuture)
	_OA_Write("Sync", "OutlookWinsOnConflict", _OA_CheckTo01(GUICtrlRead($idConflictOutlook)))
	_OA_Write("Sync", "DeleteInOutlookClearsWorkDays", _OA_CheckTo01(GUICtrlRead($idDeleteOutlookClears)))
	_OA_Write("Sync", "SyncBlank", _OA_CheckTo01(GUICtrlRead($idSyncBlank)))
	_OA_Write("Sync", "SyncWeekend", _OA_CheckTo01(GUICtrlRead($idSyncWeekend)))
	_OA_Write("Sync", "SyncTaggedBlankOrWeekend", _OA_CheckTo01(GUICtrlRead($idSyncTaggedBlankWeekend)))
	_OA_Write("Sync", "RunAtWindowsStartup", _OA_CheckTo01(GUICtrlRead($idStartup)))

	_OA_Write("Outlook", "SubjectPrefix", GUICtrlRead($idSubjectPrefix))
	_OA_Write("Outlook", "CategoryPrefix", GUICtrlRead($idCategoryPrefix))
	_OA_Write("Outlook", "ReminderSet", _OA_CheckTo01(GUICtrlRead($idReminderSet)))
	_OA_Write("Outlook", "ManagedOnly", _OA_CheckTo01(GUICtrlRead($idManagedOnly)))

	_OA_Write("Markers", "ShowMarkerTagInSubject", _OA_CheckTo01(GUICtrlRead($idShowMarkerSubject)))
	_OA_Write("Markers", "MarkerSubjectSuffix", GUICtrlRead($idMarkerSuffix))
	_OA_Write("Markers", "UseSeparateMarkerCategory", _OA_CheckTo01(GUICtrlRead($idSeparateMarkerCategory)))
	_OA_Write("Markers", "MarkerCategoryName", GUICtrlRead($idMarkerCategoryName))
	_OA_Write("Markers", "ReminderWhenMarkerExists", _OA_CheckTo01(GUICtrlRead($idReminderMarker)))
	_OA_Write("Markers", "ReminderMinutesBeforeStart", $iReminderMinutes)

	_OA_Write("Safety", "EnableOutlookCleanup", _OA_CheckTo01(GUICtrlRead($idCleanupEnabled)))
	_OA_Write("Safety", "CleanupPastYears", $iCleanupPast)
	_OA_Write("Safety", "CleanupFutureYears", $iCleanupFuture)
	_OA_Write("Safety", "CleanupPrefixOnlyItems", _OA_CheckTo01(GUICtrlRead($idCleanupPrefixOnly)))
	_OA_Write("Safety", "PauseAfterOutlookCleanup", _OA_CheckTo01(GUICtrlRead($idPauseAfterCleanup)))
	_OA_Write("Safety", "CreateBackupBeforeOutlookChanges", _OA_CheckTo01(GUICtrlRead($idBackupBeforeOutlookChanges)))
	_OA_Write("Safety", "BlockMassChanges", _OA_CheckTo01(GUICtrlRead($idBlockMassChanges)))
	_OA_Write("Safety", "MaxWorkDaysChangesPerSync", $iMaxChanges)
	_OA_Write("Safety", "MaxChangePercentPerSync", $iMaxPercent)
	_OA_Write("Safety", "MaxClearsPerSync", $iMaxClears)
	_OA_Write("Safety", "BlockIncompleteOutlookRead", _OA_CheckTo01(GUICtrlRead($idBlockIncompleteRead)))

	_OA_Write("Logging", "VerboseMode", _OA_CheckTo01(GUICtrlRead($idVerboseMode)))
	If _OA_CheckTo01(GUICtrlRead($idVerboseMode)) = "1" Then
		_OA_Write("Advanced", "LogLevel", "Verbose")
	Else
		_OA_Write("Advanced", "LogLevel", "Normal")
	EndIf

	RegWrite($g_sOutlookAgentDB, "LastSettingsSavedByWorkDays", "REG_SZ", StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))
	_OutlookAgent_UpdateStartupRunKey()
	Return 1
EndFunc   ;==>_OutlookAgent_SaveSettings


Func _OutlookAgent_StatusCodeToLabel($sStatus)
	Switch $sStatus
		Case "O"
			Return "On Site"
		Case "R"
			Return "Remote"
		Case "H"
			Return "Holiday"
		Case "P"
			Return "PTO"
		Case "T"
			Return "Travel"
		Case "S"
			Return "Sick"
		Case "W"
			Return "Weekend"
		Case "B"
			Return "Blank"
		Case Else
			Return ""
	EndSwitch
EndFunc   ;==>_OutlookAgent_StatusCodeToLabel

Func _OutlookAgent_RequestSyncNow()
	_OutlookAgent_EnsureDefaults()
	GUICtrlSetData($Button_OutlookSync, "Sync...")

	Local $sNow = StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC) & "." & @MSEC
	RegWrite($g_sOutlookAgentDB, "Sync_ForceNowRequest", "REG_SZ", $sNow)
	RegWrite($g_sOutlookAgentDB, "Sync_ForceNowRequestedBy", "REG_SZ", "WorkDays")

	If _OutlookAgent_IsRunning() Then
		Sleep(500)
		GUICtrlSetData($Button_OutlookSync, "Sync")
		Return 1
	EndIf

	If Not _OutlookAgent_IsInstalled() Then
		GUICtrlSetData($Button_OutlookSync, "Sync")
		Local $iInstall = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO, $MB_TOPMOST), "WorkDays Outlook Agent", "The Outlook Agent is not installed yet." & @CRLF & @CRLF & "Install it now?", 0, $Form_WorkDays)
		If $iInstall <> $IDYES Then Return 0
		If Not _OutlookAgent_Install() Then Return 0
		GUICtrlSetData($Button_OutlookSync, "Sync...")
	EndIf

	; Start the resident agent instead of a one-time COM sync. The agent keeps the request queued
	; and waits until a visible Outlook session is fully initialized.
	Run(_OutlookAgent_RunCommand(), $g_sOutlookAgentDir)
	Sleep(500)
	GUICtrlSetData($Button_OutlookSync, "Sync")
	Return 1
EndFunc   ;==>_OutlookAgent_RequestSyncNow

Func _OutlookAgent_CheckRefreshNotification()
	Static $sLastSeenSeq = "__INIT__"

	; First, check if the agent blocked a sync for safety reasons.
	Local $sGuardStatus = RegRead($g_sOutlookAgentDB, "LastSyncGuardStatus")
	If @error Then $sGuardStatus = ""
	Local $sGuardAt = RegRead($g_sOutlookAgentDB, "LastSyncGuardAt")
	If @error Then $sGuardAt = ""
	Local $sGuardAck = RegRead($g_sOutlookAgentDB, "LastSyncGuardAcknowledgedValue")
	If @error Then $sGuardAck = ""

	If $sGuardStatus = "BLOCKED" And $sGuardAt <> "" And $sGuardAt <> $sGuardAck Then
		$g_bOutlookAgentSyncBlockedPending = True
		$g_sOutlookAgentLastGuardStatus = $sGuardAt
		$g_sOutlookAgentSyncBlockedReason = RegRead($g_sOutlookAgentDB, "LastSyncGuardReason")
		If @error Then $g_sOutlookAgentSyncBlockedReason = ""
		$g_sOutlookAgentSyncBlockedPlanFile = RegRead($g_sOutlookAgentDB, "LastSyncPlanFile")
		If @error Then $g_sOutlookAgentSyncBlockedPlanFile = ""
		GUICtrlSetData($Button_Update, "SYNC BLOCKED - Review")
		GUICtrlSetColor($Button_Update, 0xFFFFFF)
		GUICtrlSetBkColor($Button_Update, 0xF39C12)
		GUICtrlSetFont($Button_Update, 9, 900)
		GUICtrlSetState($Button_Update, $GUI_SHOW)
		Return 1
	EndIf

	Local $sSeq = RegRead($g_sOutlookAgentDB, "LastDatabaseChangeSeq")
	If @error Then $sSeq = ""

	If $sLastSeenSeq = "__INIT__" Then
		$sLastSeenSeq = $sSeq
		Return 0
	EndIf

	If $sSeq = "" Or $sSeq = $sLastSeenSeq Then Return 0

	$sLastSeenSeq = $sSeq
	$g_bOutlookAgentRefreshPending = True
	$g_sOutlookAgentPendingSeq = $sSeq
	$g_sOutlookAgentPendingDate = RegRead($g_sOutlookAgentDB, "LastDatabaseChangeDate")
	If @error Then $g_sOutlookAgentPendingDate = ""
	$g_sOutlookAgentPendingStatus = RegRead($g_sOutlookAgentDB, "LastDatabaseChangeStatus")
	If @error Then $g_sOutlookAgentPendingStatus = ""

	Local $sText = "OUTLOOK CHANGE - Click to refresh"
	If $g_sOutlookAgentPendingDate <> "" Then $sText = "OUTLOOK CHANGE " & $g_sOutlookAgentPendingDate & " - Refresh"
	GUICtrlSetData($Button_Update, $sText)
	GUICtrlSetColor($Button_Update, 0xFFFFFF)
	GUICtrlSetBkColor($Button_Update, 0x2F80ED)
	GUICtrlSetFont($Button_Update, 9, 900)
	GUICtrlSetState($Button_Update, $GUI_SHOW)
	Return 1
EndFunc   ;==>_OutlookAgent_CheckRefreshNotification

Func _OutlookAgent_ShowSyncBlockedGuard()
	Local $sMsg = "The Outlook Agent blocked the last sync to protect the WorkDays database." & @CRLF & @CRLF
	If $g_sOutlookAgentSyncBlockedReason <> "" Then $sMsg &= "Reason:" & @CRLF & $g_sOutlookAgentSyncBlockedReason & @CRLF & @CRLF
	Local $sBackup = RegRead($g_sOutlookAgentDB, "LastPreSyncBackup")
	If @error Then $sBackup = ""
	If $sBackup <> "" Then $sMsg &= "Backup created before blocking/applying changes:" & @CRLF & $sBackup & @CRLF & @CRLF
	If $g_sOutlookAgentSyncBlockedPlanFile <> "" Then $sMsg &= "Sync plan:" & @CRLF & $g_sOutlookAgentSyncBlockedPlanFile & @CRLF & @CRLF
	$sMsg &= "No mass database changes were applied." & @CRLF & @CRLF & "Open the last sync plan now?"

	Local $iAns = MsgBox(BitOR($MB_ICONWARNING, $MB_YESNO, $MB_TOPMOST), "WorkDays Outlook Agent", $sMsg, 0, $Form_WorkDays)
	If $iAns = $IDYES And $g_sOutlookAgentSyncBlockedPlanFile <> "" And FileExists($g_sOutlookAgentSyncBlockedPlanFile) Then ShellExecute($g_sOutlookAgentSyncBlockedPlanFile)

	RegWrite($g_sOutlookAgentDB, "LastSyncGuardAcknowledgedValue", "REG_SZ", $g_sOutlookAgentLastGuardStatus)
	RegWrite($g_sOutlookAgentDB, "LastSyncGuardAcknowledgedAt", "REG_SZ", StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))

	$g_bOutlookAgentSyncBlockedPending = False
	$g_sOutlookAgentSyncBlockedReason = ""
	$g_sOutlookAgentSyncBlockedPlanFile = ""
	$g_sOutlookAgentLastGuardStatus = ""

	If $g_bWorkDaysUpdaterAvailable Or FileExists(@ScriptDir & "\WorkDays.tmp") Then
		$g_bWorkDaysUpdaterAvailable = True
		GUICtrlSetData($Button_Update, "UPDATE AVAILABLE - Click to execute")
		GUICtrlSetColor($Button_Update, 0xFFFFFF)
		GUICtrlSetBkColor($Button_Update, 0xFF0000)
		GUICtrlSetState($Button_Update, $GUI_SHOW)
	Else
		GUICtrlSetState($Button_Update, $GUI_HIDE)
	EndIf
	Return 1
EndFunc   ;==>_OutlookAgent_ShowSyncBlockedGuard

Func _OutlookAgent_RefreshWorkDaysFromAgentChange()
	Local $sMsg = "The Outlook Agent changed the WorkDays database."
	If $g_sOutlookAgentPendingDate <> "" Then
		$sMsg &= @CRLF & @CRLF & "Date: " & $g_sOutlookAgentPendingDate
		Local $sLabel = _OutlookAgent_StatusCodeToLabel($g_sOutlookAgentPendingStatus)
		If $sLabel <> "" Then $sMsg &= @CRLF & "Status: " & $sLabel
	EndIf
	$sMsg &= @CRLF & @CRLF & "Refreshing the screen now."

	MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", $sMsg, 0, $Form_WorkDays)
	_Reload()

	RegWrite($g_sOutlookAgentDB, "LastDatabaseChangeAcknowledgedByWorkDays", "REG_SZ", $g_sOutlookAgentPendingSeq)
	RegWrite($g_sOutlookAgentDB, "LastDatabaseChangeAcknowledgedAt", "REG_SZ", StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))

	$g_bOutlookAgentRefreshPending = False
	$g_sOutlookAgentPendingSeq = ""
	$g_sOutlookAgentPendingDate = ""
	$g_sOutlookAgentPendingStatus = ""

	If $g_bWorkDaysUpdaterAvailable Or FileExists(@ScriptDir & "\WorkDays.tmp") Then
		$g_bWorkDaysUpdaterAvailable = True
		GUICtrlSetData($Button_Update, "UPDATE AVAILABLE - Click to execute")
		GUICtrlSetColor($Button_Update, 0xFFFFFF)
		GUICtrlSetBkColor($Button_Update, 0xFF0000)
		GUICtrlSetState($Button_Update, $GUI_SHOW)
	Else
		GUICtrlSetState($Button_Update, $GUI_HIDE)
	EndIf
	Return 1
EndFunc   ;==>_OutlookAgent_RefreshWorkDaysFromAgentChange

Func _OutlookAgent_SettingsWindow()
	_OutlookAgent_EnsureDefaults()

	; Compact wide layout to fit smaller screens without exceeding monitor height.
	$g_hOutlookAgentSettingsWindow = GUICreate("WorkDays Outlook Agent", 960, 640, -1, -1, $DS_MODALFRAME, BitOR($WS_EX_TOPMOST, $WS_EX_MDICHILD), $Form_WorkDays)
	Local $hAgent = $g_hOutlookAgentSettingsWindow
	GUISetBkColor(0xF7FBFF, $hAgent)
	GUISetFont(9, 400, 0, "Segoe UI", $hAgent)

	GUICtrlCreateLabel("Outlook Agent", 18, 12, 240, 24)
	GUICtrlSetFont(-1, 13, 700, 0, "Segoe UI")
	GUICtrlSetColor(-1, 0x0B4F8A)

	Local $lblStatus = GUICtrlCreateLabel(_OutlookAgent_StatusText(), 18, 42, 920, 38)
	GUICtrlSetColor($lblStatus, 0x1D3557)

	Local $btnInstall = GUICtrlCreateButton("Install / Update", 18, 88, 120, 28)
	Local $btnStart = GUICtrlCreateButton("Start", 148, 88, 80, 28)
	Local $btnStop = GUICtrlCreateButton("Stop", 238, 88, 80, 28)
	Local $btnOpenLog = GUICtrlCreateButton("Open log", 328, 88, 90, 28)
	Local $btnUninstall = GUICtrlCreateButton("Uninstall", 818, 88, 120, 28)

	; Left column: sync behavior.
	GUICtrlCreateGroup("Sync behavior", 18, 128, 460, 168)
	GUICtrlCreateLabel("Sync every", 34, 154, 70, 20)
	Local $inpInterval = GUICtrlCreateInput(_OA_Read("Sync", "IntervalMinutes", "15"), 104, 150, 46, 22, $ES_NUMBER)
	GUICtrlCreateLabel("minutes", 156, 154, 60, 20)
	GUICtrlCreateLabel("Range", 226, 154, 42, 20)
	Local $inpPast = GUICtrlCreateInput(_OA_Read("Sync", "PastDays", "60"), 274, 150, 50, 22, $ES_NUMBER)
	GUICtrlCreateLabel("past /", 330, 154, 42, 20)
	Local $inpFuture = GUICtrlCreateInput(_OA_Read("Sync", "FutureDays", "370"), 370, 150, 58, 22, $ES_NUMBER)
	GUICtrlCreateLabel("future", 432, 154, 45, 20)
	Local $chkConflictOutlook = GUICtrlCreateCheckbox("Outlook wins when both sides changed before sync", 34, 182, 385, 20)
	Local $chkDeleteOutlookClears = GUICtrlCreateCheckbox("Deleting the Outlook item clears WorkDays", 34, 208, 330, 20)
	Local $chkSyncBlank = GUICtrlCreateCheckbox("Sync Blank days", 34, 234, 135, 20)
	Local $chkSyncWeekend = GUICtrlCreateCheckbox("Sync Weekend days", 178, 234, 145, 20)
	Local $chkSyncTaggedBlankWeekend = GUICtrlCreateCheckbox("Blank/Weekend only with marker", 34, 260, 235, 20)
	Local $chkStartup = GUICtrlCreateCheckbox("Start Outlook Agent with Windows", 258, 260, 200, 20)

	; Right column: Outlook display.
	GUICtrlCreateGroup("Outlook display", 488, 128, 450, 118)
	GUICtrlCreateLabel("Subject prefix", 504, 154, 82, 20)
	Local $inpSubjectPrefix = GUICtrlCreateInput(_OA_Read("Outlook", "SubjectPrefix", "WorkDays -"), 590, 150, 130, 22)
	GUICtrlCreateLabel("Category", 732, 154, 58, 20)
	Local $inpCategoryPrefix = GUICtrlCreateInput(_OA_Read("Outlook", "CategoryPrefix", "WorkDays -"), 794, 150, 130, 22)
	Local $chkReminderSet = GUICtrlCreateCheckbox("Always use Outlook reminder", 504, 184, 210, 20)
	Local $chkManagedOnly = GUICtrlCreateCheckbox("Only read items created by the agent", 720, 184, 210, 20)
	GUICtrlCreateLabel("Appointments are created as all-day Free time events.", 504, 216, 410, 18)
	GUICtrlSetColor(-1, 0x577590)

	; Left column: marker alerts.
	GUICtrlCreateGroup("Marker alerts", 18, 308, 460, 142)
	Local $chkShowMarkerSubject = GUICtrlCreateCheckbox("Add marker indicator to subject", 34, 334, 220, 20)
	GUICtrlCreateLabel("Suffix", 270, 338, 45, 18)
	Local $inpMarkerSuffix = GUICtrlCreateInput(_OA_Read("Markers", "MarkerSubjectSuffix", " [Marker]"), 320, 334, 130, 22)
	Local $chkSeparateMarkerCategory = GUICtrlCreateCheckbox("Add separate Outlook category", 34, 364, 225, 20)
	GUICtrlCreateLabel("Category", 270, 368, 60, 18)
	Local $inpMarkerCategory = GUICtrlCreateInput(_OA_Read("Markers", "MarkerCategoryName", "WorkDays - Marker"), 332, 364, 120, 22)
	Local $chkReminderMarker = GUICtrlCreateCheckbox("Show Outlook reminder when marker exists", 34, 396, 280, 20)
	GUICtrlCreateLabel("Minutes before all-day event", 34, 424, 165, 18)
	Local $inpReminderMinutes = GUICtrlCreateInput(_OA_Read("Markers", "ReminderMinutesBeforeStart", "540"), 205, 420, 58, 22, $ES_NUMBER)

	; Left column: cleanup safety.
	GUICtrlCreateGroup("Safety", 18, 462, 460, 88)
	Local $chkCleanupEnabled = GUICtrlCreateCheckbox("Allow Outlook cleanup from WorkDays", 34, 488, 250, 20)
	Local $chkCleanupPrefixOnly = GUICtrlCreateCheckbox("Cleanup old prefix-only items", 272, 488, 185, 20)
	Local $chkPauseAfterCleanup = GUICtrlCreateCheckbox("Keep agent stopped after cleanup", 34, 518, 230, 20)
	GUICtrlCreateLabel("Cleanup range", 266, 522, 85, 18)
	Local $inpCleanupPast = GUICtrlCreateInput(_OA_Read("Safety", "CleanupPastYears", "10"), 352, 518, 36, 22, $ES_NUMBER)
	GUICtrlCreateLabel("past /", 392, 522, 38, 18)
	Local $inpCleanupFuture = GUICtrlCreateInput(_OA_Read("Safety", "CleanupFutureYears", "10"), 432, 518, 36, 22, $ES_NUMBER)

	; Right column: database protection.
	GUICtrlCreateGroup("Sync Safety", 488, 258, 450, 168)
	Local $chkBackupBeforeOutlookChanges = GUICtrlCreateCheckbox("Create backup before Outlook changes WorkDays", 504, 284, 300, 20)
	Local $chkBlockMassChanges = GUICtrlCreateCheckbox("Block mass changes", 812, 284, 125, 20)
	GUICtrlCreateLabel("Max changes", 504, 314, 80, 18)
	Local $inpMaxChanges = GUICtrlCreateInput(_OA_Read("Safety", "MaxWorkDaysChangesPerSync", "20"), 588, 310, 46, 22, $ES_NUMBER)
	GUICtrlCreateLabel("Max %", 648, 314, 50, 18)
	Local $inpMaxPercent = GUICtrlCreateInput(_OA_Read("Safety", "MaxChangePercentPerSync", "15"), 700, 310, 46, 22, $ES_NUMBER)
	GUICtrlCreateLabel("Max clears", 760, 314, 70, 18)
	Local $inpMaxClears = GUICtrlCreateInput(_OA_Read("Safety", "MaxClearsPerSync", "0"), 834, 310, 46, 22, $ES_NUMBER)
	Local $chkBlockIncompleteRead = GUICtrlCreateCheckbox("Block sync if Outlook read looks incomplete", 504, 342, 285, 20)
	Local $btnOpenBackupFolder = GUICtrlCreateButton("Open backup folder", 504, 370, 130, 24)
	Local $btnOpenPlan = GUICtrlCreateButton("Open last sync plan", 646, 370, 140, 24)
	GUICtrlCreateLabel("Recommended defaults: backup ON, mass-change block ON, max clears 0, Outlook deletion does not clear WorkDays.", 504, 400, 410, 18)
	GUICtrlSetColor(-1, 0x577590)

	; Right column: diagnostics.
	GUICtrlCreateGroup("Diagnostics", 488, 462, 450, 88)
	Local $chkVerboseMode = GUICtrlCreateCheckbox("Verbose diagnostic log", 504, 494, 180, 20)
	GUICtrlCreateLabel("Logs item inspection, candidate decision, date parsing, registry write, state update, safety plan, backup and guard decision.", 690, 482, 230, 60)
	GUICtrlSetColor(-1, 0x577590)

	_OA_SetCheck($chkConflictOutlook, _OA_Read("Sync", "OutlookWinsOnConflict", "1"))
	_OA_SetCheck($chkDeleteOutlookClears, _OA_Read("Sync", "DeleteInOutlookClearsWorkDays", "0"))
	_OA_SetCheck($chkSyncBlank, _OA_Read("Sync", "SyncBlank", "0"))
	_OA_SetCheck($chkSyncWeekend, _OA_Read("Sync", "SyncWeekend", "0"))
	_OA_SetCheck($chkSyncTaggedBlankWeekend, _OA_Read("Sync", "SyncTaggedBlankOrWeekend", "1"))
	_OA_SetCheck($chkStartup, _OA_Read("Sync", "RunAtWindowsStartup", "0"))
	_OA_SetCheck($chkReminderSet, _OA_Read("Outlook", "ReminderSet", "0"))
	_OA_SetCheck($chkManagedOnly, _OA_Read("Outlook", "ManagedOnly", "0"))
	_OA_SetCheck($chkShowMarkerSubject, _OA_Read("Markers", "ShowMarkerTagInSubject", "1"))
	_OA_SetCheck($chkSeparateMarkerCategory, _OA_Read("Markers", "UseSeparateMarkerCategory", "1"))
	_OA_SetCheck($chkReminderMarker, _OA_Read("Markers", "ReminderWhenMarkerExists", "1"))
	_OA_SetCheck($chkCleanupEnabled, _OA_Read("Safety", "EnableOutlookCleanup", "1"))
	_OA_SetCheck($chkCleanupPrefixOnly, _OA_Read("Safety", "CleanupPrefixOnlyItems", "0"))
	_OA_SetCheck($chkPauseAfterCleanup, _OA_Read("Safety", "PauseAfterOutlookCleanup", "1"))
	_OA_SetCheck($chkBackupBeforeOutlookChanges, _OA_Read("Safety", "CreateBackupBeforeOutlookChanges", "1"))
	_OA_SetCheck($chkBlockMassChanges, _OA_Read("Safety", "BlockMassChanges", "1"))
	_OA_SetCheck($chkBlockIncompleteRead, _OA_Read("Safety", "BlockIncompleteOutlookRead", "1"))
	_OA_SetCheck($chkVerboseMode, _OA_Read("Logging", "VerboseMode", "0"))

	Local $btnClean = GUICtrlCreateButton("Clean Outlook WorkDays items...", 18, 575, 210, 30)
	Local $btnSave = GUICtrlCreateButton("Close", 838, 575, 100, 30)

	GUISetState(@SW_SHOW, $hAgent)

	While 1
		Local $nAgentMsg = GUIGetMsg()
		Switch $nAgentMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($hAgent)
				$g_hOutlookAgentSettingsWindow = 0
				WinActivate("Work Days")
				Return 0

			Case $btnInstall
				_OutlookAgent_Install()
				GUICtrlSetData($lblStatus, _OutlookAgent_StatusText())

			Case $btnStart
				_OutlookAgent_Start()
				GUICtrlSetData($lblStatus, _OutlookAgent_StatusText())

			Case $btnStop
				_OutlookAgent_Stop()
				GUICtrlSetData($lblStatus, _OutlookAgent_StatusText())

			Case $btnOpenLog
				_OutlookAgent_OpenLog()

			Case $btnUninstall
				_OutlookAgent_Uninstall()
				GUICtrlSetData($lblStatus, _OutlookAgent_StatusText())

			Case $btnOpenBackupFolder
				DirCreate($g_sOutlookAgentDir & "\Backup")
				ShellExecute($g_sOutlookAgentDir & "\Backup")

			Case $btnOpenPlan
				Local $sPlan = RegRead($g_sOutlookAgentDB, "LastSyncPlanFile")
				If @error Or $sPlan = "" Or Not FileExists($sPlan) Then
					MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), "WorkDays Outlook Agent", "The last sync plan file has not been created yet.", 0, $hAgent)
				Else
					ShellExecute($sPlan)
				EndIf

			Case $btnClean
				_OutlookAgent_SaveSettings($inpInterval, $inpPast, $inpFuture, $chkConflictOutlook, $chkDeleteOutlookClears, $chkSyncBlank, $chkSyncWeekend, $chkSyncTaggedBlankWeekend, $chkStartup, $inpSubjectPrefix, $inpCategoryPrefix, $chkReminderSet, $chkManagedOnly, $chkShowMarkerSubject, $inpMarkerSuffix, $chkSeparateMarkerCategory, $inpMarkerCategory, $chkReminderMarker, $inpReminderMinutes, $chkCleanupEnabled, $inpCleanupPast, $inpCleanupFuture, $chkCleanupPrefixOnly, $chkPauseAfterCleanup, $chkVerboseMode, $chkBackupBeforeOutlookChanges, $chkBlockMassChanges, $inpMaxChanges, $inpMaxPercent, $inpMaxClears, $chkBlockIncompleteRead)
				_OutlookAgent_CleanOutlookFromWorkDays()
				GUICtrlSetData($lblStatus, _OutlookAgent_StatusText())

			Case $btnSave
				_OutlookAgent_SaveSettings($inpInterval, $inpPast, $inpFuture, $chkConflictOutlook, $chkDeleteOutlookClears, $chkSyncBlank, $chkSyncWeekend, $chkSyncTaggedBlankWeekend, $chkStartup, $inpSubjectPrefix, $inpCategoryPrefix, $chkReminderSet, $chkManagedOnly, $chkShowMarkerSubject, $inpMarkerSuffix, $chkSeparateMarkerCategory, $inpMarkerCategory, $chkReminderMarker, $inpReminderMinutes, $chkCleanupEnabled, $inpCleanupPast, $inpCleanupFuture, $chkCleanupPrefixOnly, $chkPauseAfterCleanup, $chkVerboseMode, $chkBackupBeforeOutlookChanges, $chkBlockMassChanges, $inpMaxChanges, $inpMaxPercent, $inpMaxClears, $chkBlockIncompleteRead)
				GUIDelete($hAgent)
				$g_hOutlookAgentSettingsWindow = 0
				WinActivate("Work Days")
				Return 1
		EndSwitch
	WEnd
EndFunc   ;==>_OutlookAgent_SettingsWindow
