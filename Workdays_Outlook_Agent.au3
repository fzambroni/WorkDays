#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Icon=CalendarSync.ico
#AutoIt3Wrapper_Res_Description=Work Day Sync Agent
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_ProductName=Work Day Sync Agent
#AutoIt3Wrapper_Res_CompanyName=Fabricio Zambroni
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2026 Fabricio Zambroni
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <TrayConstants.au3>

Opt("MustDeclareVars", 1)
Opt("TrayMenuMode", 3)
Opt("TrayOnEventMode", 0)

Global Const $g_sAppTitle = "WorkDays Outlook Agent"
Global Const $g_sDB = "HKEY_CURRENT_USER\Software\WorkDays"
Global Const $g_sAgentDB = "HKEY_CURRENT_USER\Software\WorkDays\OutlookAgent"
Global Const $g_sAgentDir = @LocalAppDataDir & "\WorkDays"
Global Const $g_sState = $g_sAgentDir & "\Workdays_Outlook_Agent_State.ini"
Global Const $g_sLog = $g_sAgentDir & "\Workdays_Outlook_Agent.log"
Global Const $g_sSep = Chr(29)

Global Const $OL_APPOINTMENT_ITEM = 1
Global Const $OL_FOLDER_CALENDAR = 9
Global Const $OL_FREE = 0
Global Const $OL_TEXT = 1

Global $g_oComError = ObjEvent("AutoIt.Error", "_ComErrorHandler")
Global $g_bPaused = False
Global $g_hTimer = TimerInit()
Global $g_iTrayStatus = 0
Global $g_iTraySyncNow = 0
Global $g_iTrayPause = 0
Global $g_iTrayLog = 0
Global $g_iTrayCleanOutlook = 0
Global $g_iTrayExit = 0

DirCreate($g_sAgentDir)
_EnsureConfig()
_HandleCommandLine()

If _Singleton($g_sAppTitle, 1) = 0 Then
	MsgBox($MB_ICONINFORMATION, $g_sAppTitle, "WorkDays Outlook Agent is already running.")
	Exit
EndIf

_ApplyStartupSetting()
_CreateTray()
_Log("Agent started.")
_SyncNow()
$g_hTimer = TimerInit()

While 1
	Local $iMsg = TrayGetMsg()
	Switch $iMsg
		Case $g_iTraySyncNow
			_SyncNow()
			$g_hTimer = TimerInit()
		Case $g_iTrayPause
			_TogglePause()
		Case $g_iTrayLog
			If FileExists($g_sLog) Then
				ShellExecute($g_sLog)
			Else
				MsgBox($MB_ICONINFORMATION, $g_sAppTitle, "The log file has not been created yet.")
			EndIf
		Case $g_iTrayCleanOutlook
			_CleanOutlookCalendarFromTray()
		Case $g_iTrayExit
			_Log("Agent closed by user.")
			Exit
	EndSwitch

	If Not $g_bPaused Then
		Local $iIntervalMin = Number(_Cfg("Sync", "IntervalMinutes", "15"))
		If $iIntervalMin < 1 Then $iIntervalMin = 15
		If TimerDiff($g_hTimer) >= ($iIntervalMin * 60000) Then
			_SyncNow()
			$g_hTimer = TimerInit()
		EndIf
	EndIf

	Sleep(250)
WEnd

Func _HandleCommandLine()
	If $CmdLine[0] < 1 Then Return
	Local $sCmd = StringLower(StringStripWS($CmdLine[1], 3))
	Switch $sCmd
		Case "/cleanoutlook"
			_Log("Outlook cleanup requested by WorkDays.")
			Local $iDeleted = _CleanOutlookCalendar()
			If @error Then
				_Log("Outlook cleanup failed. Error code: " & @error)
				Exit 1
			EndIf
			_Log("Outlook cleanup completed. Deleted items: " & $iDeleted)
			Exit 0
		Case "/synconce", "/syncnow"
			_Log("One-time sync requested by WorkDays.")
			Local $iChanges = _RunSync()
			If @error Then
				_Log("One-time sync failed. Error code: " & @error)
				Exit 1
			EndIf
			_Log("One-time sync completed. Changes: " & $iChanges)
			Exit 0
	EndSwitch
EndFunc

Func _EnsureConfig()
	; Settings are owned by the main WorkDays application and stored in the registry.
	; The agent only creates missing defaults for upgrade safety.
	_EnsureRegDefault("Sync", "IntervalMinutes", "15")
	_EnsureRegDefault("Sync", "PastDays", "60")
	_EnsureRegDefault("Sync", "FutureDays", "370")
	_EnsureRegDefault("Sync", "OutlookWinsOnConflict", "1")
	_EnsureRegDefault("Sync", "DeleteInOutlookClearsWorkDays", "0")
	_EnsureRegDefault("Sync", "SyncBlank", "0")
	_EnsureRegDefault("Sync", "SyncWeekend", "0")
	_EnsureRegDefault("Sync", "SyncTaggedBlankOrWeekend", "1")
	_EnsureRegDefault("Sync", "RunAtWindowsStartup", "0")

	_EnsureRegDefault("Outlook", "SubjectPrefix", "WorkDays -")
	_EnsureRegDefault("Outlook", "CategoryPrefix", "WorkDays -")
	_EnsureRegDefault("Outlook", "ReminderSet", "0")
	_EnsureRegDefault("Outlook", "ManagedOnly", "0")

	_EnsureRegDefault("Markers", "ShowMarkerTagInSubject", "1")
	_EnsureRegDefault("Markers", "MarkerSubjectSuffix", " [Marker]")
	_EnsureRegDefault("Markers", "UseSeparateMarkerCategory", "1")
	_EnsureRegDefault("Markers", "MarkerCategoryName", "WorkDays - Marker")
	_EnsureRegDefault("Markers", "ReminderWhenMarkerExists", "1")
	_EnsureRegDefault("Markers", "ReminderMinutesBeforeStart", "540")

	_EnsureRegDefault("Safety", "EnableOutlookCleanup", "1")
	_EnsureRegDefault("Safety", "CleanupPastYears", "10")
	_EnsureRegDefault("Safety", "CleanupFutureYears", "10")
	_EnsureRegDefault("Safety", "CleanupPrefixOnlyItems", "0")
	_EnsureRegDefault("Safety", "PauseAfterOutlookCleanup", "1")
	_EnsureRegDefault("Safety", "CleanupConfirmationPhrase", "CLEAN WORKDAYS OUTLOOK")

	_EnsureRegDefault("Advanced", "LogLevel", "Normal")
EndFunc

Func _SettingName($sSection, $sKey)
	Return $sSection & "_" & $sKey
EndFunc

Func _EnsureRegDefault($sSection, $sKey, $sDefault)
	RegRead($g_sAgentDB, _SettingName($sSection, $sKey))
	If @error Then RegWrite($g_sAgentDB, _SettingName($sSection, $sKey), "REG_SZ", String($sDefault))
EndFunc

Func _Cfg($sSection, $sKey, $sDefault)
	Local $sValue = RegRead($g_sAgentDB, _SettingName($sSection, $sKey))
	If @error Or $sValue = "" Then Return $sDefault
	Return String($sValue)
EndFunc

Func _SetCfg($sSection, $sKey, $sValue)
	Return RegWrite($g_sAgentDB, _SettingName($sSection, $sKey), "REG_SZ", String($sValue))
EndFunc

Func _CreateTray()
	TraySetToolTip($g_sAppTitle)
	$g_iTrayStatus = TrayCreateItem("Last sync: starting...")
	TrayItemSetState($g_iTrayStatus, $TRAY_DISABLE)
	TrayCreateItem("")
	$g_iTraySyncNow = TrayCreateItem("Sync now")
	$g_iTrayPause = TrayCreateItem("Pause sync")
	$g_iTrayLog = TrayCreateItem("Open log")
	TrayCreateItem("")
	$g_iTrayCleanOutlook = TrayCreateItem("Clean Outlook WorkDays items...")
	TrayCreateItem("")
	$g_iTrayExit = TrayCreateItem("Exit")
	TraySetState($TRAY_ICONSTATE_SHOW)
EndFunc

Func _TogglePause()
	$g_bPaused = Not $g_bPaused
	If $g_bPaused Then
		TrayItemSetText($g_iTrayPause, "Resume sync")
		TrayItemSetText($g_iTrayStatus, "Paused")
		TraySetToolTip($g_sAppTitle & " - paused")
		_Log("Sync paused.")
	Else
		TrayItemSetText($g_iTrayPause, "Pause sync")
		TraySetToolTip($g_sAppTitle)
		_Log("Sync resumed.")
		_SyncNow()
		$g_hTimer = TimerInit()
	EndIf
EndFunc

Func _ApplyStartupSetting()
	; Startup is configured only by the main WorkDays application.
	; The agent only applies the registry value it receives from WorkDays.
	If _Cfg("Sync", "RunAtWindowsStartup", "0") = "1" Then
		RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "WorkDays Outlook Agent", "REG_SZ", _AgentRunCommand())
	Else
		RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "WorkDays Outlook Agent")
	EndIf
EndFunc

Func _AgentRunCommand()
	If @Compiled Then Return '"' & @ScriptFullPath & '"'
	Return '"' & @AutoItExe & '" "' & @ScriptFullPath & '"'
EndFunc

Func _SyncNow()
	TrayItemSetText($g_iTrayStatus, "Syncing...")
	TraySetToolTip($g_sAppTitle & " - syncing")

	Local $iChanges = _RunSync()
	Local $sNow = StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC)

	If @error Then
		TrayItemSetText($g_iTrayStatus, "Last sync failed: " & @HOUR & ":" & @MIN)
		TraySetToolTip($g_sAppTitle & " - last sync failed")
		_Log("Sync failed. Error code: " & @error)
	Else
		TrayItemSetText($g_iTrayStatus, "Last sync: " & @HOUR & ":" & @MIN & " | Changes: " & $iChanges)
		TraySetToolTip($g_sAppTitle & " - last sync " & $sNow)
		IniWrite($g_sState, "Global", "LastSync", $sNow)
		_Log("Sync completed. Changes: " & $iChanges)
	EndIf
EndFunc

Func _RunSync()
	Local $oOutlook = _GetOutlookApplication()
	If Not IsObj($oOutlook) Then Return SetError(1, 0, 0)

	Local $oNs = $oOutlook.GetNamespace("MAPI")
	If Not IsObj($oNs) Then Return SetError(2, 0, 0)

	Local $oCalendar = $oNs.GetDefaultFolder($OL_FOLDER_CALENDAR)
	If Not IsObj($oCalendar) Then Return SetError(3, 0, 0)

	Local $iPast = Number(_Cfg("Sync", "PastDays", "60"))
	Local $iFuture = Number(_Cfg("Sync", "FutureDays", "370"))
	If $iPast < 0 Then $iPast = 60
	If $iFuture < 1 Then $iFuture = 370

	Local $sStartISO = _ISOAddDays(_TodayISO(), -$iPast)
	Local $sEndISO = _ISOAddDays(_TodayISO(), $iFuture)
	Local $oOutlookMap = _LoadOutlookMap($oCalendar, $sStartISO, $sEndISO)
	If Not IsObj($oOutlookMap) Then Return SetError(4, 0, 0)

	Local $iChanges = 0
	Local $iDays = _ISODiffDays($sStartISO, $sEndISO)
	Local $i

	For $i = 0 To $iDays
		Local $sDateISO = _ISOAddDays($sStartISO, $i)
		Local $sRegRec = _ReadRegistryDay($sDateISO)
		Local $sRegStatus = _RecordStatus($sRegRec)
		Local $sRegMarker = _RecordMarker($sRegRec)
		Local $sRegHash = _RecordHash($sRegStatus, $sRegMarker)

		Local $sOutRec = ""
		Local $bHasOutlook = False
		If $oOutlookMap.Exists($sDateISO) Then
			$sOutRec = $oOutlookMap.Item($sDateISO)
			$bHasOutlook = True
		EndIf

		Local $sOutEntryID = _OutlookRecordPart($sOutRec, 0)
		Local $sOutStatus = _OutlookRecordPart($sOutRec, 1)
		Local $sOutMarker = _OutlookRecordPart($sOutRec, 2)
		Local $sOutHash = ""
		If $bHasOutlook Then $sOutHash = _RecordHash($sOutStatus, $sOutMarker)

		Local $sStateRegHash = IniRead($g_sState, $sDateISO, "RegHash", "")
		Local $sStateOutHash = IniRead($g_sState, $sDateISO, "OutHash", "")
		Local $sStateEntryID = IniRead($g_sState, $sDateISO, "EntryID", "")

		Local $bRegChanged = ($sRegHash <> $sStateRegHash)
		Local $bOutChanged = ($sOutHash <> $sStateOutHash)

		If Not $bHasOutlook And $sStateOutHash <> "" And Not $bRegChanged Then
			; The Outlook item disappeared after being synced. By default this is kept safe and ignored.
			If _Cfg("Sync", "DeleteInOutlookClearsWorkDays", "0") = "1" Then
				_WriteRegistryDay($sDateISO, "B", "")
				_UpdateState($sDateISO, _RecordHash("B", ""), "", "")
				$iChanges += 1
				_Log("Outlook deletion cleared WorkDays date " & $sDateISO)
			Else
				If _ShouldSync($sRegStatus, $sRegMarker) Then
					$sOutEntryID = _CreateOrUpdateOutlookItem($oOutlook, $oNs, $sDateISO, "", $sRegStatus, $sRegMarker)
					_UpdateState($sDateISO, $sRegHash, $sRegHash, $sOutEntryID)
					$iChanges += 1
					_Log("Re-created missing Outlook item for " & $sDateISO)
				EndIf
			EndIf
			ContinueLoop
		EndIf

		If $bHasOutlook And $bOutChanged And Not $bRegChanged Then
			; Outlook changed after the last sync. Pull it into Work Days.
			_WriteRegistryDay($sDateISO, $sOutStatus, $sOutMarker)
			_UpdateState($sDateISO, $sOutHash, $sOutHash, $sOutEntryID)
			_EnsureOutlookItemFree($oNs, $sOutEntryID, $sDateISO, $sOutStatus, $sOutMarker)
			$iChanges += 1
			_Log("Pulled Outlook change into WorkDays: " & $sDateISO & " -> " & _StatusLabel($sOutStatus))
			ContinueLoop
		EndIf

		If $bRegChanged And Not $bOutChanged Then
			; Work Days changed after the last sync. Push it into Outlook.
			If _ShouldSync($sRegStatus, $sRegMarker) Then
				If $sOutEntryID = "" Then $sOutEntryID = $sStateEntryID
				$sOutEntryID = _CreateOrUpdateOutlookItem($oOutlook, $oNs, $sDateISO, $sOutEntryID, $sRegStatus, $sRegMarker)
				_UpdateState($sDateISO, $sRegHash, $sRegHash, $sOutEntryID)
				$iChanges += 1
				_Log("Pushed WorkDays change into Outlook: " & $sDateISO & " -> " & _StatusLabel($sRegStatus))
			Else
				If $bHasOutlook Then
					_DeleteOutlookItem($oNs, $sOutEntryID)
					$iChanges += 1
					_Log("Removed Outlook item because WorkDays date is not configured to sync: " & $sDateISO)
				EndIf
				_UpdateState($sDateISO, $sRegHash, "", "")
			EndIf
			ContinueLoop
		EndIf

		If $bRegChanged And $bOutChanged Then
			If _Cfg("Sync", "OutlookWinsOnConflict", "1") = "1" And $bHasOutlook Then
				_WriteRegistryDay($sDateISO, $sOutStatus, $sOutMarker)
				_UpdateState($sDateISO, $sOutHash, $sOutHash, $sOutEntryID)
				_EnsureOutlookItemFree($oNs, $sOutEntryID, $sDateISO, $sOutStatus, $sOutMarker)
				$iChanges += 1
				_Log("Conflict resolved by Outlook: " & $sDateISO)
			Else
				If _ShouldSync($sRegStatus, $sRegMarker) Then
					If $sOutEntryID = "" Then $sOutEntryID = $sStateEntryID
					$sOutEntryID = _CreateOrUpdateOutlookItem($oOutlook, $oNs, $sDateISO, $sOutEntryID, $sRegStatus, $sRegMarker)
					_UpdateState($sDateISO, $sRegHash, $sRegHash, $sOutEntryID)
					$iChanges += 1
					_Log("Conflict resolved by WorkDays: " & $sDateISO)
				Else
					_UpdateState($sDateISO, $sRegHash, "", "")
				EndIf
			EndIf
			ContinueLoop
		EndIf

		; First run or state repair: create missing Outlook items from existing Work Days data.
		; If Outlook was intentionally cleaned, the state keeps the current WorkDays hash
		; with no Outlook hash. In that case, do not republish until WorkDays changes again.
		If Not $bHasOutlook And _ShouldSync($sRegStatus, $sRegMarker) And $sStateRegHash = "" And $sStateOutHash = "" And $sStateEntryID = "" Then
			$sOutEntryID = _CreateOrUpdateOutlookItem($oOutlook, $oNs, $sDateISO, "", $sRegStatus, $sRegMarker)
			_UpdateState($sDateISO, $sRegHash, $sRegHash, $sOutEntryID)
			$iChanges += 1
			_Log("Created Outlook item for existing WorkDays date: " & $sDateISO)
			ContinueLoop
		EndIf

		If $bHasOutlook Then
			_EnsureOutlookItemFree($oNs, $sOutEntryID, $sDateISO, $sOutStatus, $sOutMarker)
			_UpdateState($sDateISO, $sOutHash, $sOutHash, $sOutEntryID)
		EndIf
	Next

	Return $iChanges
EndFunc

Func _CleanOutlookCalendarFromTray()
	If _Cfg("Safety", "EnableOutlookCleanup", "1") <> "1" Then
		MsgBox($MB_ICONINFORMATION, $g_sAppTitle, "Outlook cleanup is disabled in settings.")
		Return
	EndIf

	Local $sPhrase = _Cfg("Safety", "CleanupConfirmationPhrase", "CLEAN WORKDAYS OUTLOOK")
	Local $sMsg = "This will delete WorkDays calendar items from Outlook only." & @CRLF & @CRLF & _
		"Your WorkDays data will remain stored in the WorkDays app." & @CRLF & _
		"After cleanup, sync will be paused if PauseAfterOutlookCleanup is enabled." & @CRLF & @CRLF & _
		"Continue?"
	If MsgBox(BitOR($MB_ICONWARNING, $MB_YESNO, $MB_DEFBUTTON2), $g_sAppTitle, $sMsg) <> $IDYES Then Return

	Local $sTyped = InputBox($g_sAppTitle, "Type exactly this confirmation phrase:" & @CRLF & @CRLF & $sPhrase, "", "", 420, 150)
	If @error Then Return
	If StringStripWS($sTyped, 3) <> $sPhrase Then
		MsgBox($MB_ICONINFORMATION, $g_sAppTitle, "Cleanup cancelled. The confirmation phrase did not match.")
		Return
	EndIf

	TrayItemSetText($g_iTrayStatus, "Cleaning Outlook...")
	Local $iDeleted = _CleanOutlookCalendar()
	If @error Then
		MsgBox($MB_ICONERROR, $g_sAppTitle, "Outlook cleanup failed. Check the log file for details.")
		TrayItemSetText($g_iTrayStatus, "Cleanup failed")
		Return
	EndIf

	If _Cfg("Safety", "PauseAfterOutlookCleanup", "1") = "1" And Not $g_bPaused Then
		_TogglePause()
	EndIf

	MsgBox($MB_ICONINFORMATION, $g_sAppTitle, "Outlook cleanup completed." & @CRLF & @CRLF & "Deleted items: " & $iDeleted)
	_Log("Outlook cleanup completed. Deleted items: " & $iDeleted)
EndFunc

Func _CleanOutlookCalendar()
	Local $oOutlook = _GetOutlookApplication()
	If Not IsObj($oOutlook) Then Return SetError(1, 0, 0)

	Local $oNs = $oOutlook.GetNamespace("MAPI")
	If Not IsObj($oNs) Then Return SetError(2, 0, 0)

	Local $oCalendar = $oNs.GetDefaultFolder($OL_FOLDER_CALENDAR)
	If Not IsObj($oCalendar) Then Return SetError(3, 0, 0)

	Local $iPastYears = Number(_Cfg("Safety", "CleanupPastYears", "10"))
	Local $iFutureYears = Number(_Cfg("Safety", "CleanupFutureYears", "10"))
	If $iPastYears < 0 Then $iPastYears = 10
	If $iFutureYears < 0 Then $iFutureYears = 10

	Local $sStartISO = _ISOAddDays(_TodayISO(), -365 * $iPastYears)
	Local $sEndISO = _ISOAddDays(_TodayISO(), 365 * $iFutureYears)
	Local $oItems = $oCalendar.Items
	If Not IsObj($oItems) Then Return SetError(4, 0, 0)

	$oItems.IncludeRecurrences = False
	$oItems.Sort("[Start]")
	Local $sFilter = "[Start] >= '" & _OutlookFilterDate($sStartISO) & "' AND [Start] < '" & _OutlookFilterDate(_ISOAddDays($sEndISO, 1)) & "'"
	Local $oRange = $oItems.Restrict($sFilter)
	If Not IsObj($oRange) Then Return SetError(5, 0, 0)

	Local $oDeleteList = ObjCreate("Scripting.Dictionary")
	If Not IsObj($oDeleteList) Then Return SetError(6, 0, 0)

	Local $oItem
	For $oItem In $oRange
		If Not IsObj($oItem) Then ContinueLoop
		If Not _IsOutlookCleanupCandidate($oItem) Then ContinueLoop

		Local $sEntryID = String($oItem.EntryID)
		If $sEntryID = "" Then ContinueLoop
		Local $sDateISO = _GetUserProp($oItem, "WorkDaysDate")
		If Not _IsISODate($sDateISO) Then $sDateISO = _OutlookDateToISO($oItem.Start)
		If Not $oDeleteList.Exists($sEntryID) Then $oDeleteList.Add($sEntryID, $sDateISO)
	Next

	Local $iDeleted = 0
	Local $vEntryID
	For $vEntryID In $oDeleteList.Keys
		Local $sDate = $oDeleteList.Item($vEntryID)
		If _DeleteOutlookItem($oNs, String($vEntryID)) Then
			$iDeleted += 1
			If _IsISODate($sDate) Then _MarkDateAsOutlookCleaned($sDate)
		EndIf
	Next

	Return $iDeleted
EndFunc

Func _IsOutlookCleanupCandidate($oItem)
	If Not IsObj($oItem) Then Return False
	If _GetUserProp($oItem, "WorkDaysManaged") = "1" Then Return True

	If _Cfg("Safety", "CleanupPrefixOnlyItems", "0") <> "1" Then Return False

	Local $sSubject = String($oItem.Subject)
	Local $sPrefix = _SubjectPrefix()
	If StringLeft(StringLower(StringStripWS($sSubject, 3)), StringLen(StringLower($sPrefix))) = StringLower($sPrefix) Then Return True
	If StringRegExp($sSubject, "(?i)^\s*\[\s*WD\s*[:\-]\s*[A-Z]") Then Return True
	Return False
EndFunc

Func _MarkDateAsOutlookCleaned($sDateISO)
	Local $sRegRec = _ReadRegistryDay($sDateISO)
	Local $sRegStatus = _RecordStatus($sRegRec)
	Local $sRegMarker = _RecordMarker($sRegRec)
	_UpdateState($sDateISO, _RecordHash($sRegStatus, $sRegMarker), "", "")
EndFunc

Func _GetOutlookApplication()
	Local $oOutlook = ObjGet("", "Outlook.Application")
	If Not IsObj($oOutlook) Then $oOutlook = ObjCreate("Outlook.Application")
	If Not IsObj($oOutlook) Then Return SetError(1, 0, 0)
	Return $oOutlook
EndFunc

Func _LoadOutlookMap($oCalendar, $sStartISO, $sEndISO)
	Local $oMap = ObjCreate("Scripting.Dictionary")
	If Not IsObj($oMap) Then Return SetError(1, 0, 0)

	Local $oItems = $oCalendar.Items
	If Not IsObj($oItems) Then Return $oMap

	$oItems.IncludeRecurrences = True
	$oItems.Sort("[Start]")

	Local $sFilter = "[Start] >= '" & _OutlookFilterDate($sStartISO) & "' AND [Start] < '" & _OutlookFilterDate(_ISOAddDays($sEndISO, 1)) & "'"
	Local $oRange = $oItems.Restrict($sFilter)
	If Not IsObj($oRange) Then Return $oMap
	Local $oItem

	For $oItem In $oRange
		If Not IsObj($oItem) Then ContinueLoop
		If Not _IsWorkDaysCandidate($oItem) Then ContinueLoop

		Local $sDateISO = _GetUserProp($oItem, "WorkDaysDate")
		If Not _IsISODate($sDateISO) Then $sDateISO = _OutlookDateToISO($oItem.Start)
		If Not _IsISODate($sDateISO) Then ContinueLoop

		; Visible Outlook data wins. This lets users change the day classification directly in Outlook.
		; Internal UserProperties are kept only as a fallback for old/managed items.
		Local $sStatus = _GetOutlookItemStatus($oItem)
		If Not _IsKnownStatus($sStatus) Then ContinueLoop

		Local $sMarker = _CleanOutlookMarker($oItem.Body)
		Local $sEntryID = String($oItem.EntryID)
		Local $sManaged = _GetUserProp($oItem, "WorkDaysManaged")
		Local $sRec = $sEntryID & $g_sSep & $sStatus & $g_sSep & $sMarker & $g_sSep & $sManaged

		If $oMap.Exists($sDateISO) Then
			; Prefer managed items over manually created prefix-only items.
			Local $sExisting = $oMap.Item($sDateISO)
			Local $sExistingManaged = _OutlookRecordPart($sExisting, 3)
			If $sExistingManaged = "1" And $sManaged <> "1" Then ContinueLoop
			_Log("Duplicate WorkDays Outlook item detected for " & $sDateISO & ". The agent will use the most recent candidate it found.")
			$oMap.Item($sDateISO) = $sRec
		Else
			$oMap.Add($sDateISO, $sRec)
		EndIf
	Next

	Return $oMap
EndFunc

Func _IsWorkDaysCandidate($oItem)
	If Not IsObj($oItem) Then Return False
	If _GetUserProp($oItem, "WorkDaysManaged") = "1" Then Return True

	If _Cfg("Outlook", "ManagedOnly", "0") = "1" Then Return False

	If _IsKnownStatus(_ParseStatusFromCategories($oItem.Categories)) Then Return True

	Local $sSubject = String($oItem.Subject)
	Local $sPrefix = _SubjectPrefix()
	If StringLeft(StringLower(StringStripWS($sSubject, 3)), StringLen(StringLower($sPrefix))) = StringLower($sPrefix) Then Return True
	If StringRegExp($sSubject, "(?i)^\s*\[\s*WD\s*[:\-]\s*[A-Z]") Then Return True
	Return False
EndFunc

Func _CreateOrUpdateOutlookItem($oOutlook, $oNs, $sDateISO, $sEntryID, $sStatus, $sMarker)
	Local $oItem = 0
	If $sEntryID <> "" Then $oItem = _GetOutlookItemByEntryID($oNs, $sEntryID)
	If Not IsObj($oItem) Then $oItem = $oOutlook.CreateItem($OL_APPOINTMENT_ITEM)
	If Not IsObj($oItem) Then Return ""

	$oItem.Subject = _BuildSubject($sStatus, $sMarker)
	$oItem.Start = _OutlookFilterDate($sDateISO)
	$oItem.End = _OutlookFilterDate(_ISOAddDays($sDateISO, 1))
	$oItem.AllDayEvent = True
	$oItem.BusyStatus = $OL_FREE
	_SetOutlookReminder($oItem, $sMarker)
	$oItem.Body = $sMarker
	$oItem.Categories = _BuildCategories($sStatus, $sMarker)
	_SetUserProp($oItem, "WorkDaysManaged", "1")
	_SetUserProp($oItem, "WorkDaysDate", $sDateISO)
	_SetUserProp($oItem, "WorkDaysStatus", $sStatus)
	_SetUserProp($oItem, "WorkDaysHash", _RecordHash($sStatus, $sMarker))
	$oItem.Save()

	Return String($oItem.EntryID)
EndFunc

Func _EnsureOutlookItemFree($oNs, $sEntryID, $sDateISO, $sStatus, $sMarker)
	If $sEntryID = "" Then Return 0
	Local $oItem = _GetOutlookItemByEntryID($oNs, $sEntryID)
	If Not IsObj($oItem) Then Return 0

	Local $bChanged = False
	If String($oItem.Subject) <> _BuildSubject($sStatus, $sMarker) Then
		$oItem.Subject = _BuildSubject($sStatus, $sMarker)
		$bChanged = True
	EndIf
	If $oItem.BusyStatus <> $OL_FREE Then
		$oItem.BusyStatus = $OL_FREE
		$bChanged = True
	EndIf
	If Not $oItem.AllDayEvent Then
		$oItem.AllDayEvent = True
		$bChanged = True
	EndIf
	If String($oItem.Categories) <> _BuildCategories($sStatus, $sMarker) Then
		$oItem.Categories = _BuildCategories($sStatus, $sMarker)
		$bChanged = True
	EndIf
	If String($oItem.Body) <> $sMarker Then
		$oItem.Body = $sMarker
		$bChanged = True
	EndIf
	If _SetOutlookReminder($oItem, $sMarker) Then $bChanged = True
	_SetUserProp($oItem, "WorkDaysManaged", "1")
	_SetUserProp($oItem, "WorkDaysDate", $sDateISO)
	_SetUserProp($oItem, "WorkDaysStatus", $sStatus)
	_SetUserProp($oItem, "WorkDaysHash", _RecordHash($sStatus, $sMarker))
	If $bChanged Then $oItem.Save()
	Return 1
EndFunc

Func _DeleteOutlookItem($oNs, $sEntryID)
	If $sEntryID = "" Then Return 0
	Local $oItem = _GetOutlookItemByEntryID($oNs, $sEntryID)
	If Not IsObj($oItem) Then Return 0
	$oItem.Delete()
	Return 1
EndFunc

Func _GetOutlookItemByEntryID($oNs, $sEntryID)
	If $sEntryID = "" Then Return 0
	Local $oItem = $oNs.GetItemFromID($sEntryID)
	If Not IsObj($oItem) Then Return 0
	Return $oItem
EndFunc

Func _ReadRegistryDay($sDateISO)
	Local $a = StringSplit($sDateISO, "-")
	If @error Or $a[0] <> 3 Then Return ""
	Local $sYear = $a[1]
	Local $sMonth = $a[2]
	Local $sDay = $a[3]
	Local $sValue = RegRead($g_sDB & "\" & $sYear & "\" & $sMonth, $sDay)
	If @error Then Return ""
	Return String($sValue)
EndFunc

Func _WriteRegistryDay($sDateISO, $sStatus, $sMarker)
	If Not _IsKnownStatus($sStatus) Then Return 0
	Local $a = StringSplit($sDateISO, "-")
	If @error Or $a[0] <> 3 Then Return 0
	RegWrite($g_sDB & "\" & $a[1] & "\" & $a[2], $a[3], "REG_SZ", $sStatus & $sMarker)
	Return 1
EndFunc

Func _ShouldSync($sStatus, $sMarker)
	If Not _IsKnownStatus($sStatus) Then Return False
	If $sStatus = "B" Then
		If _Cfg("Sync", "SyncBlank", "0") = "1" Then Return True
		If StringStripWS($sMarker, 3) <> "" And _Cfg("Sync", "SyncTaggedBlankOrWeekend", "1") = "1" Then Return True
		Return False
	EndIf
	If $sStatus = "W" Then
		If _Cfg("Sync", "SyncWeekend", "0") = "1" Then Return True
		If StringStripWS($sMarker, 3) <> "" And _Cfg("Sync", "SyncTaggedBlankOrWeekend", "1") = "1" Then Return True
		Return False
	EndIf
	Return True
EndFunc

Func _UpdateState($sDateISO, $sRegHash, $sOutHash, $sEntryID)
	IniWrite($g_sState, $sDateISO, "RegHash", $sRegHash)
	IniWrite($g_sState, $sDateISO, "OutHash", $sOutHash)
	IniWrite($g_sState, $sDateISO, "EntryID", $sEntryID)
EndFunc

Func _RecordStatus($sRecord)
	If StringLen($sRecord) < 1 Then Return ""
	Local $sStatus = StringUpper(StringLeft($sRecord, 1))
	If _IsKnownStatus($sStatus) Then Return $sStatus
	Return ""
EndFunc

Func _RecordMarker($sRecord)
	If StringLen($sRecord) <= 1 Then Return ""
	Return StringTrimLeft($sRecord, 1)
EndFunc

Func _RecordHash($sStatus, $sMarker)
	If $sStatus = "" And $sMarker = "" Then Return ""
	Return _HashString($sStatus & "|" & $sMarker)
EndFunc

Func _HashString($sValue)
	Local $iHash = 5381
	Local $i
	For $i = 1 To StringLen($sValue)
		$iHash = BitAND(($iHash * 33) + AscW(StringMid($sValue, $i, 1)), 0x7FFFFFFF)
	Next
	Return Hex($iHash, 8)
EndFunc

Func _OutlookRecordPart($sRecord, $iIndex)
	If $sRecord = "" Then Return ""
	Local $a = StringSplit($sRecord, $g_sSep, 2)
	If @error Or Not IsArray($a) Then Return ""
	If $iIndex < 0 Or $iIndex > UBound($a) - 1 Then Return ""
	Return $a[$iIndex]
EndFunc


Func _SubjectPrefix()
	Local $sPrefix = _Cfg("Outlook", "SubjectPrefix", "WorkDays -")
	If StringRight($sPrefix, 1) = "-" Then $sPrefix &= " "
	Return $sPrefix
EndFunc

Func _CategoryPrefix()
	Local $sPrefix = _Cfg("Outlook", "CategoryPrefix", "WorkDays -")
	If StringRight($sPrefix, 1) = "-" Then $sPrefix &= " "
	Return $sPrefix
EndFunc

Func _BuildSubject($sStatus, $sMarker = "")
	Local $sSubject = _SubjectPrefix() & _StatusLabel($sStatus)
	If _HasMarker($sMarker) And _Cfg("Markers", "ShowMarkerTagInSubject", "1") = "1" Then
		$sSubject &= _Cfg("Markers", "MarkerSubjectSuffix", " [Marker]")
	EndIf
	Return $sSubject
EndFunc

Func _BuildCategories($sStatus, $sMarker = "")
	Local $sCategories = _CategoryPrefix() & _StatusLabel($sStatus)
	If _HasMarker($sMarker) And _Cfg("Markers", "UseSeparateMarkerCategory", "1") = "1" Then
		$sCategories &= ", " & _Cfg("Markers", "MarkerCategoryName", "WorkDays - Marker")
	EndIf
	Return $sCategories
EndFunc

Func _HasMarker($sMarker)
	Return StringStripWS(String($sMarker), 3) <> ""
EndFunc

Func _SetOutlookReminder(ByRef $oItem, $sMarker)
	If Not IsObj($oItem) Then Return False
	Local $bChanged = False
	Local $bReminderSet = (_Cfg("Outlook", "ReminderSet", "0") = "1")
	If _HasMarker($sMarker) And _Cfg("Markers", "ReminderWhenMarkerExists", "1") = "1" Then $bReminderSet = True

	If $oItem.ReminderSet <> $bReminderSet Then
		$oItem.ReminderSet = $bReminderSet
		$bChanged = True
	EndIf

	If $bReminderSet Then
		Local $iMinutes = Number(_Cfg("Markers", "ReminderMinutesBeforeStart", "540"))
		If $iMinutes < 0 Then $iMinutes = 0
		If $oItem.ReminderMinutesBeforeStart <> $iMinutes Then
			$oItem.ReminderMinutesBeforeStart = $iMinutes
			$bChanged = True
		EndIf
	EndIf

	Return $bChanged
EndFunc

Func _StatusLabel($sStatus)
	Switch StringUpper($sStatus)
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
		Case "B"
			Return "Blank"
		Case "W"
			Return "Weekend"
	EndSwitch
	Return "Unknown"
EndFunc

Func _GetOutlookItemStatus($oItem)
	If Not IsObj($oItem) Then Return ""

	Local $sStatus = _ParseStatusFromCategories($oItem.Categories)
	If _IsKnownStatus($sStatus) Then Return $sStatus

	$sStatus = _ParseStatusFromSubject($oItem.Subject)
	If _IsKnownStatus($sStatus) Then Return $sStatus

	$sStatus = _GetUserProp($oItem, "WorkDaysStatus")
	If _IsKnownStatus($sStatus) Then Return $sStatus

	Return ""
EndFunc

Func _ParseStatusFromCategories($sCategories)
	Local $sRaw = String($sCategories)
	If StringStripWS($sRaw, 3) = "" Then Return ""

	Local $sPrefix = StringLower(_CategoryPrefix())
	Local $aCategories = StringSplit($sRaw, ",")
	If Not IsArray($aCategories) Then Return ""

	For $i = 1 To $aCategories[0]
		Local $sCat = StringLower(StringStripWS($aCategories[$i], 3))
		If $sCat = "" Then ContinueLoop
		If StringLeft($sCat, StringLen($sPrefix)) = $sPrefix Then
			Local $sAfter = StringStripWS(StringTrimLeft($sCat, StringLen($sPrefix)), 3)
			Local $sStatus = _ParseStatusLabel($sAfter)
			If _IsKnownStatus($sStatus) Then Return $sStatus
		EndIf
		Local $sDirect = _ParseStatusLabel($sCat)
		If _IsKnownStatus($sDirect) Then Return $sDirect
	Next

	Return ""
EndFunc

Func _ParseStatusLabel($sText)
	Local $s = StringLower(StringStripWS(String($sText), 3))
	If $s = "" Then Return ""
	If $s = "o" Or $s = "on site" Or $s = "onsite" Then Return "O"
	If $s = "r" Or $s = "remote" Then Return "R"
	If $s = "h" Or $s = "holiday" Then Return "H"
	If $s = "p" Or $s = "pto" Or $s = "paid time off" Then Return "P"
	If $s = "t" Or $s = "travel" Then Return "T"
	If $s = "s" Or $s = "sick" Then Return "S"
	If $s = "b" Or $s = "blank" Then Return "B"
	If $s = "w" Or $s = "weekend" Then Return "W"
	Return ""
EndFunc

Func _ParseStatusFromSubject($sSubject)
	Local $s = StringLower(StringStripWS(String($sSubject), 3))
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*O\s*\]") Then Return "O"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*R\s*\]") Then Return "R"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*H\s*\]") Then Return "H"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*P\s*\]") Then Return "P"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*T\s*\]") Then Return "T"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*S\s*\]") Then Return "S"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*B\s*\]") Then Return "B"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*W\s*\]") Then Return "W"

	If StringInStr($s, "on site") Or StringInStr($s, "onsite") Then Return "O"
	If StringInStr($s, "remote") Then Return "R"
	If StringInStr($s, "holiday") Then Return "H"
	If StringInStr($s, "pto") Or StringInStr($s, "paid time off") Then Return "P"
	If StringInStr($s, "travel") Then Return "T"
	If StringInStr($s, "sick") Then Return "S"
	If StringInStr($s, "blank") Then Return "B"
	If StringInStr($s, "weekend") Then Return "W"

	; Short codes are supported when the user creates an Outlook item like "WorkDays - O".
	Local $sPrefix = StringLower(_SubjectPrefix())
	If StringLeft($s, StringLen($sPrefix)) = $sPrefix Then
		Local $sAfter = StringStripWS(StringTrimLeft($s, StringLen($sPrefix)), 3)
		Local $sCode = StringUpper(StringLeft($sAfter, 1))
		If _IsKnownStatus($sCode) Then Return $sCode
	EndIf

	Return ""
EndFunc

Func _IsKnownStatus($sStatus)
	Switch StringUpper($sStatus)
		Case "O", "R", "H", "P", "T", "S", "B", "W"
			Return True
	EndSwitch
	Return False
EndFunc

Func _CleanOutlookMarker($sBody)
	Local $s = String($sBody)
	$s = StringReplace($s, Chr(0), "")
	Return StringStripWS($s, 2)
EndFunc

Func _GetUserProp($oItem, $sName)
	If Not IsObj($oItem) Then Return ""
	Local $oProp = $oItem.UserProperties.Find($sName)
	If IsObj($oProp) Then Return String($oProp.Value)
	Return ""
EndFunc

Func _SetUserProp(ByRef $oItem, $sName, $sValue)
	If Not IsObj($oItem) Then Return 0
	Local $oProp = $oItem.UserProperties.Find($sName)
	If Not IsObj($oProp) Then $oProp = $oItem.UserProperties.Add($sName, $OL_TEXT, True)
	If IsObj($oProp) Then
		$oProp.Value = $sValue
		Return 1
	EndIf
	Return 0
EndFunc

Func _TodayISO()
	Return StringFormat("%04d-%02d-%02d", @YEAR, @MON, @MDAY)
EndFunc

Func _ISOAddDays($sISO, $iDays)
	Local $sDate = StringReplace($sISO, "-", "/") & " 00:00:00"
	Local $sNew = _DateAdd("D", $iDays, $sDate)
	Return StringReplace(StringLeft($sNew, 10), "/", "-")
EndFunc

Func _ISODiffDays($sStartISO, $sEndISO)
	Return _DateDiff("D", StringReplace($sStartISO, "-", "/") & " 00:00:00", StringReplace($sEndISO, "-", "/") & " 00:00:00")
EndFunc

Func _IsISODate($sISO)
	Return StringRegExp($sISO, "^\d{4}-\d{2}-\d{2}$") = 1
EndFunc

Func _OutlookFilterDate($sISO)
	Local $a = StringSplit($sISO, "-")
	If @error Or $a[0] <> 3 Then Return ""
	Return Number($a[2]) & "/" & Number($a[3]) & "/" & $a[1] & " 12:00 AM"
EndFunc

Func _OutlookDateToISO($vDate)
	Local $s = StringStripWS(String($vDate), 3)
	Local $aMatch = StringRegExp($s, "(\d{4})[\-/](\d{1,2})[\-/](\d{1,2})", 1)
	If IsArray($aMatch) And UBound($aMatch) = 3 Then
		Return StringFormat("%04d-%02d-%02d", Number($aMatch[0]), Number($aMatch[1]), Number($aMatch[2]))
	EndIf

	$aMatch = StringRegExp($s, "(\d{1,2})[\-/](\d{1,2})[\-/](\d{4})", 1)
	If IsArray($aMatch) And UBound($aMatch) = 3 Then
		Return StringFormat("%04d-%02d-%02d", Number($aMatch[2]), Number($aMatch[0]), Number($aMatch[1]))
	EndIf

	Return ""
EndFunc

Func _Log($sText)
	Local $sLine = StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC) & " | " & $sText & @CRLF
	FileWrite($g_sLog, $sLine)
EndFunc

Func _ComErrorHandler($oError)
	If Not IsObj($oError) Then Return
	_Log("COM error: 0x" & Hex($oError.number, 8) & " | " & $oError.windescription & " | line " & $oError.scriptline)
EndFunc
