#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Icon=CalendarSync.ico
#AutoIt3Wrapper_Res_Description=Work Day Sync Agent
#AutoIt3Wrapper_Res_Fileversion=1.0.1.0
#AutoIt3Wrapper_Res_ProductName=Work Day Sync Agent
#AutoIt3Wrapper_Res_CompanyName=Fabricio Zambroni
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2026 Fabricio Zambroni
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <TrayConstants.au3>
#include "Workdays_Backup.au3"

Opt("MustDeclareVars", 1)
Opt("TrayMenuMode", 3)
Opt("TrayOnEventMode", 0)

Global Const $g_sAppTitle = "WorkDays Outlook Agent - Version: " & FileGetVersion(@ScriptFullPath)
Global Const $g_sDB = "HKEY_CURRENT_USER\Software\WorkDays"
Global Const $g_sAgentDB = "HKEY_CURRENT_USER\Software\WorkDays\OutlookAgent"
Global Const $g_sAgentDir = @ScriptDir
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
Global $g_sLastForceSyncRequest = ""

Global $g_iPlanOutlookToWorkDaysChanges = 0
Global $g_iPlanClears = 0
Global $g_iPlanCurrentRecords = 0
Global $g_iPlanSyncableRecords = 0
Global $g_iPlanOutlookCandidates = 0
Global $g_sLastSyncPlanFile = ""
Global $g_sLastPreSyncBackup = ""
Global $g_sLastSyncGuardReason = ""

DirCreate($g_sAgentDir)
_EnsureConfig()
$g_sLastForceSyncRequest = RegRead($g_sAgentDB, "Sync_ForceNowRequest")
If @error Then $g_sLastForceSyncRequest = ""
_HandleCommandLine()

If _Singleton($g_sAppTitle, 1) = 0 Then
	MsgBox(BitOR($MB_ICONINFORMATION, $MB_TOPMOST), $g_sAppTitle, "WorkDays Outlook Agent is already running.")
	Exit
EndIf

_ApplyStartupSetting()
_CreateTray()
_Log("Agent started. Executable/script folder: " & @ScriptDir & " | Log: " & $g_sLog & " | Settings: " & $g_sAgentDB)
_VLog("Verbose mode enabled. State file: " & $g_sState)
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

	_CheckForcedSyncRequest()

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
	_EnsureRegDefault("Outlook", "DateOrder", "Auto")

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
	_EnsureRegDefault("Safety", "CreateBackupBeforeOutlookChanges", "1")
	_EnsureRegDefault("Safety", "BlockMassChanges", "1")
	_EnsureRegDefault("Safety", "MaxWorkDaysChangesPerSync", "20")
	_EnsureRegDefault("Safety", "MaxChangePercentPerSync", "15")
	_EnsureRegDefault("Safety", "MaxClearsPerSync", "0")
	_EnsureRegDefault("Safety", "BlockIncompleteOutlookRead", "1")
	_EnsureRegDefault("Safety", "IncompleteReadMinOutlookItems", "3")
	_EnsureRegDefault("Safety", "IncompleteReadMinRatioPercent", "20")

	_EnsureRegDefault("Advanced", "LogLevel", "Normal")
	_EnsureRegDefault("Logging", "VerboseMode", "0")
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


Func _CheckForcedSyncRequest()
	Local $sRequest = RegRead($g_sAgentDB, "Sync_ForceNowRequest")
	If @error Then $sRequest = ""
	If $sRequest = "" Or $sRequest = $g_sLastForceSyncRequest Then Return 0

	$g_sLastForceSyncRequest = $sRequest
	SetError(0)
	RegWrite($g_sAgentDB, "Sync_ForceNowAccepted", "REG_SZ", StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))
	_Log("Immediate sync request detected from WorkDays. Request=" & $sRequest)


	_SyncNow()
	$g_hTimer = TimerInit()
	RegWrite($g_sAgentDB, "Sync_ForceNowCompleted", "REG_SZ", StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))
	Return 1
EndFunc

Func _NotifyWorkDaysDatabaseChanged($sDateISO, $sStatus, $sMarker)
	Local $sSeq = RegRead($g_sAgentDB, "LastDatabaseChangeSeq")
	If @error Or $sSeq = "" Then $sSeq = "0"
	Local $iSeq = Number($sSeq) + 1
	Local $sNow = StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC)

	RegWrite($g_sAgentDB, "LastDatabaseChangeSeq", "REG_SZ", String($iSeq))
	RegWrite($g_sAgentDB, "LastDatabaseChangeAt", "REG_SZ", $sNow)
	RegWrite($g_sAgentDB, "LastDatabaseChangeDate", "REG_SZ", $sDateISO)
	RegWrite($g_sAgentDB, "LastDatabaseChangeStatus", "REG_SZ", $sStatus)
	RegWrite($g_sAgentDB, "LastDatabaseChangeMarkerLength", "REG_SZ", StringLen($sMarker))
	_VLog("WorkDays refresh notification updated: seq=" & $iSeq & " date=" & $sDateISO & " status=" & $sStatus & " markerLen=" & StringLen($sMarker))
	Return $iSeq
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
	_VLog("RunSync started.")
	_VLogConfigSnapshot()
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
	_VLog("Sync range: " & $sStartISO & " to " & $sEndISO & " | PastDays=" & $iPast & " | FutureDays=" & $iFuture & " | ManagedOnly=" & _Cfg("Outlook", "ManagedOnly", "0"))
	Local $oOutlookMap = _LoadOutlookMap($oCalendar, $sStartISO, $sEndISO)
	If Not IsObj($oOutlookMap) Then Return SetError(4, 0, 0)
	_VLog("Outlook candidate map loaded. Count=" & $oOutlookMap.Count)

	_BuildSyncSafetyPlan($oOutlookMap, $sStartISO, $sEndISO)
	If Not _ValidateAndPrepareSyncSafetyPlan() Then
		_Log("Sync blocked by safety guard. Reason: " & $g_sLastSyncGuardReason)
		Return SetError(20, 0, 0)
	EndIf

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

		If _VerboseEnabled() And ($bHasOutlook Or $sRegHash <> "" Or $sStateRegHash <> "" Or $sStateOutHash <> "") Then
			_VLog("Day decision " & $sDateISO & ": regStatus='" & $sRegStatus & "' regMarkerLen=" & StringLen($sRegMarker) & " regHash='" & $sRegHash & "' hasOutlook=" & _BoolText($bHasOutlook) & " outStatus='" & $sOutStatus & "' outMarkerLen=" & StringLen($sOutMarker) & " outHash='" & $sOutHash & "' stateRegHash='" & $sStateRegHash & "' stateOutHash='" & $sStateOutHash & "' stateEntryIdShort='" & _EntryIdShort($sStateEntryID) & "' regChanged=" & _BoolText($bRegChanged) & " outChanged=" & _BoolText($bOutChanged))
		EndIf

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

	_SetSyncGuardStatus("OK", "Sync applied safely. Outlook-to-WorkDays changes=" & $g_iPlanOutlookToWorkDaysChanges)
	Return $iChanges
EndFunc

Func _BuildSyncSafetyPlan($oOutlookMap, $sStartISO, $sEndISO)
	$g_iPlanOutlookToWorkDaysChanges = 0
	$g_iPlanClears = 0
	$g_iPlanCurrentRecords = 0
	$g_iPlanSyncableRecords = 0
	$g_iPlanOutlookCandidates = 0
	$g_sLastPreSyncBackup = ""
	$g_sLastSyncGuardReason = ""

	If IsObj($oOutlookMap) Then $g_iPlanOutlookCandidates = $oOutlookMap.Count

	Local $sLogDir = $g_sAgentDir & "\Logs"
	DirCreate($sLogDir)
	$g_sLastSyncPlanFile = $sLogDir & "\LastSyncPlan.txt"
	Local $hPlan = FileOpen($g_sLastSyncPlanFile, 2)

	If $hPlan <> -1 Then
		FileWriteLine($hPlan, "WORKDAYS OUTLOOK AGENT - LAST SYNC PLAN")
		FileWriteLine($hPlan, "Generated=" & StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC))
		FileWriteLine($hPlan, "Range=" & $sStartISO & " to " & $sEndISO)
		FileWriteLine($hPlan, "")
	EndIf

	Local $iDays = _ISODiffDays($sStartISO, $sEndISO)
	Local $i
	For $i = 0 To $iDays
		Local $sDateISO = _ISOAddDays($sStartISO, $i)
		Local $sRegRec = _ReadRegistryDay($sDateISO)
		Local $sRegStatus = _RecordStatus($sRegRec)
		Local $sRegMarker = _RecordMarker($sRegRec)
		Local $sRegHash = _RecordHash($sRegStatus, $sRegMarker)
		If $sRegHash <> "" Then $g_iPlanCurrentRecords += 1
		If _ShouldSync($sRegStatus, $sRegMarker) Then $g_iPlanSyncableRecords += 1

		Local $sOutRec = ""
		Local $bHasOutlook = False
		If IsObj($oOutlookMap) And $oOutlookMap.Exists($sDateISO) Then
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

		Local $sAction = ""
		Local $sTargetStatus = $sRegStatus
		Local $sTargetMarker = $sRegMarker

		If Not $bHasOutlook And $sStateOutHash <> "" And Not $bRegChanged Then
			If _Cfg("Sync", "DeleteInOutlookClearsWorkDays", "0") = "1" Then
				$sAction = "OUTLOOK_DELETE_CLEAR_WORKDAYS"
				$sTargetStatus = "B"
				$sTargetMarker = ""
			EndIf
		ElseIf $bHasOutlook And $bOutChanged And Not $bRegChanged Then
			$sAction = "PULL_OUTLOOK_CHANGE"
			$sTargetStatus = $sOutStatus
			$sTargetMarker = $sOutMarker
		ElseIf $bRegChanged And $bOutChanged Then
			If _Cfg("Sync", "OutlookWinsOnConflict", "1") = "1" And $bHasOutlook Then
				$sAction = "CONFLICT_PULL_OUTLOOK"
				$sTargetStatus = $sOutStatus
				$sTargetMarker = $sOutMarker
			EndIf
		EndIf

		If $sAction <> "" Then
			$g_iPlanOutlookToWorkDaysChanges += 1
			If $sTargetStatus = "B" And StringStripWS($sTargetMarker, 3) = "" Then $g_iPlanClears += 1
			If $hPlan <> -1 Then FileWriteLine($hPlan, $sDateISO & " | " & $sAction & " | current=" & $sRegStatus & " | outlook=" & $sOutStatus & " | target=" & $sTargetStatus & " | markerLen=" & StringLen($sTargetMarker) & " | entry=" & _EntryIdShort($sOutEntryID) & " | stateEntry=" & _EntryIdShort($sStateEntryID))
		EndIf
	Next

	If $hPlan <> -1 Then
		FileWriteLine($hPlan, "")
		FileWriteLine($hPlan, "SUMMARY")
		FileWriteLine($hPlan, "OutlookToWorkDaysChanges=" & $g_iPlanOutlookToWorkDaysChanges)
		FileWriteLine($hPlan, "Clears=" & $g_iPlanClears)
		FileWriteLine($hPlan, "CurrentWorkDaysRecordsInRange=" & $g_iPlanCurrentRecords)
		FileWriteLine($hPlan, "SyncableWorkDaysRecordsInRange=" & $g_iPlanSyncableRecords)
		FileWriteLine($hPlan, "OutlookCandidatesFound=" & $g_iPlanOutlookCandidates)
		FileClose($hPlan)
	EndIf

	RegWrite($g_sAgentDB, "LastSyncPlanFile", "REG_SZ", $g_sLastSyncPlanFile)
	_VLog("Sync safety plan built. OutlookToWorkDaysChanges=" & $g_iPlanOutlookToWorkDaysChanges & " clears=" & $g_iPlanClears & " currentRecords=" & $g_iPlanCurrentRecords & " syncableRecords=" & $g_iPlanSyncableRecords & " outlookCandidates=" & $g_iPlanOutlookCandidates & " planFile=" & $g_sLastSyncPlanFile)
	Return 1
EndFunc

Func _ValidateAndPrepareSyncSafetyPlan()
	Local $sReason = ""
	Local $iChanges = $g_iPlanOutlookToWorkDaysChanges

	If $iChanges = 0 Then
		_SetSyncGuardStatus("OK", "No Outlook-to-WorkDays database changes planned.")
		Return 1
	EndIf

	If _Cfg("Safety", "CreateBackupBeforeOutlookChanges", "1") = "1" Then
		$g_sLastPreSyncBackup = _CreateAgentPreSyncBackup("Agent_PreSync")
		If @error Or $g_sLastPreSyncBackup = "" Then
			$sReason = "Backup failed. No WorkDays database changes were applied."
			_SetSyncGuardStatus("BLOCKED", $sReason)
			Return SetError(1, 0, 0)
		EndIf
	EndIf

	If _Cfg("Safety", "BlockMassChanges", "1") = "1" Then
		Local $iMaxChanges = Number(_Cfg("Safety", "MaxWorkDaysChangesPerSync", "20"))
		If $iMaxChanges < 1 Then $iMaxChanges = 1
		If $iChanges > $iMaxChanges Then $sReason &= "Too many Outlook-to-WorkDays changes planned: " & $iChanges & " > " & $iMaxChanges & ". "

		Local $iMaxClears = Number(_Cfg("Safety", "MaxClearsPerSync", "0"))
		If $iMaxClears < 0 Then $iMaxClears = 0
		If $g_iPlanClears > $iMaxClears Then $sReason &= "Too many WorkDays clears planned: " & $g_iPlanClears & " > " & $iMaxClears & ". "

		Local $iMaxPercent = Number(_Cfg("Safety", "MaxChangePercentPerSync", "15"))
		If $iMaxPercent < 1 Then $iMaxPercent = 1
		If $g_iPlanCurrentRecords > 0 Then
			Local $nPct = ($iChanges * 100.0) / $g_iPlanCurrentRecords
			If $nPct > $iMaxPercent Then $sReason &= "Change percentage is too high: " & StringFormat("%.1f", $nPct) & "% > " & $iMaxPercent & "%. "
		EndIf
	EndIf

	If _Cfg("Safety", "BlockIncompleteOutlookRead", "1") = "1" Then
		Local $iMinItems = Number(_Cfg("Safety", "IncompleteReadMinOutlookItems", "3"))
		Local $iMinRatio = Number(_Cfg("Safety", "IncompleteReadMinRatioPercent", "20"))
		If $iMinItems < 0 Then $iMinItems = 0
		If $iMinRatio < 1 Then $iMinRatio = 1
		If $g_iPlanSyncableRecords >= 10 Then
			Local $nRatio = 0
			If $g_iPlanSyncableRecords > 0 Then $nRatio = ($g_iPlanOutlookCandidates * 100.0) / $g_iPlanSyncableRecords
			If $g_iPlanOutlookCandidates < $iMinItems Or $nRatio < $iMinRatio Then
				$sReason &= "Outlook read looks incomplete: candidates=" & $g_iPlanOutlookCandidates & ", syncable WorkDays records=" & $g_iPlanSyncableRecords & ", ratio=" & StringFormat("%.1f", $nRatio) & "%. "
			EndIf
		EndIf
	EndIf

	If StringStripWS($sReason, 3) <> "" Then
		_SetSyncGuardStatus("BLOCKED", $sReason)
		Return SetError(2, 0, 0)
	EndIf

	_SetSyncGuardStatus("OK", "Safety checks passed. Backup=" & $g_sLastPreSyncBackup)
	Return 1
EndFunc

Func _CreateAgentPreSyncBackup($sPrefix = "Agent_PreSync")
	Local $sBackupDir = $g_sAgentDir & "\Backup"
	DirCreate($sBackupDir)
	Local $sFile = $sBackupDir & "\" & $sPrefix & "_" & @YEAR & "_" & @MON & "_" & @MDAY & "_" & @HOUR & @MIN & @SEC & ".bkp"
	Local $sCreated = _WD_Backup_Create($g_sDB, $sFile, $sBackupDir, $sPrefix)
	If @error Or $sCreated = "" Then
		_Log("Pre-sync backup FAILED. Target=" & $sFile & " error=" & @error)
		Return SetError(1, 0, "")
	EndIf
	RegWrite($g_sAgentDB, "LastPreSyncBackup", "REG_SZ", $sCreated)
	_Log("Pre-sync backup created: " & $sCreated)
	Return $sCreated
EndFunc

Func _SetSyncGuardStatus($sStatus, $sReason)
	$g_sLastSyncGuardReason = $sReason
	Local $sNow = StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC)
	RegWrite($g_sAgentDB, "LastSyncGuardStatus", "REG_SZ", $sStatus)
	RegWrite($g_sAgentDB, "LastSyncGuardReason", "REG_SZ", $sReason)
	RegWrite($g_sAgentDB, "LastSyncGuardAt", "REG_SZ", $sNow)
	RegWrite($g_sAgentDB, "LastSyncGuardChanges", "REG_SZ", String($g_iPlanOutlookToWorkDaysChanges))
	RegWrite($g_sAgentDB, "LastSyncGuardClears", "REG_SZ", String($g_iPlanClears))
	RegWrite($g_sAgentDB, "LastSyncGuardOutlookCandidates", "REG_SZ", String($g_iPlanOutlookCandidates))
	RegWrite($g_sAgentDB, "LastSyncGuardWorkDaysRecords", "REG_SZ", String($g_iPlanCurrentRecords))
	RegWrite($g_sAgentDB, "LastSyncPlanFile", "REG_SZ", $g_sLastSyncPlanFile)
	If $g_sLastPreSyncBackup <> "" Then RegWrite($g_sAgentDB, "LastPreSyncBackup", "REG_SZ", $g_sLastPreSyncBackup)
	If $sStatus = "BLOCKED" Then _Log("SYNC BLOCKED: " & $sReason & " Plan=" & $g_sLastSyncPlanFile & " Backup=" & $g_sLastPreSyncBackup)
	Return 1
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
	If _IsSubjectWorkDaysCandidate($sSubject) Then Return True
	Return False
EndFunc

Func _MarkDateAsOutlookCleaned($sDateISO)
	Local $sRegRec = _ReadRegistryDay($sDateISO)
	Local $sRegStatus = _RecordStatus($sRegRec)
	Local $sRegMarker = _RecordMarker($sRegRec)
	_UpdateState($sDateISO, _RecordHash($sRegStatus, $sRegMarker), "", "")
EndFunc

Func _GetOutlookApplication()
	_VLog("Trying to connect to an existing Outlook.Application COM instance.")
	Local $oOutlook = ObjGet("", "Outlook.Application")
	If IsObj($oOutlook) Then
		_VLog("Connected to existing Outlook.Application instance.")
	Else
		_VLog("No existing Outlook.Application instance found. Trying ObjCreate.")
		$oOutlook = ObjCreate("Outlook.Application")
		If IsObj($oOutlook) Then _VLog("Created new Outlook.Application COM instance.")
	EndIf
	If Not IsObj($oOutlook) Then
		_Log("Unable to connect to Outlook.Application COM object.")
		Return SetError(1, 0, 0)
	EndIf
	Return $oOutlook
EndFunc

Func _LoadOutlookMap($oCalendar, $sStartISO, $sEndISO)
	Local $oMap = ObjCreate("Scripting.Dictionary")
	If Not IsObj($oMap) Then Return SetError(1, 0, 0)

	Local $oItems = $oCalendar.Items
	If Not IsObj($oItems) Then
		_Log("Outlook calendar Items collection is not available.")
		Return $oMap
	EndIf

	Local $sFolderName = ""
	Local $sFolderPath = ""
	Local $iTotalCount = -1
	$sFolderName = String($oCalendar.Name)
	$sFolderPath = String($oCalendar.FolderPath)
	$iTotalCount = Number($oItems.Count)
	_VLog("Default Outlook calendar folder: name='" & _DbgValue($sFolderName, 120) & "' path='" & _DbgValue($sFolderPath, 220) & "' totalItems=" & $iTotalCount)

	$oItems.IncludeRecurrences = True
	$oItems.Sort("[Start]")

	Local $sFilter = "[Start] >= '" & _OutlookFilterDate($sStartISO) & "' AND [Start] < '" & _OutlookFilterDate(_ISOAddDays($sEndISO, 1)) & "'"
	_VLog("Outlook Restrict filter: " & $sFilter)
	Local $oRange = $oItems.Restrict($sFilter)
	If Not IsObj($oRange) Then
		_Log("Outlook Restrict returned no valid range object. Filter: " & $sFilter)
		Return $oMap
	EndIf

	Local $iRangeCount = -1
	$iRangeCount = Number($oRange.Count)
	_VLog("Outlook restricted range Count=" & $iRangeCount & " | Note: if your manual item is not listed below, it is outside this default calendar/range or the date filter did not match.")
	_VLog("Manual subject formats accepted: W - On Site | W - Remote | WD - PTO | WorkDays - Travel | [WD:O] | [WD:R]")

	Local $oItem
	Local $iIndex = 0
	Local $iAccepted = 0
	Local $iRejected = 0
	Local $iDateRejected = 0
	Local $iStatusRejected = 0
	Local $iDuplicate = 0

	For $oItem In $oRange
		$iIndex += 1
		If Not IsObj($oItem) Then
			$iRejected += 1
			_VLog("Outlook item #" & $iIndex & " skipped: not an object.")
			ContinueLoop
		EndIf

		_VLog("Outlook item #" & $iIndex & " raw: " & _OutlookItemDebugSummary($oItem))

		Local $sDecision = _WorkDaysCandidateDebugReason($oItem)
		If StringLeft($sDecision, 7) <> "accept:" Then
			$iRejected += 1
			_VLog("Outlook item #" & $iIndex & " rejected as candidate: " & $sDecision)
			ContinueLoop
		EndIf

		$iAccepted += 1
		_VLog("Outlook item #" & $iIndex & " accepted as candidate: " & $sDecision)

		Local $sDateProp = _GetUserProp($oItem, "WorkDaysDate")
		Local $sDateISO = $sDateProp
		If _IsISODate($sDateISO) Then
			_VLog("Outlook item #" & $iIndex & " date source: WorkDaysDate user property = " & $sDateISO)
		Else
			$sDateISO = _OutlookDateToISOInRange($oItem.Start, $sStartISO, $sEndISO)
			_VLog("Outlook item #" & $iIndex & " date source: Start property raw='" & _DbgValue(String($oItem.Start), 120) & "' convertedISO='" & $sDateISO & "'")
		EndIf

		If Not _IsISODate($sDateISO) Then
			$iDateRejected += 1
			_VLog("Outlook item #" & $iIndex & " skipped: date could not be converted. summary=" & _OutlookItemDebugSummary($oItem))
			ContinueLoop
		EndIf

		; Visible Outlook data wins. This lets users change the day classification directly in Outlook.
		; Internal UserProperties are kept only as a fallback for old/managed items.
		Local $sCatStatus = _ParseStatusFromCategories($oItem.Categories)
		Local $sSubjectStatus = _ParseStatusFromSubject($oItem.Subject)
		Local $sPropStatus = _GetUserProp($oItem, "WorkDaysStatus")
		Local $sStatus = _GetOutlookItemStatus($oItem)
		_VLog("Outlook item #" & $iIndex & " status parse: category='" & $sCatStatus & "' subject='" & $sSubjectStatus & "' userProp='" & $sPropStatus & "' chosen='" & $sStatus & "'")
		If Not _IsKnownStatus($sStatus) Then
			$iStatusRejected += 1
			_VLog("Outlook item #" & $iIndex & " skipped: status could not be parsed after candidate acceptance.")
			ContinueLoop
		EndIf

		Local $sMarker = _CleanOutlookMarker($oItem.Body)
		Local $sEntryID = String($oItem.EntryID)
		Local $sManaged = _GetUserProp($oItem, "WorkDaysManaged")
		Local $sRec = $sEntryID & $g_sSep & $sStatus & $g_sSep & $sMarker & $g_sSep & $sManaged

		_VLog("Outlook item #" & $iIndex & " mapped: date=" & $sDateISO & " status=" & $sStatus & " (" & _StatusLabel($sStatus) & ") markerLen=" & StringLen($sMarker) & " managed='" & $sManaged & "' entryIdShort='" & _EntryIdShort($sEntryID) & "'")

		If $oMap.Exists($sDateISO) Then
			$iDuplicate += 1
			; Prefer managed items over manually created prefix-only items.
			Local $sExisting = $oMap.Item($sDateISO)
			Local $sExistingManaged = _OutlookRecordPart($sExisting, 3)
			If $sExistingManaged = "1" And $sManaged <> "1" Then
				_VLog("Duplicate candidate ignored for " & $sDateISO & ": existing item is managed and current item is not managed.")
				ContinueLoop
			EndIf
			_Log("Duplicate WorkDays Outlook item detected for " & $sDateISO & ". The agent will use the most recent candidate it found.")
			$oMap.Item($sDateISO) = $sRec
		Else
			$oMap.Add($sDateISO, $sRec)
		EndIf
	Next

	_VLog("Outlook scan summary: scanned=" & $iIndex & " accepted=" & $iAccepted & " rejected=" & $iRejected & " dateRejected=" & $iDateRejected & " statusRejected=" & $iStatusRejected & " duplicates=" & $iDuplicate & " mappedDates=" & $oMap.Count)

	Return $oMap
EndFunc


Func _IsWorkDaysCandidate($oItem)
	Return StringLeft(_WorkDaysCandidateDebugReason($oItem), 7) = "accept:"
EndFunc

Func _WorkDaysCandidateReason($oItem)
	Local $sDecision = _WorkDaysCandidateDebugReason($oItem)
	If StringLeft($sDecision, 7) = "accept:" Then Return StringStripWS(StringTrimLeft($sDecision, 7), 3)
	Return ""
EndFunc

Func _WorkDaysCandidateDebugReason($oItem)
	If Not IsObj($oItem) Then Return "reject: not an Outlook object"

	Local $sManaged = _GetUserProp($oItem, "WorkDaysManaged")
	Local $sDateProp = _GetUserProp($oItem, "WorkDaysDate")
	Local $sStatusProp = _GetUserProp($oItem, "WorkDaysStatus")
	If $sManaged = "1" Then Return "accept: WorkDaysManaged=1 user property; WorkDaysDate='" & $sDateProp & "' WorkDaysStatus='" & $sStatusProp & "'"

	If _Cfg("Outlook", "ManagedOnly", "0") = "1" Then Return "reject: Outlook_ManagedOnly=1 and item does not have WorkDaysManaged=1"

	Local $sCatStatus = _ParseStatusFromCategories($oItem.Categories)
	If _IsKnownStatus($sCatStatus) Then Return "accept: recognized Outlook category status=" & $sCatStatus & " rawCategories='" & _DbgValue(String($oItem.Categories), 160) & "'"

	Local $sSubject = String($oItem.Subject)
	If _IsSubjectWorkDaysCandidate($sSubject) Then
		Local $sSubjectStatus = _ParseStatusFromSubject($sSubject)
		Return "accept: recognized WorkDays subject pattern; parsedSubjectStatus='" & $sSubjectStatus & "' normalizedSubject='" & _DbgValue(_NormalizeOutlookTextForParsing($sSubject), 160) & "'"
	EndIf

	Return "reject: no WorkDaysManaged user property, no recognized WorkDays category, and subject did not match supported WorkDays/manual patterns. normalizedSubject='" & _DbgValue(_NormalizeOutlookTextForParsing($sSubject), 160) & "' categories='" & _DbgValue(String($oItem.Categories), 160) & "'"
EndFunc

Func _IsSubjectWorkDaysCandidate($sSubject)
	Local $s = _NormalizeOutlookTextForParsing($sSubject)
	If $s = "" Then Return False

	Local $sLower = StringLower($s)
	Local $sPrefix = _SubjectPrefix()
	If StringLeft($sLower, StringLen(StringLower($sPrefix))) = StringLower($sPrefix) Then Return True
	If StringRegExp($s, "(?i)^\s*\[\s*WD\s*[:\-]\s*[ORHPTSBW]\s*\]") Then Return True

	; Manual Outlook shorthand supported by WorkDays:
	; W - On Site, W - Remote, W: Travel, WD - PTO, WorkDay - Sick, WorkDays - Holiday, etc.
	If StringRegExp($s, "(?i)^\s*(WORKDAYS|WORKDAY|WD|W)\s*[:\-]\s*(ON\s*SITE|ONSITE|REMOTE|HOLIDAY|PAID\s+TIME\s+OFF|PTO|TRAVEL|SICK|SICK\s+DAY|BLANK|WEEKEND|O|R|H|P|T|S|B|W)\b") Then Return True

	Return False
EndFunc


Func _CreateOrUpdateOutlookItem($oOutlook, $oNs, $sDateISO, $sEntryID, $sStatus, $sMarker)
	Local $oItem = 0
	Local $bExisting = False
	If $sEntryID <> "" Then
		_VLog("CreateOrUpdate: trying existing Outlook item date=" & $sDateISO & " entryIdShort='" & _EntryIdShort($sEntryID) & "'")
		$oItem = _GetOutlookItemByEntryID($oNs, $sEntryID)
	EndIf
	If IsObj($oItem) Then $bExisting = True
	If Not IsObj($oItem) Then
		_VLog("CreateOrUpdate: creating new Outlook appointment for date=" & $sDateISO)
		$oItem = $oOutlook.CreateItem($OL_APPOINTMENT_ITEM)
	EndIf
	If Not IsObj($oItem) Then
		_Log("CreateOrUpdate FAILED: could not create Outlook item for date=" & $sDateISO)
		Return ""
	EndIf

	Local $sSubject = _BuildSubject($sStatus, $sMarker)
	Local $sCategories = _BuildCategories($sStatus, $sMarker)
	$oItem.Subject = $sSubject
	$oItem.Start = _OutlookFilterDate($sDateISO)
	$oItem.End = _OutlookFilterDate(_ISOAddDays($sDateISO, 1))
	$oItem.AllDayEvent = True
	$oItem.BusyStatus = $OL_FREE
	_SetOutlookReminder($oItem, $sMarker)
	$oItem.Body = $sMarker
	$oItem.Categories = $sCategories
	_SetUserProp($oItem, "WorkDaysManaged", "1")
	_SetUserProp($oItem, "WorkDaysDate", $sDateISO)
	_SetUserProp($oItem, "WorkDaysStatus", $sStatus)
	$oItem.Save()
	_VLog("CreateOrUpdate saved: existing=" & _BoolText($bExisting) & " date=" & $sDateISO & " status=" & $sStatus & " subject='" & _DbgValue($sSubject, 160) & "' categories='" & _DbgValue($sCategories, 160) & "' markerLen=" & StringLen($sMarker) & " entryIdShort='" & _EntryIdShort(String($oItem.EntryID)) & "'")
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
	Local $a = StringSplit($sDateISO, "-")
	If @error Or $a[0] <> 3 Then Return 0
	Local $sRegPath = $g_sDB & "\" & $a[1] & "\" & $a[2]
	Local $sRegName = $a[3]
	Local $sValue = $sStatus & $sMarker
	Local $iOK = RegWrite($sRegPath, $sRegName, "REG_SZ", $sValue)
	If $iOK Then
		Local $sReadBack = RegRead($sRegPath, $sRegName)
		_VLog("Registry write OK: date=" & $sDateISO & " path='" & $sRegPath & "' name='" & $sRegName & "' status='" & $sStatus & "' markerLen=" & StringLen($sMarker) & " valuePreview='" & _DbgValue($sValue, 120) & "' readBackPreview='" & _DbgValue($sReadBack, 120) & "'")
		_NotifyWorkDaysDatabaseChanged($sDateISO, $sStatus, $sMarker)
	Else
		_Log("Registry write FAILED: " & $sDateISO & " status=" & $sStatus & " path=" & $sRegPath & " name=" & $sRegName & " error=" & @error)
	EndIf
	Return $iOK
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
	_VLog("State updated: date=" & $sDateISO & " RegHash='" & $sRegHash & "' OutHash='" & $sOutHash & "' EntryIDShort='" & _EntryIdShort($sEntryID) & "'")
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
	Local $s = StringLower(_NormalizeOutlookTextForParsing($sText))
	If $s = "" Then Return ""
	If $s = "o" Or $s = "on site" Or $s = "onsite" Then Return "O"
	If $s = "r" Or $s = "remote" Then Return "R"
	If $s = "h" Or $s = "holiday" Then Return "H"
	If $s = "p" Or $s = "pto" Or $s = "paid time off" Then Return "P"
	If $s = "t" Or $s = "travel" Then Return "T"
	If $s = "s" Or $s = "sick" Or $s = "sick day" Then Return "S"
	If $s = "b" Or $s = "blank" Then Return "B"
	If $s = "w" Or $s = "weekend" Then Return "W"
	Return ""
EndFunc

Func _ParseStatusFromSubject($sSubject)
	Local $s = StringLower(_NormalizeOutlookTextForParsing($sSubject))

	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*O\s*\]") Then Return "O"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*R\s*\]") Then Return "R"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*H\s*\]") Then Return "H"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*P\s*\]") Then Return "P"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*T\s*\]") Then Return "T"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*S\s*\]") Then Return "S"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*B\s*\]") Then Return "B"
	If StringRegExp($s, "(?i)\[\s*WD\s*[:\-]\s*W\s*\]") Then Return "W"

	; Manual Outlook shorthand. These are only parsed when the subject clearly starts as a WorkDays item.
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(o|on\s*site|onsite)\b") Then Return "O"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(r|remote)\b") Then Return "R"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(h|holiday)\b") Then Return "H"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(p|pto|paid\s+time\s+off)\b") Then Return "P"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(t|travel)\b") Then Return "T"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(s|sick|sick\s+day)\b") Then Return "S"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(b|blank)\b") Then Return "B"
	If StringRegExp($s, "(?i)^\s*(workdays|workday|wd|w)\s*[:\-]\s*(w|weekend)\b") Then Return "W"

	; Full WorkDays prefix and short code, for example "WorkDays - O".
	Local $sPrefix = StringLower(_SubjectPrefix())
	If StringLeft($s, StringLen($sPrefix)) = $sPrefix Then
		Local $sAfter = StringStripWS(StringTrimLeft($s, StringLen($sPrefix)), 3)
		Local $sStatus = _ParseStatusLabel($sAfter)
		If _IsKnownStatus($sStatus) Then Return $sStatus
		Local $sCode = StringUpper(StringLeft($sAfter, 1))
		If _IsKnownStatus($sCode) Then Return $sCode
	EndIf

	; Fallback for old managed items whose subject may have extra text.
	If StringInStr($s, "on site") Or StringInStr($s, "onsite") Then Return "O"
	If StringInStr($s, "remote") Then Return "R"
	If StringInStr($s, "holiday") Then Return "H"
	If StringInStr($s, "pto") Or StringInStr($s, "paid time off") Then Return "P"
	If StringInStr($s, "travel") Then Return "T"
	If StringInStr($s, "sick") Then Return "S"
	If StringInStr($s, "blank") Then Return "B"
	If StringInStr($s, "weekend") Then Return "W"

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

Func _OutlookCompactDateToISO($sRaw)
	Local $s = StringStripWS(String($sRaw), 3)

	; Outlook COM can return compact date/time strings for all-day appointments,
	; especially for manually-created items. Example: 20260702000000.
	; That means YYYYMMDDHHMMSS and must be mapped to 2026-07-02.
	If StringRegExp($s, "^\d{14}$") Then
		Return _BuildISOIfValid(Number(StringLeft($s, 4)), Number(StringMid($s, 5, 2)), Number(StringMid($s, 7, 2)))
	EndIf

	; Also support compact date-only values in case Outlook returns YYYYMMDD.
	If StringRegExp($s, "^\d{8}$") Then
		Return _BuildISOIfValid(Number(StringLeft($s, 4)), Number(StringMid($s, 5, 2)), Number(StringMid($s, 7, 2)))
	EndIf

	Return ""
EndFunc

Func _OutlookDateToISO($vDate)
	Local $s = StringStripWS(String($vDate), 3)
	Local $sCompactISO = _OutlookCompactDateToISO($s)
	If _IsISODate($sCompactISO) Then Return $sCompactISO
	Local $aMatch = StringRegExp($s, "(\d{4})[\-/](\d{1,2})[\-/](\d{1,2})", 1)
	If IsArray($aMatch) And UBound($aMatch) = 3 Then
		Return _BuildISOIfValid(Number($aMatch[0]), Number($aMatch[1]), Number($aMatch[2]))
	EndIf

	$aMatch = StringRegExp($s, "(\d{1,2})[\-/](\d{1,2})[\-/](\d{4})", 1)
	If IsArray($aMatch) And UBound($aMatch) = 3 Then
		; Backward-compatible fallback: month/day/year.
		Return _BuildISOIfValid(Number($aMatch[2]), Number($aMatch[0]), Number($aMatch[1]))
	EndIf

	Return ""
EndFunc

Func _OutlookDateToISOInRange($vDate, $sStartISO, $sEndISO)
	Local $s = StringStripWS(String($vDate), 3)
	Local $sCompactISO = _OutlookCompactDateToISO($s)
	If _IsISODate($sCompactISO) Then
		_VLog("Date parser compact Outlook raw='" & _DbgValue($s, 120) & "' -> " & $sCompactISO & " inRange=" & _BoolText(_ISOInRange($sCompactISO, $sStartISO, $sEndISO)))
		Return $sCompactISO
	EndIf

	Local $aMatch = StringRegExp($s, "(\d{4})[\-/](\d{1,2})[\-/](\d{1,2})", 1)
	If IsArray($aMatch) And UBound($aMatch) = 3 Then
		Local $sYMD = _BuildISOIfValid(Number($aMatch[0]), Number($aMatch[1]), Number($aMatch[2]))
		_VLog("Date parser YMD raw='" & _DbgValue($s, 120) & "' -> " & $sYMD)
		Return $sYMD
	EndIf

	$aMatch = StringRegExp($s, "(\d{1,2})[\-/](\d{1,2})[\-/](\d{4})", 1)
	If IsArray($aMatch) And UBound($aMatch) = 3 Then
		Local $iA = Number($aMatch[0])
		Local $iB = Number($aMatch[1])
		Local $iY = Number($aMatch[2])
		Local $sMDY = _BuildISOIfValid($iY, $iA, $iB)
		Local $sDMY = _BuildISOIfValid($iY, $iB, $iA)
		Local $bMDY = _ISOInRange($sMDY, $sStartISO, $sEndISO)
		Local $bDMY = _ISOInRange($sDMY, $sStartISO, $sEndISO)
		Local $sDateOrder = StringUpper(_Cfg("Outlook", "DateOrder", "Auto"))

		_VLog("Date parser ambiguous raw='" & _DbgValue($s, 120) & "' candidateMDY='" & $sMDY & "' inRange=" & _BoolText($bMDY) & " candidateDMY='" & $sDMY & "' inRange=" & _BoolText($bDMY) & " DateOrder=" & $sDateOrder & " syncRange=" & $sStartISO & ".." & $sEndISO)

		If $sDateOrder = "DMY" And $sDMY <> "" Then Return $sDMY
		If $sDateOrder = "MDY" And $sMDY <> "" Then Return $sMDY

		If $bMDY And Not $bDMY Then Return $sMDY
		If $bDMY And Not $bMDY Then Return $sDMY
		If $bMDY And $bDMY Then
			_VLog("Date parser warning: both MDY and DMY candidates are inside the sync range. Using MDY for backward compatibility. Set Outlook_DateOrder=DMY in the registry if needed.")
			Return $sMDY
		EndIf

		If $sMDY <> "" Then Return $sMDY
		If $sDMY <> "" Then Return $sDMY
	EndIf

	_VLog("Date parser failed raw='" & _DbgValue($s, 120) & "'")
	Return ""
EndFunc

Func _BuildISOIfValid($iYear, $iMonth, $iDay)
	If $iYear < 1900 Or $iYear > 2200 Then Return ""
	If $iMonth < 1 Or $iMonth > 12 Then Return ""
	If $iDay < 1 Or $iDay > 31 Then Return ""
	Local $sCheck = StringFormat("%04d/%02d/%02d", $iYear, $iMonth, $iDay)
	If Not _DateIsValid($sCheck) Then Return ""
	Return StringFormat("%04d-%02d-%02d", $iYear, $iMonth, $iDay)
EndFunc

Func _ISOInRange($sISO, $sStartISO, $sEndISO)
	If Not _IsISODate($sISO) Then Return False
	If _ISODiffDays($sStartISO, $sISO) < 0 Then Return False
	If _ISODiffDays($sISO, $sEndISO) < 0 Then Return False
	Return True
EndFunc



Func _VLogConfigSnapshot()
	If Not _VerboseEnabled() Then Return
	_VLog("Runtime: compiled=" & _BoolText(@Compiled) & " scriptFullPath='" & @ScriptFullPath & "' scriptDir='" & @ScriptDir & "' autoItExe='" & @AutoItExe & "' os=" & @OSVersion & " arch=" & @OSArch & " osLang=" & @OSLang)
	_VLog("Paths: registryRoot='" & $g_sDB & "' agentRegistry='" & $g_sAgentDB & "' state='" & $g_sState & "' log='" & $g_sLog & "'")
	_VLog("Settings snapshot: IntervalMinutes=" & _Cfg("Sync", "IntervalMinutes", "15") & " PastDays=" & _Cfg("Sync", "PastDays", "60") & " FutureDays=" & _Cfg("Sync", "FutureDays", "370") & " OutlookWinsOnConflict=" & _Cfg("Sync", "OutlookWinsOnConflict", "1") & " DeleteInOutlookClearsWorkDays=" & _Cfg("Sync", "DeleteInOutlookClearsWorkDays", "0"))
	_VLog("Settings snapshot: SyncBlank=" & _Cfg("Sync", "SyncBlank", "0") & " SyncWeekend=" & _Cfg("Sync", "SyncWeekend", "0") & " SyncTaggedBlankOrWeekend=" & _Cfg("Sync", "SyncTaggedBlankOrWeekend", "1") & " RunAtWindowsStartup=" & _Cfg("Sync", "RunAtWindowsStartup", "0"))
	_VLog("Settings snapshot: SubjectPrefix='" & _Cfg("Outlook", "SubjectPrefix", "WorkDays -") & "' CategoryPrefix='" & _Cfg("Outlook", "CategoryPrefix", "WorkDays -") & "' ManagedOnly=" & _Cfg("Outlook", "ManagedOnly", "0") & " DateOrder=" & _Cfg("Outlook", "DateOrder", "Auto") & " VerboseMode=" & _Cfg("Logging", "VerboseMode", "0") & " LogLevel=" & _Cfg("Advanced", "LogLevel", "Normal"))
	_VLog("Safety snapshot: CreateBackupBeforeOutlookChanges=" & _Cfg("Safety", "CreateBackupBeforeOutlookChanges", "1") & " BlockMassChanges=" & _Cfg("Safety", "BlockMassChanges", "1") & " MaxWorkDaysChangesPerSync=" & _Cfg("Safety", "MaxWorkDaysChangesPerSync", "20") & " MaxChangePercentPerSync=" & _Cfg("Safety", "MaxChangePercentPerSync", "15") & " MaxClearsPerSync=" & _Cfg("Safety", "MaxClearsPerSync", "0") & " BlockIncompleteOutlookRead=" & _Cfg("Safety", "BlockIncompleteOutlookRead", "1"))
EndFunc

Func _OutlookItemDebugSummary($oItem)
	If Not IsObj($oItem) Then Return "not an object"
	Local $sSubject = String($oItem.Subject)
	Local $sCategories = String($oItem.Categories)
	Local $sStart = String($oItem.Start)
	Local $sEnd = String($oItem.End)
	Local $sAllDay = String($oItem.AllDayEvent)
	Local $sBusy = String($oItem.BusyStatus)
	Local $sClass = String($oItem.Class)
	Local $sMsgClass = String($oItem.MessageClass)
	Local $sEntryID = String($oItem.EntryID)
	Local $sManaged = _GetUserProp($oItem, "WorkDaysManaged")
	Local $sWDDate = _GetUserProp($oItem, "WorkDaysDate")
	Local $sWDStatus = _GetUserProp($oItem, "WorkDaysStatus")
	Return "class='" & _DbgValue($sClass, 20) & "' messageClass='" & _DbgValue($sMsgClass, 50) & "' subject='" & _DbgValue($sSubject, 180) & "' normalizedSubject='" & _DbgValue(_NormalizeOutlookTextForParsing($sSubject), 180) & "' start='" & _DbgValue($sStart, 80) & "' end='" & _DbgValue($sEnd, 80) & "' allDay=" & $sAllDay & " busy=" & $sBusy & " categories='" & _DbgValue($sCategories, 180) & "' WorkDaysManaged='" & $sManaged & "' WorkDaysDate='" & $sWDDate & "' WorkDaysStatus='" & $sWDStatus & "' entryIdShort='" & _EntryIdShort($sEntryID) & "'"
EndFunc

Func _NormalizeOutlookTextForParsing($sText)
	Local $s = String($sText)
	$s = StringReplace($s, ChrW(160), " ") ; non-breaking space
	$s = StringReplace($s, ChrW(8209), "-") ; non-breaking hyphen
	$s = StringReplace($s, ChrW(8210), "-") ; figure dash
	$s = StringReplace($s, ChrW(8211), "-") ; en dash
	$s = StringReplace($s, ChrW(8212), "-") ; em dash
	$s = StringReplace($s, ChrW(8722), "-") ; minus sign
	$s = StringReplace($s, @TAB, " ")
	While StringInStr($s, "  ")
		$s = StringReplace($s, "  ", " ")
	WEnd
	Return StringStripWS($s, 3)
EndFunc

Func _DbgValue($sValue, $iMax = 220)
	Local $s = String($sValue)
	$s = StringReplace($s, @CR, "\\r")
	$s = StringReplace($s, @LF, "\\n")
	$s = StringReplace($s, @TAB, "\\t")
	If StringLen($s) > $iMax Then $s = StringLeft($s, $iMax) & "..."
	Return $s
EndFunc

Func _EntryIdShort($sEntryID)
	Local $s = String($sEntryID)
	If StringLen($s) <= 18 Then Return $s
	Return StringLeft($s, 8) & "..." & StringRight($s, 8)
EndFunc

Func _BoolText($bValue)
	If $bValue Then Return "true"
	Return "false"
EndFunc

Func _VerboseEnabled()
	If _Cfg("Logging", "VerboseMode", "0") = "1" Then Return True
	If StringLower(_Cfg("Advanced", "LogLevel", "Normal")) = "verbose" Then Return True
	Return False
EndFunc

Func _VLog($sText)
	If _VerboseEnabled() Then _Log("[VERBOSE] " & $sText)
EndFunc

Func _Log($sText)
	Local $sLine = StringFormat("%04d-%02d-%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC) & " | " & $sText & @CRLF
	FileWrite($g_sLog, $sLine)
EndFunc

Func _ComErrorHandler($oError)
	If Not IsObj($oError) Then Return
	_Log("COM error: 0x" & Hex($oError.number, 8) & " | " & $oError.windescription & " | line " & $oError.scriptline)
EndFunc
