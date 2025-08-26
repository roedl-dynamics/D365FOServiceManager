#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=zahnrad.ico
#AutoIt3Wrapper_Res_Description=Program to stop and start services via a GUI
#AutoIt3Wrapper_Res_Fileversion=1.0.0.9
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_CompanyName=Visionet Dynamics GmbH
#AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.0
 Author:         Carsten Weda

 Script Function:
	Manage Microsoft services via a GUI.

#ce ----------------------------------------------------------------------------

Opt("MustDeclareVars", 1)
Opt("TrayAutoPause", 0)

#include <Array.au3>
#include <AutoItConstants.au3>
#include <ButtonConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GUIStatusBar.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <Process.au3>
#include <ProgressConstants.au3>
#include <SecurityConstants.au3>
#include <Services.au3>
#include <WindowsConstants.au3>


#Region ### START Koda GUI section ### Form=
Global $hGui = GUICreate("Visionet Dynamics - D365 Service Maintain", 536, 400, -1, -1)

Global $idProgress = GUICtrlCreateProgress(28, 22, 481, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idStopButton = GUICtrlCreateButton("Stop all", 28, 110, 225, 73, $WS_BORDER)
Global $idStartButton = GUICtrlCreateButton("Start all", 283, 110, 225, 73,  $WS_BORDER)
;~ Global $idAdd = GUICtrlCreateButton("Add", 135, 352, 100, 20)
Global $idAdd = GUICtrlCreateButton("Add", 170, 352, 64, 20)
Global $idRefresh = GUICtrlCreateButton("", 250, 352, 20, 20, $BS_ICON)
GUICtrlSetImage(-1, "C:\Windows\System32\imageres.dll", -229, 0)
Global $idInput = GUICtrlCreateInput("", 283, 352, 225, 21)
;~ Global $idRemove = GUICtrlCreateButton("Remove", 28, 352, 100, 20)
Global $idRemove = GUICtrlCreateButton("Remove", 28, 352, 64, 20)
Global $idListview = GUICtrlCreateListView("Service | Name| Status | Starting-, StoppingID", 28, 194, 481, 141, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT))
Global $idReset = GUICtrlCreateButton("Reset", 99, 352, 64, 20)

_GUICtrlListView_SetColumnWidth($idListview, 0, 183)
_GUICtrlListView_SetColumnWidth($idListview, 1, 183)
_GUICtrlListView_SetColumnWidth($idListview, 2, 110)

Global $idContextmenu = GUICtrlCreateContextMenu($idListview)
Global $idSubmenuStart = GUICtrlCreateMenuItem("Start", $idContextmenu)
Global $idSubmenuStop = GUICtrlCreateMenuItem("Stop", $idContextmenu)
Global $idSubmenuRemove = GUICtrlCreateMenuItem("Remove", $idContextmenu)

Global $oShell = ObjCreate("shell.application"), $Enter_KEY = GUICtrlCreateDummy(), $sStatus, $sServiceAction

Dim $Arr[1][2] = [["{ENTER}", $Enter_KEY]]
GUISetAccelerators($Arr)

_setServices()

GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
#EndRegion ### END Koda GUI section ###

;Local $class = "[CLASS:D365FOServiceManager.exe]"
Local $state
Local $counter = 0


; Sets the status and name of the services in the predefined list on startup. UPD: included in _setServices()

;_refreshStatus()

; Actions to execute on certain interactions with the GUI

While 1
	$state = WinGetState($hGui)
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			ExitLoop
		Case $idStartButton
			_GUICtrlListView_SetItemSelected($idListview, -1, False)
			$sServiceAction = "NET START "
			$sStatus = "Running"
			_manageServices()
		Case $idStopButton
			_GUICtrlListView_SetItemSelected($idListview, -1, False)
			$sServiceAction = "NET STOP "
			$sStatus = "Stopped"
			_manageServices()
		Case $idAdd
			_addService()
		Case $Enter_KEY
			If ControlGetHandle($hGui, "", $idInput) = ControlGetHandle($hGui, "", ControlGetFocus($hGui))  Then
				_addService()
			EndIf
		Case $idRemove
			_removeService()
		Case $idReset
			_setServices()
		Case $idSubmenuStart
			$sServiceAction = "NET START "
			$sStatus = "Running"
			_manageServices()
		Case $idSubmenuStop
			$sServiceAction = "NET STOP "
			$sStatus = "Stopped"
			_manageServices()
		Case $idSubmenuRemove
			_removeService()
		Case $idRefresh
			_refreshStatus()
	EndSwitch

	;Prüft ob das Fenster wieder Aktiv ist und führt einen Refresh durch
	If BitAND($state,$WIN_STATE_ACTIVE) and $counter == 0 then
		_refreshStatus()
		$counter = $counter+1
		;ConsoleWrite("Refresh wurde durchgeführt"&@CRLF)
	EndIf
	If not BitAND($state,$WIN_STATE_ACTIVE) and $counter == 1 then
		$counter = $counter-1
		;ConsoleWrite("Refresh wurde nicht durchgeführt"&@CRLF)
	EndIf

WEnd
Exit 0


Func WM_NOTIFY($hWnd, $MsgID, $wParam, $lParam)
	#forceref $hWnd, $MsgID, $wParam
	Local $tagNMHDR = DllStructCreate("int;int;int", $lParam)
	If @error Then Return $GUI_RUNDEFMSG
	Local $code = DllStructGetData($tagNMHDR, 3)
	If $code = $NM_RCLICK Then

		; Disable start/stop in context menu if more than 1 service is selected
		If _GUICtrlListView_GetSelectedCount($idListview) > 1 Then
			GUICtrlSetState($idSubmenuStart, $GUI_DISABLE)
			GUICtrlSetState($idSubmenuStop, $GUI_DISABLE)
		Else
			GUICtrlSetState($idSubmenuStart, $GUI_ENABLE)
			GUICtrlSetState($idSubmenuStop, $GUI_ENABLE)
		EndIf

		; If 1 service is selected
		If _GUICtrlListView_GetSelectedCount($idListview) = 1 Then
			If _GUICtrlListView_GetItemText($idListview, Int(_GUICtrlListView_GetSelectedIndices($idListview)), 2) = "Running" Then
				GUICtrlSetState($idSubmenuStart, $GUI_DISABLE)
			ElseIf _GUICtrlListView_GetItemText($idListview, Int(_GUICtrlListView_GetSelectedIndices($idListview)), 2) = "Stopped" Then
				GUICtrlSetState($idSubmenuStop, $GUI_DISABLE)
			Else
				GUICtrlSetState($idSubmenuStart, $GUI_DISABLE)
				GUICtrlSetState($idSubmenuStop, $GUI_DISABLE)
			EndIf
		EndIf
	EndIf
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY


; Start/Stop services in the list

Func _manageServices()
	Local $aPID[0]
	Local $iListlength = _GUICtrlListView_GetItemCount($idListview) - 1
	GUICtrlSetData($idProgress, 0)
	WinSetOnTop($hGui, "", $WINDOWS_ONTOP)

	; Manage only 1 selected service with context menu
	If _GUICtrlListView_GetSelectedCount($idListview) = 1 Then
		Local $iServiceIndex = Int(_GUICtrlListView_GetSelectedIndices($idListview))

		If _Service_Exists(_GUICtrlListView_GetItemText($idListview, $iServiceIndex)) Then
			Local $iPID = Run($sServiceAction & _GUICtrlListView_GetItemText($idListview, $iServiceIndex))
			_GUICtrlListView_SetItem($idListview, $iPID, $iServiceIndex, 3)
			_ArrayAdd($aPID, $iPID)
		EndIf

	; Manage all services in the list
	Else
		For $i = 0 To $iListlength
			Local $sService = _GUICtrlListView_GetItemText($idListview, $i)
			If _Service_Exists($sService) Then
				Local $iPID = Run($sServiceAction & $sService)
				_GUICtrlListView_SetItem($idListview, $iPID, $i, 3)
				_ArrayAdd($aPID, $iPID)
			EndIf
		Next
	EndIf

	; Continiously check if the starting processes are running and update the progressbar
	While True
		Local $ProcessActive = False
		Local $iRunningProcesses = 0

		For $i = 0 To UBound($aPID) - 1

			If ProcessExists($aPID[$i]) Then
					$ProcessActive = True
			ElseIf IsDeclared("iServiceIndex") Then
				_GUICtrlListView_SetItem($idListview, $sStatus, $iServiceIndex, 2)
				_GUICtrlListView_SetItem($idListview, "", $iServiceIndex, 3)
				$iRunningProcesses += 1
			Else
				_GUICtrlListView_SetItem($idListview, $sStatus, _GUICtrlListView_FindInText($idListview, $aPID[$i]), 2)
				_GUICtrlListView_SetItem($idListview, "", _GUICtrlListView_FindInText($idListview, $aPID[$i]), 3)
				$iRunningProcesses += 1
			EndIf
		Next

		GUICtrlSetData($idProgress, 100 / UBound($aPID) * $iRunningProcesses)

		If Not $ProcessActive Then
			Sleep(1000)
			GUICtrlSetData($idProgress, 0)
			ExitLoop
		Else
			Sleep(50)
		EndIf
	WEnd
	WinSetOnTop($hGui, "", $WINDOWS_NOONTOP)
EndFunc   ;==>_manageServices

; Adds services written in the input into the list

Func _addService()
	Local $sServiceAdd = GUICtrlRead($idInput)

	; Check if service exists
	If _Service_Exists($sServiceAdd) = 0 Then
		MsgBox(0, "Error", $sServiceAdd & " does NOT EXIST")
		Return
	EndIf

	; Check if service includes illegal characters
	If StringRegExp(GUICtrlRead($idInput), "[^0-9a-zA-Z._]") Then
		MsgBox($MB_SYSTEMMODAL, "", "Illegal use of characters.")
	Else

		; Check if service is already listed
		If $sServiceAdd = _GUICtrlListView_GetItemText($idListview, _GUICtrlListView_FindText($idListview, $sServiceAdd)) Then
			MsgBox($MB_SYSTEMMODAL, "", "This service is already listed")
		Else

			; Set the status of the service
			If $oShell.IsServiceRunning($sServiceAdd) Then
				$sStatus = "Running"
			Else
				$sStatus = "Stopped"
			EndIf

			; Add the service into the list
			GUICtrlCreateListViewItem($sServiceAdd & "|" & _Service_QueryDisplayName($sServiceAdd) & "|" & $sStatus, $idListview)
		EndIf
	EndIf
EndFunc   ;==>_addService

; Remove services (either selected ones from list, the one written in the input or both)

Func _removeService()
	Local $sServiceRemove = GUICtrlRead($idInput)

	If _GUICtrlListView_GetSelectedCount($idListview) >= 1 Then
		_GUICtrlListView_DeleteItemsSelected($idListview)
	EndIf
	If $sServiceRemove = _GUICtrlListView_GetItemText($idListview, _GUICtrlListView_FindText($idListview, $sServiceRemove)) Then
		_GUICtrlListView_DeleteItem($idListview, _GUICtrlListView_FindText($idListview, $sServiceRemove))
	EndIf
EndFunc   ;==>_removeService

; Refresh the status of the services

Func _refreshStatus()
	For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
		Local $sService = _GUICtrlListView_GetItemText($idListview, $i)
		Local $serviceExist = _Service_Exists(_GUICtrlListView_GetItemText($idListview, $i))

		If $serviceExist = 0 Then
			$sStatus = "Unknown"
		EndIf
		If $oShell.IsServiceRunning($sService) Then
			$sStatus = "Running"
		ElseIf $serviceExist = 1 And Not $oShell.IsServiceRunning($sService) Then
			$sStatus = "Stopped"
		EndIf
		_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName($sService), $i, 1)
		_GUICtrlListView_SetItem($idListview, $sStatus, $i, 2)
	Next
	ConsoleWrite("Refresh wurde durchgeführt"&@CRLF)
EndFunc   ;==>_refreshStatus

Func _setServices()
	_GUICtrlListView_DeleteAllItems($idListview)
	GUICtrlCreateListViewItem("MR2012ProcessService|" & $sStatus, $idListview)
	GUICtrlCreateListViewItem("DynamicsAxBatch|" & $sStatus, $idListview)
	GUICtrlCreateListViewItem("Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe|" & $sStatus, $idListview)
	GUICtrlCreateListViewItem("W3SVC|" & $sStatus, $idListview)
	_refreshStatus()
EndFunc
