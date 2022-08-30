#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=zahnrad_black.ico
#AutoIt3Wrapper_Res_Description=Program to stop and start services via a GUI
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_CompanyName=Rödl Dynamics GmbH
#AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Opt("MustDeclareVars", 1)

#include <Array.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GUIStatusBar.au3>
#include <MsgBoxConstants.au3>
#include <Process.au3>
#include <ProgressConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <SecurityConstants.au3>
#include "Services.au3"

#Region ### START Koda GUI section ### Form=
Global $hGui = GUICreate("Rödl Dynamics - D365 Service Maintain", 536, 411, -1, -1)

Global $idProgress = GUICtrlCreateProgress(28, 22, 481, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idStopButton = GUICtrlCreateButton("Stop", 28, 110, 225, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idStartButton = GUICtrlCreateButton("Start", 283, 110, 225, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idAdd = GUICtrlCreateButton("Add", 350, 376, 100, 20)
Global $idAddService = GUICtrlCreateInput("", 283, 349, 225, 21)
Global $idRemove = GUICtrlCreateButton("Remove", 102, 349, 100, 20)
Global $idListview = GUICtrlCreateListView("Service | Name| Status", 28, 194, 481, 141, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT))

_GUICtrlListView_SetColumnWidth($idListview, 0, 183)
_GUICtrlListView_SetColumnWidth($idListview, 1, 183)
_GUICtrlListView_SetColumnWidth($idListview, 2, 110)

GUISetState(@SW_SHOW)

Global $sStatus
GUICtrlCreateListViewItem("MR2012ProcessService|" & $sStatus, $idListview)
GUICtrlCreateListViewItem("DynamicsAxBatch|" & $sStatus, $idListview)
GUICtrlCreateListViewItem("Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe|" & $sStatus, $idListview)
GUICtrlCreateListViewItem("W3SVC|" & $sStatus, $idListview)

Global $iRunningProcesses = 0

Global $aiExist, $aiStatus
#EndRegion ### END Koda GUI section ###



For $i = 0 To _GUICtrlListView_GetItemCount($idListview)
	Local $obShell = ObjCreate("shell.application")
	If $obShell.IsServiceRunning(_GUICtrlListView_GetItemText($idListview, $i)) Then
		$sStatus = "Running"
		$iRunningProcesses += 1
		_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName(_GUICtrlListView_GetItemText($idListview, $i)), $i, 1)
		_GUICtrlListView_SetItem($idListview, $sStatus, $i, 2)
	Else
		$sStatus = "Stopped"
		_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName(_GUICtrlListView_GetItemText($idListview, $i)), $i, 1)
		_GUICtrlListView_SetItem($idListview, $sStatus, $i, 2)
	EndIf
Next


While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			ExitLoop
		Case $idStopButton
			_Stop()
		Case $idStartButton
			_Start()
		Case $idAdd
			_addService()
		Case $idRemove
			_removeService()
	EndSwitch
WEnd
Exit 0

;start the services
Func _Start()
	Local $oShell = ObjCreate("shell.application")

	For $iNetStart = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
		Local $sService = _GUICtrlListView_GetItemText($idListview, $iNetStart)
		$aiExist = _Service_Exists($sService)
		$aiStatus = _Service_QueryStatus($sService)

		If $aiExist = 0 Then
			MsgBox(0, "Error", _GUICtrlListView_GetItemText($idListview, $iNetStart) & " does NOT EXIST")
			_GUICtrlListView_DeleteItem($idListview, $iNetStart)
			ExitLoop
		ElseIf $oShell.IsServiceRunning($sService) Then
			$sStatus = "Running"

			_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName(_GUICtrlListView_GetItemText($idListview, $iNetStart)), $iNetStart, 1)
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStart, 2)
		Else
			Run("NET START " & $sService)
			WinActivate($hGui)
			Do
				Sleep(100)
			Until $oShell.IsServiceRunning($sService)
			$sStatus = "Running"

			$iRunningProcesses += 1
			GUICtrlSetData($idProgress, 100 / _GUICtrlListView_GetItemCount($idListview) * $iRunningProcesses)
			_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName(_GUICtrlListView_GetItemText($idListview, $iNetStart)), $iNetStart, 1)
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStart, 2)
			Sleep(500)
		EndIf
	Next
EndFunc   ;==>_Start


;stop the services
Func _Stop()
	Local $oShell = ObjCreate("shell.application")
	Local $iStoppedProcesses

	For $iNetStop = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
		Local $sServiceStop = _GUICtrlListView_GetItemText($idListview, $iNetStop)
		If Not $oShell.IsServiceRunning($sServiceStop) Then
			$sStatus = "Stopped"
			$iStoppedProcesses += 1
			_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName(_GUICtrlListView_GetItemText($idListview, $iNetStop)), $iNetStop, 1)
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStop, 2)
		Else
			Run("NET STOP " & $sServiceStop)
			WinActivate($hGui)
			Do
				Sleep(100)
			Until Not $oShell.IsServiceRunning($sServiceStop)
			$sStatus = "Stopped"
			$iRunningProcesses -= 1
			$iStoppedProcesses += 1
;~ 			GUICtrlSetData($idProgress, 100 / _GUICtrlListView_GetItemCount($idListview) * $iStoppedProcesses)
			_GUICtrlListView_SetItem($idListview, _Service_QueryDisplayName(_GUICtrlListView_GetItemText($idListview, $iNetStop)), $iNetStop, 1)
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStop, 2)
			Sleep(500)
		EndIf
		GUICtrlSetData($idProgress, 100 / _GUICtrlListView_GetItemCount($idListview) * $iStoppedProcesses)
	Next
	$iStoppedProcesses = 0
EndFunc   ;==>_Stop

;add service
Func _addService()
	Local $sServiceAdd = GUICtrlRead($idAddService)
	Local $oShell = ObjCreate("shell.application")

	If StringRegExp(GUICtrlRead($idAddService), "[^0-9,a-z,A-Z,.,_]") Then
		MsgBox($MB_SYSTEMMODAL, "", "Illegal use of characters.")
	Else
		If $sServiceAdd = _GUICtrlListView_GetItemText($idListview, _GUICtrlListView_FindText($idListview, $sServiceAdd)) Then
			MsgBox($MB_SYSTEMMODAL, "", "This service is already listed")
		Else
			If $oShell.IsServiceRunning($sServiceAdd) Then
				$sStatus = "Running"
				$iRunningProcesses += 1
			Else
				$sStatus = "Stopped"
			EndIf
			GUICtrlCreateListViewItem($sServiceAdd & "|" & _Service_QueryDisplayName($sServiceAdd) & "|" & $sStatus, $idListview)
		EndIf
	EndIf
EndFunc   ;==>_addService

Func _removeService()
	_GUICtrlListView_DeleteItemsSelected($idListview)
EndFunc   ;==>_removeService
