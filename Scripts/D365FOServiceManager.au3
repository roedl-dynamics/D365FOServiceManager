#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=zahnrad_black.ico
#AutoIt3Wrapper_Res_Description=Program to stop and start services via a GUI
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
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
Global $hGui = GUICreate("Rödl Dynamics - D365 Service Maintain", 535, 440, -1, -1)

Global $idProgress = GUICtrlCreateProgress(28, 22, 481, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idStopButton = GUICtrlCreateButton("Stop", 27, 110, 225, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idStartButton = GUICtrlCreateButton("Start", 283, 110, 225, 73, BitOR($PBS_SMOOTH, $WS_BORDER))
Global $idAdd = GUICtrlCreateButton("Add", 350, 280, 100, 20)
Global $idAddService = GUICtrlCreateInput("", 283, 245, 225, 20)
Global $idRemove = GUICtrlCreateButton("Remove", 350, 340, 100, 20)
Global $idListview = GUICtrlCreateListView("Service |Status", 27, 210, 225, 197, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT))

_GUICtrlListView_SetColumnWidth($idListview, 0, 110)
_GUICtrlListView_SetColumnWidth($idListview, 1, 110)
GUISetState(@SW_SHOW)

Global $sStatus
GUICtrlCreateListViewItem("MR2012ProcessService|" & $sStatus, $idListview)
GUICtrlCreateListViewItem("DynamicsAxBatch|" & $sStatus, $idListview)
GUICtrlCreateListViewItem("Microsoft.Dynamics.AX.Framework.Tools.DMF.SSISHelperService.exe|" & $sStatus, $idListview)
GUICtrlCreateListViewItem("W3SVC|" & $sStatus, $idListview)

Global $guiAddServicegLabel = GUICtrlCreateLabel("Add service:", 365, 220, 110, 20)
Global $guiRemoveServicegLabel = GUICtrlCreateLabel("Remove service:", 362, 320, 110, 20)

Global $iRunningProcesses = 0

Global $aiExist
Global $aiStatus
#EndRegion ### END Koda GUI section ###


;Check and set service status on startup

For $i = 0 To _GUICtrlListView_GetItemCount($idListview)
	Local $obShell = ObjCreate("shell.application")
	If $obShell.IsServiceRunning(_GUICtrlListView_GetItemText($idListview, $i)) Then
		$sStatus = "Running"
		$iRunningProcesses += 1
		_GUICtrlListView_SetItem($idListview, $sStatus, $i, 1)
	Else
		$sStatus = "Stopped"
		_GUICtrlListView_SetItem($idListview, $sStatus, $i, 1)
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

	;Check if service is already running before trying to start it and set the status respectively

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
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStart, 1)
		Else
			Run("NET START " & $sService)
			WinActivate($hGui)
			Do
				Sleep(100)
			Until $oShell.IsServiceRunning($sService)
			$sStatus = "Running"
			$iRunningProcesses += 1
			GUICtrlSetData($idProgress, 100 / _GUICtrlListView_GetItemCount($idListview) * $iRunningProcesses)
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStart, 1)
			Sleep(500)
		EndIf
	Next
EndFunc   ;==>Start


;stop the services
Func _Stop()
	Local $oShell = ObjCreate("shell.application")
	Local $iStoppedProcesses

	;Check if service is already stopped before trying to stop it and set the status respectively

	For $iNetStop = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
		Local $sServiceStop = _GUICtrlListView_GetItemText($idListview, $iNetStop)
		If Not $oShell.IsServiceRunning($sServiceStop) Then
			$sStatus = "Stopped"
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStop, 1)
		Else
			Run("NET STOP " & $sServiceStop)
			WinActivate($hGui)
			Do
				Sleep(1000)
			Until Not $oShell.IsServiceRunning($sServiceStop)
			$sStatus = "Stopped"
			$iRunningProcesses -= 1
			$iStoppedProcesses += 1
			GUICtrlSetData($idProgress, 100 / _GUICtrlListView_GetItemCount($idListview) * $iStoppedProcesses)
			_GUICtrlListView_SetItem($idListview, $sStatus, $iNetStop, 1)
			Sleep(500)
		EndIf
	Next
	$iStoppedProcesses = 0
EndFunc   ;==>Stop

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
			GUICtrlCreateListViewItem($sServiceAdd & "|" & $sStatus, $idListview)
		EndIf
	EndIf
EndFunc   ;==>addService

Func _removeService()
	_GUICtrlListView_DeleteItemsSelected($idListview)
EndFunc   ;==>removeService