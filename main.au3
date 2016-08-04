#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=checkup.ico
#AutoIt3Wrapper_Outfile=Check (Beta).exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=For use by Bootlab Computers Staff
#AutoIt3Wrapper_Res_Description=Post Format Checkup (Beta)
#AutoIt3Wrapper_Res_Fileversion=0.2.1.0
#AutoIt3Wrapper_Res_ProductVersion=0.2.1.0
#AutoIt3Wrapper_Res_LegalCopyright=© Angah ICT. All rights reserved.
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=ProductName|Post Format Checkup (Beta)
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ;**** Includes added by AutoIt3Wrapper ***
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>
#include <WinAPIFiles.au3>
#include <Date.au3>
#include "Functions.au3"
#EndRegion ;**** Includes added by AutoIt3Wrapper ***

Opt("GUIOnEventMode", 1) ;0=disabled, 1=OnEvent mode enabled
Opt("TrayIconHide", 1) ;0=show, 1=hide tray icon

Global $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrHandler")
Global $sFilesPath = @ScriptDir & "\files\"
Global $sWinOEMLogo = @WindowsDir & "\OEMLogo.bmp"
Global $sSettingsIni = @ScriptDir & "\settings.ini"
Global $sCustomOEMLogo = $sFilesPath & "OEMLogo.bmp"
Global $iCheck = $sFilesPath & "refresh.ico"
Global $iApply = $sFilesPath & "apply.ico"
Global $iShow = $sFilesPath & "open.ico"
Global $iEdit = $sFilesPath & "edit.ico"
Global $iGreen = $sFilesPath & "circle-green.ico"
Global $iYellow = $sFilesPath & "circle-yellow.ico"
Global $iRed = $sFilesPath & "circle-red.ico"
Global $iGrey = $sFilesPath & "circle-grey.ico"
Global $iWindows = $sFilesPath & "windows.ico"
Global $iOffice = $sFilesPath & "office.ico"
Global $iWarning = $sFilesPath & "warning.ico"
Global $sOEMKey, $sAUKey, $sOEMLogo, $sOEMLogoTmp, $pOEMLogo, $iManufacturer, $iModel, $iHours, $iPhone, $iURL

If @OSArch = "X86" Then
	$sOEMKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
	$sAUKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
Else
	$sOEMKey = "HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
	$sAUKey = "HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
EndIf

#Region ; Various Checkings

; Check Windows Version
If Not (@OSVersion = "WIN_10" Or @OSVersion = "WIN_81" Or @OSVersion = "WIN_8" Or @OSVersion = "WIN_7") Then
	MsgBox(64, "Info", "Unsupported Operating System")
	Exit
EndIf
; Check missing files
If IniRead($sSettingsIni, "Custom", "Warning", "") <> "No" Then
	If Not FileExists(@ScriptDir & "\settings.ini") Then
		MsgBox(48, "Warning", "Settings file missing")
	EndIf
	If Not FileExists($sFilesPath & "\Microsoft Toolkit.exe") Then
		MsgBox(48, "Warning", "Microsoft Toolkit missing")
	EndIf
	If Not FileExists($sFilesPath & "\Windows Loader.exe") Then
		MsgBox(48, "Warning", "Windows Loader missing")
	EndIf
EndIf

#EndRegion ; Various Checkings

#Region ; Form 1

$guiForm = GUICreate("Post Format Checkup", 660, 460)

$lIntro = GUICtrlCreateLabel("This is an overview of current computer various item to be checked after format. Use buttons to check and apply neccessary patch.", 20, 10, 620, 50)

$gCompName = GUICtrlCreateGroup("Computer Name", 20, 72, 200, 90)
$bCheckCompName = GUICtrlCreateButton("Check", 138, 128, 35, 25, $BS_ICON)
$bApplyCompName = GUICtrlCreateButton("Apply", 173, 128, 35, 25, $BS_ICON)
$iCompName1 = GUICtrlCreateIcon("", -1, 30, 102, 16, 16)
$lCompName1 = GUICtrlCreateLabel("", 55, 102, 150, 17)
GUICtrlSetImage($bCheckCompName, $iCheck, -1, 0)
GUICtrlSetImage($bApplyCompName, $iApply, -1, 0)
GUICtrlSetTip($bCheckCompName, "Re-check computer's name")
GUICtrlSetTip($bApplyCompName, "Apply computer's name")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gDriveName = GUICtrlCreateGroup("Drive Name", 230, 72, 200, 90)
$bCheckDriveName = GUICtrlCreateButton("Check", 348, 128, 35, 25, $BS_ICON)
$bApplyDriveName = GUICtrlCreateButton("Apply", 383, 128, 35, 25, $BS_ICON)
$iDriveName1 = GUICtrlCreateIcon("", -1, 240, 102, 16, 16)
$lDriveName1 = GUICtrlCreateLabel("", 265, 102, 150, 17)
GUICtrlSetImage($bCheckDriveName, $iCheck, -1, 0)
GUICtrlSetImage($bApplyDriveName, $iApply, -1, 0)
GUICtrlSetTip($bCheckDriveName, "Re-check drive's name")
GUICtrlSetTip($bApplyDriveName, "Apply drive's name")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gAntivirus = GUICtrlCreateGroup("Antivirus", 440, 72, 200, 90)
$bCheckAV = GUICtrlCreateButton("Check", 593, 128, 35, 25, $BS_ICON)
$iAV1 = GUICtrlCreateIcon("", -1, 450, 102, 16, 16)
$lAV1 = GUICtrlCreateLabel("", 475, 102, 150, 17)
GUICtrlSetImage($bCheckAV, $iCheck, -1, 0)
GUICtrlSetTip($bCheckAV, "Re-check antivirus's status")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gDevice = GUICtrlCreateGroup("Device Manager", 20, 172, 200, 90)
$bCheckDevice = GUICtrlCreateButton("Check", 138, 228, 35, 25, $BS_ICON)
$bShowDevice = GUICtrlCreateButton("Open", 173, 228, 35, 25, $BS_ICON)
$iDevice1 = GUICtrlCreateIcon("", -1, 30, 202, 16, 16)
$lDevice1 = GUICtrlCreateLabel("", 55, 202, 150, 17)
GUICtrlSetImage($bShowDevice, $iShow, -1, 0)
GUICtrlSetImage($bCheckDevice, $iCheck, -1, 0)
GUICtrlSetTip($bCheckDevice, "Re-check driver's status")
GUICtrlSetTip($bShowDevice, "Open Device Manager")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gSHFolders = GUICtrlCreateGroup("11 Folders", 230, 172, 200, 90)
$bCheckSHFolders = GUICtrlCreateButton("Check", 348, 228, 35, 25, $BS_ICON)
$bShowSHFolders = GUICtrlCreateButton("Show", 383, 228, 35, 25, $BS_ICON)
$iSHFolders1 = GUICtrlCreateIcon("", -1, 240, 202, 16, 16)
$lSHFolders1 = GUICtrlCreateLabel("", 265, 202, 150, 17)
GUICtrlSetImage($bCheckSHFolders, $iCheck, -1, 0)
GUICtrlSetImage($bShowSHFolders, $iShow, -1, 0)
GUICtrlSetTip($bCheckSHFolders, "Re-check User folder's status")
GUICtrlSetTip($bShowSHFolders, "Open User's folder")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gOEMInfo = GUICtrlCreateGroup("OEM Info", 440, 172, 200, 90)
$bEditOEMInfo = GUICtrlCreateButton("Edit", 488, 228, 35, 25, $BS_ICON)
$bCheckOEMInfo = GUICtrlCreateButton("Check", 523, 228, 35, 25, $BS_ICON)
$bShowOEMInfo = GUICtrlCreateButton("Show", 558, 228, 35, 25, $BS_ICON)
$bApplyOEMInfo = GUICtrlCreateButton("Apply", 593, 228, 35, 25, $BS_ICON)
$iOEMInfo1 = GUICtrlCreateIcon("", -1, 450, 202, 16, 16)
$lOEMInfo1 = GUICtrlCreateLabel("", 475, 202, 150, 17)
GUICtrlSetImage($bEditOEMInfo, $iEdit, -1, 0)
GUICtrlSetImage($bCheckOEMInfo, $iCheck, -1, 0)
GUICtrlSetImage($bShowOEMInfo, $iShow, -1, 0)
GUICtrlSetImage($bApplyOEMInfo, $iApply, -1, 0)
GUICtrlSetTip($bCheckOEMInfo, "Customize OEM's Information")
GUICtrlSetTip($bCheckOEMInfo, "Re-check OEM's Information")
GUICtrlSetTip($bApplyOEMInfo, "Apply OEM's Information")
GUICtrlSetTip($bShowOEMInfo, "Show System Page (OEM Information)")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gUpdate = GUICtrlCreateGroup("Windows Update", 20, 272, 200, 120)
$bCheckUpdate = GUICtrlCreateButton("Check", 103, 360, 35, 25, $BS_ICON)
$bShowUpdate = GUICtrlCreateButton("Show", 138, 360, 35, 25, $BS_ICON)
$bApplyUpdate = GUICtrlCreateButton("Apply", 173, 360, 35, 25, $BS_ICON)
$iUpdate1 = GUICtrlCreateIcon("", -1, 30, 302, 16, 16)
$iUpdate2 = GUICtrlCreateIcon("", -1, 30, 332, 16, 16)
$lUpdate1 = GUICtrlCreateLabel("", 55, 302, 150, 17)
$lUpdate2 = GUICtrlCreateLabel("", 55, 332, 150, 17)
GUICtrlSetImage($bCheckUpdate, $iCheck, -1, 0)
GUICtrlSetImage($bShowUpdate, $iShow, -1, 0)
GUICtrlSetImage($bApplyUpdate, $iApply, -1, 0)
GUICtrlSetTip($bCheckUpdate, "Re-check Windows Update's status")
GUICtrlSetTip($bApplyUpdate, "Apply Windows Update's tweak")
GUICtrlSetTip($bShowUpdate, "Show Windows Update Setting's page")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gActivation = GUICtrlCreateGroup("Activation", 230, 272, 200, 120)
$bCheckActivation = GUICtrlCreateButton("Check", 313, 360, 35, 25, $BS_ICON)
$bOfficeActivation = GUICtrlCreateButton("Office", 348, 360, 35, 25, $BS_ICON)
$bWindowsActivation = GUICtrlCreateButton("Windows", 383, 360, 35, 25, $BS_ICON)
$iActivation1 = GUICtrlCreateIcon("", -1, 240, 302, 16, 16)
$iActivation2 = GUICtrlCreateIcon("", -1, 240, 332, 16, 16)
$lActivation1 = GUICtrlCreateLabel("", 265, 302, 150, 17)
$lActivation2 = GUICtrlCreateLabel("", 265, 332, 150, 17)
GUICtrlSetImage($bCheckActivation, $iCheck, -1, 0)
GUICtrlSetImage($bOfficeActivation, $iOffice, -1, 0)
GUICtrlSetImage($bWindowsActivation, $iWindows, -1, 0)
GUICtrlSetTip($bCheckActivation, "Re-check activation's status")
GUICtrlSetTip($bOfficeActivation, "Run Office activation tool")
GUICtrlSetTip($bWindowsActivation, "Run Windows activation tool")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gForm = GUICtrlCreateGroup("This section will apply to ALL", 20, 400, 200, 50)
$bCheckForm = GUICtrlCreateButton("Check All", 25, 418, 35, 25, $BS_ICON)
$bApplyForm = GUICtrlCreateButton("Apply All", 60, 418, 35, 25, $BS_ICON)
GUICtrlSetImage($bCheckForm, $iCheck, -1, 0)
GUICtrlSetImage($bApplyForm, $iApply, -1, 0)
GUICtrlSetTip($bCheckForm, "Check all")
GUICtrlSetTip($bApplyForm, "Patch all")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$bAbout = GUICtrlCreateButton("About", 585, 418, 60, 25)

#EndRegion ; Form 1

#Region ; Form 2
$fOEMInfo = GUICreate("OEM Info Customizer", 450, 260, -1, -1, $WS_POPUPWINDOW)
$iManufacturer = GUICtrlCreateInput("Manufacturer", 88, 50, 200, 21)
$iModel = GUICtrlCreateInput("Model", 88, 77, 200, 21)
$iHours = GUICtrlCreateInput("Support Hours", 88, 104, 200, 21)
$iPhone = GUICtrlCreateInput("Support Phone", 88, 131, 200, 21)
$iURL = GUICtrlCreateInput("Support URL", 88, 158, 200, 21)
$bCurrent = GUICtrlCreateButton("Current", 10, 220, 65, 25)
$bLoad = GUICtrlCreateButton("Load", 75, 220, 65, 25)
$bClear = GUICtrlCreateButton("Clear", 140, 220, 65, 25)
$bSave = GUICtrlCreateButton("Save", 205, 220, 65, 25)
$bApply = GUICtrlCreateButton("Apply", 270, 220, 65, 25)
$bClose = GUICtrlCreateButton("Close", 372, 220, 65, 25)
$lManufacturer = GUICtrlCreateLabel("Manufacturer", 10, 53, 67, 17)
$lDivide = GUICtrlCreateLabel("", 10, 210, 428, 2, $SS_SUNKEN)
$lModel = GUICtrlCreateLabel("Model", 10, 80, 33, 17)
$lHours = GUICtrlCreateLabel("Hours", 10, 107, 32, 17)
$lPhone = GUICtrlCreateLabel("Phone", 10, 134, 35, 17)
$lURL = GUICtrlCreateLabel("URL", 10, 161, 26, 17)
$pOEMLogo = GUICtrlCreatePic($sFilesPath & "default.gif", 310, 54, 120, 120)
$lImage = GUICtrlCreateLabel("120 x 120 BMP Image", 318, 180, 118, 25)
$lTitle = GUICtrlCreateLabel("OEM Information Customizer", 104, 8, 235, 24)
GUICtrlSetFont($lTitle, 12, 800, 0, "MS Sans Serif")
GUICtrlSetState($iModel, $GUI_DISABLE)

_ReadOEM("REG")

#EndRegion ; Form 2

#Region ; SPLASH!

$frmSplash = GUICreate("", 300, 150, -1, -1, BitOR($WS_SYSMENU, $WS_POPUP), 0)
$picSplash = GUICtrlCreatePic($sFilesPath & "splash.jpg", 0, 0, 300, 150, BitOR($SS_NOTIFY, $WS_GROUP, $WS_CLIPSIBLINGS))
$txtSplash = GUICtrlCreateLabel("", 56, 80, 200, 15)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 8)
GUICtrlSetColor(-1, '0xffffff')
$proSplash = GUICtrlCreateProgress(56, 95, 193, 3)
GUISetState(@SW_SHOW, $frmSplash)
_SplashBar($proSplash, $txtSplash)
GUIDelete($frmSplash)

#EndRegion ; SPLASH!

GUISetState(@SW_SHOW, $guiForm)

#Region ; Lots of buttons

GUISetOnEvent($GUI_EVENT_CLOSE, "QUIT")
GUICtrlSetOnEvent($bApplyCompName, "MAIN")
GUICtrlSetOnEvent($bCheckCompName, "MAIN")
GUICtrlSetOnEvent($bApplyDriveName, "MAIN")
GUICtrlSetOnEvent($bCheckDriveName, "MAIN")
GUICtrlSetOnEvent($bCheckAV, "MAIN")
GUICtrlSetOnEvent($bCheckDevice, "MAIN")
GUICtrlSetOnEvent($bShowDevice, "MAIN")
GUICtrlSetOnEvent($bCheckSHFolders, "MAIN")
GUICtrlSetOnEvent($bShowSHFolders, "MAIN")
GUICtrlSetOnEvent($bCheckUpdate, "MAIN")
GUICtrlSetOnEvent($bShowUpdate, "MAIN")
GUICtrlSetOnEvent($bApplyUpdate, "MAIN")
GUICtrlSetOnEvent($bCheckActivation, "MAIN")
GUICtrlSetOnEvent($bOfficeActivation, "MAIN")
GUICtrlSetOnEvent($bWindowsActivation, "MAIN")
GUICtrlSetOnEvent($bEditOEMInfo, "MAIN")
GUICtrlSetOnEvent($bCheckOEMInfo, "MAIN")
GUICtrlSetOnEvent($bShowOEMInfo, "MAIN")
GUICtrlSetOnEvent($bApplyOEMInfo, "MAIN")
GUICtrlSetOnEvent($bCheckForm, "MAIN")
GUICtrlSetOnEvent($bApplyForm, "MAIN")
GUICtrlSetOnEvent($bAbout, "MAIN")

GUICtrlSetOnEvent($pOEMLogo, "MAIN")
GUICtrlSetOnEvent($bSave, "MAIN")
GUICtrlSetOnEvent($bLoad, "MAIN")
GUICtrlSetOnEvent($bClear, "MAIN")
GUICtrlSetOnEvent($bApply, "MAIN")
GUICtrlSetOnEvent($bClose, "MAIN")
GUICtrlSetOnEvent($bCurrent, "MAIN")

#EndRegion ; Lots of buttons

While 1
	Sleep(500)
WEnd

Func MAIN()
	Switch @GUI_CtrlId
		Case $bApplyCompName
			_ComputerName('Apply')
			If @error = 1 Then MsgBox(8192, "Info", "Computer's name was already changed.")
		Case $bCheckCompName
			_ComputerName('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bApplyDriveName
			_DriveName('Apply')
			If @error = 1 Then MsgBox(8192, "Info", "Root drive name was already changed.")
		Case $bCheckDriveName
			_DriveName('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bCheckAV
			_Antivirus('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bCheckDevice
			_Device('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bShowDevice
			_Device('Show')
		Case $bCheckSHFolders
			_ShellFolders('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bShowSHFolders
			_ShellFolders('Show')
		Case $bCheckUpdate
			_Update('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bShowUpdate
			_Update('Show')
		Case $bApplyUpdate
			_Update('Apply')
		Case $bCheckActivation
			_Activation('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bOfficeActivation
			_Activation('Office')
		Case $bWindowsActivation
			_Activation('Windows')
		Case $bEditOEMInfo
			_OEMInfo('Edit')
		Case $bCheckOEMInfo
			_OEMInfo('Check')
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bShowOEMInfo
			_OEMInfo('Show')
		Case $bApplyOEMInfo
			_OEMInfo('Apply')
			If @error = 1 Then MsgBox(8192, "Info", "Custom OEM Info was already applied.")
		Case $bCheckForm
			GUICtrlSetState($bCheckForm, $GUI_DISABLE)
			_CheckAll()
			MsgBox(64 + 8192, "Info", "Done!")
			GUICtrlSetState($bCheckForm, $GUI_ENABLE)
		Case $bApplyForm
			GUICtrlSetState($bApplyForm, $GUI_DISABLE)
			_ApplyAll()
			MsgBox(64 + 8192, "Info", "Done!")
			GUICtrlSetState($bApplyForm, $GUI_ENABLE)
		Case $bAbout
			_About()
		Case $bSave
			_WriteOEM(_ReadOEM("INP"), "INI")
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bLoad
			_ReadOEM("INI")
		Case $bClear
			_ClearOEM()
		Case $bApply
			_WriteOEM(_ReadOEM("INP"), "REG")
			MsgBox(64 + 8192, "Info", "Done!")
		Case $bCurrent
			_ReadOEM("REG")
		Case $bClose
			_CloseOEM()
		Case $pOEMLogo
			_BrowseOEM()
	EndSwitch
EndFunc   ;==>MAIN

Func _ComputerName($CASE)

	Local $sDate = _NowDate()
	Local $sCmpValue = IniRead($sSettingsIni, "Custom", "Initial", "")
	Local $sCompare = StringInStr(_ReadComputerName(), $sCmpValue)
	Local $sNewName = $sCmpValue & _DateFormat($sDate, "yyMMddHHmm")

	Switch $CASE
		Case "Apply"
			Switch $sCompare
				Case 0
					MsgBox(8192, "Info", "Computer's name will change to: " & $sNewName)
					;_SetComputerName($sNewName)
					$temp = _RenameComputer($sNewName)
					Switch $temp
						Case 0
							Switch @error
								Case 2
									MsgBox(8192 + 64, "Error", "Computername contains invalid characters.")
								Case 3
									MsgBox(8192 + 64, "Error", "Failed to create COM Object.")
							EndSwitch
						Case 1
							MsgBox(8192 + 64, "Success", "Reboot to take effect.")
					EndSwitch
					GUICtrlSetImage($iCompName1, $iYellow, 0)
				Case Else
					SetError(1)
					;MsgBox(8192, "Info", "Computer's name was already changed.")
			EndSwitch
		Case "Check"
			GUICtrlSetData($lCompName1, _ReadComputerName())
			Switch $sCompare
				Case 0
					GUICtrlSetImage($iCompName1, $iRed, 0)
				Case Else
					GUICtrlSetImage($iCompName1, $iGreen, 0)
			EndSwitch
	EndSwitch

EndFunc   ;==>_ComputerName

Func _DriveName($CASE)

	Local $sDate = _NowDate()
	Local $sCmpValue = _WinVer() & _WinEd() & _WinArch()
	Local $sCompare = StringInStr(_ReadDriveName(@HomeDrive), $sCmpValue)
	Local $sNewName = $sCmpValue & _DateFormat($sDate, "yyMMdd")

	Switch $CASE
		Case "Apply"
			Switch $sCompare
				Case 0
					MsgBox(8192, "Info", "Current root drive name will change to: " & $sNewName)
					_RenameDrive(@HomeDrive, $sNewName)
					GUICtrlSetData($lDriveName1, _ReadDriveName(@HomeDrive))
					GUICtrlSetImage($iDriveName1, $iGreen, 0)
				Case Else
					SetError(1)
					;MsgBox(8192,"Info", "Root drive name was already changed.")
			EndSwitch
		Case "Check"
			GUICtrlSetData($lDriveName1, _ReadDriveName(@HomeDrive))
			Switch $sCompare
				Case 0
					GUICtrlSetImage($iDriveName1, $iRed, 0)
				Case Else
					GUICtrlSetImage($iDriveName1, $iGreen, 0)
			EndSwitch
	EndSwitch

EndFunc   ;==>_DriveName

Func _Antivirus($CASE)

	Local $aAntivirus = _IsAntivirus()

	Switch $CASE
		Case "Check"
			If $aAntivirus[0] = 0 Then
				GUICtrlSetData($lAV1, "Antivirus: absent")
				GUICtrlSetImage($iAV1, $iRed, 0)
			ElseIf $aAntivirus[0] = 1 Then
				If $aAntivirus[1] = "" Then
					GUICtrlSetData($lAV1, "Antivirus: absent")
					GUICtrlSetImage($iAV1, $iRed, 0)
				Else
					GUICtrlSetData($lAV1, $aAntivirus[1])
					GUICtrlSetImage($iAV1, $iGreen, 0)
				EndIf
			ElseIf $aAntivirus[0] > 1 Then
				GUICtrlSetData($lAV1, "Antivirus: multiple")
				GUICtrlSetImage($iAV1, $iYellow, 0)
			EndIf
	EndSwitch

EndFunc   ;==>_Antivirus

Func _Device($CASE)

	Local $aErrorCode = _IsDeviceOK()

	Switch $CASE
		Case "Show"
			Run(@ComSpec & " /c " & "devmgmt.msc", "", @SW_HIDE)
		Case "Check"
			If $aErrorCode[0] = 1 Then
				If Not $aErrorCode[1] Then
					GUICtrlSetData($lDevice1, "Status: good")
					GUICtrlSetImage($iDevice1, $iGreen, 0)
				Else
					GUICtrlSetData($lDevice1, "Status: moderate")
					GUICtrlSetImage($iDevice1, $iYellow, 0)
				EndIf
			ElseIf $aErrorCode[0] > 1 And $aErrorCode[0] < 4 Then
				GUICtrlSetData($lDevice1, "Status: moderate")
				GUICtrlSetImage($iDevice1, $iYellow, 0)
			ElseIf $aErrorCode[0] >= 4 Then
				GUICtrlSetData($lDevice1, "Status: serious")
				GUICtrlSetImage($iDevice1, $iRed, 0)
			EndIf
	EndSwitch

EndFunc   ;==>_Device

Func _ShellFolders($CASE)

	Local $aSHFolders = _ReadShellFolders()
	Local $sDrive = @HomeDrive
	Local $sFolder = @UserProfileDir
	Local $count = 0

	For $i = 1 To $aSHFolders[0]
		$count += StringInStr($aSHFolders[$i], $sDrive)
	Next

	Switch $CASE
		Case "Show"
			Run("Explorer.exe " & $sFolder)
		Case "Check"
			If $count = $aSHFolders[0] Then
				GUICtrlSetData($lSHFolders1, "No folder relocated")
				GUICtrlSetImage($iSHFolders1, $iRed, 0)
			ElseIf $count > 0 And $count < $aSHFolders[0] Then
				GUICtrlSetData($lSHFolders1, "Some folders still remain")
				GUICtrlSetImage($iSHFolders1, $iYellow, 0)
			ElseIf $count = 0 Then
				GUICtrlSetData($lSHFolders1, "All folders relocated")
				GUICtrlSetImage($iSHFolders1, $iGreen, 0)
			Else
				GUICtrlSetData($lSHFolders1, "How many folders are they?")
				GUICtrlSetImage($iSHFolders1, $iWarning, 0
			EndIf
	EndSwitch

EndFunc   ;==>_ShellFolders

Func _Update($CASE)

	Local $strQueryService = "wuauserv" ;The service name to query for...
	Local $sStartMode = _Service_StartMode($strQueryService)

	Switch $CASE
		Case "Show"
			_OCP("Microsoft.WindowsUpdate", "PageSettings")
		Case "Check"
			; Check windows update service
			If $sStartMode = "Manual" Then
				GUICtrlSetData($lUpdate1, "Service: manual")
				GUICtrlSetImage($iUpdate1, $iGreen, 0)
			ElseIf $sStartMode = "Disabled" Then
				GUICtrlSetData($lUpdate1, "Service: disabled")
				GUICtrlSetImage($iUpdate1, $iGreen, 0)
			ElseIf $sStartMode = "Auto" Then
				GUICtrlSetData($lUpdate1, "Service: automatic")
				GUICtrlSetImage($iUpdate1, $iRed, 0)
			Else
				GUICtrlSetData($lUpdate1, "Service: unknown")
				GUICtrlSetImage($iUpdate1, $iGrey, 0)
			EndIf

			; Check windows update notification
			_AURead()
			Switch $sAU1
				Case 1
					GUICtrlSetData($lUpdate2, "Notification: disabled")
					GUICtrlSetImage($iUpdate2, $iGreen, 0)
				Case 2, 3, 4
					GUICtrlSetData($lUpdate2, "Notification: enabled")
					GUICtrlSetImage($iUpdate2, $iRed, 0)
				Case Else
					GUICtrlSetData($lUpdate2, "Notification: unknown")
					GUICtrlSetImage($iUpdate2, $iGrey, 0)
			EndSwitch
		Case "Apply"
			_AUWrite()
			_Service_ChangeState($strQueryService, "Stop")
			_Service_ChangeStartMode($strQueryService, "Disabled")
			GUICtrlSetData($lUpdate1, "Service: disabled")
			GUICtrlSetImage($iUpdate1, $iGreen, 0)
			GUICtrlSetData($lUpdate2, "Notification: disabled")
			GUICtrlSetImage($iUpdate2, $iGreen, 0)
	EndSwitch

EndFunc   ;==>_Update

Func _Activation($CASE)

	Local $crack, $IsOfficeActiv
	If @OSVersion = "WIN_7" Then
		$crack = $sFilesPath & "Windows Loader.exe"
		$IsOfficeActiv = _IsOfficeActivated()
	Else
		$crack = $sFilesPath & "Microsoft Toolkit.exe"
		$IsOfficeActiv = _IsActivated("office")
	EndIf

	Switch $CASE
		Case "Check"
			; Check Windows activation
			Switch _IsActivated("windows")
				Case True
					GUICtrlSetData($lActivation1, "Windows: activated")
					GUICtrlSetImage($iActivation1, $iGreen, 0)
				Case False
					GUICtrlSetData($lActivation1, "Windows: not activated")
					GUICtrlSetImage($iActivation1, $iRed, 0)
			EndSwitch

			; Check Office activation
			Switch $IsOfficeActiv
				Case True
					GUICtrlSetData($lActivation2, "Office: activated")
					GUICtrlSetImage($iActivation2, $iGreen, 0)
				Case False
					GUICtrlSetData($lActivation2, "Office: not activated")
					GUICtrlSetImage($iActivation2, $iRed, 0)
			EndSwitch
		Case "Office"
			GUICtrlSetState($bOfficeActivation, $GUI_DISABLE)
			Run("files\Microsoft Toolkit.exe")
			Sleep(20000)
			GUICtrlSetState($bOfficeActivation, $GUI_ENABLE)
		Case "Windows"
			GUICtrlSetState($bWindowsActivation, $GUI_DISABLE)
			Run($crack)
			Sleep(20000)
			GUICtrlSetState($bWindowsActivation, $GUI_ENABLE)
	EndSwitch

EndFunc   ;==>_Activation

Func _OEMInfo($CASE)

	Local $IsOEMApplied = False
	Local $aRegOEM = _ReadOEM("REG")
	Local $aIniOEM = _ReadOEM("INI")
	Local $Threshold = IniRead($sSettingsIni, "Custom", "OEM", "")

	$IsOEMApplied = _ArrayCompare($aRegOEM, $aIniOEM)
	If $Threshold < 20 Or $Threshold > 80 Then
		MsgBox($MB_TOPMOST, "Settings error", "Wrong threshold value")
		Return 1
	Else
		If $IsOEMApplied >= $Threshold Then
			$IsOEMApplied = True
		Else
			$IsOEMApplied = False
		EndIf
	EndIf

	Switch $CASE
		Case "Edit"
			GUISetState(@SW_SHOW, $fOEMInfo)
		Case "Check"
			Switch $IsOEMApplied
				Case True
					GUICtrlSetData($lOEMInfo1, "OEM Info: applied")
					GUICtrlSetImage($iOEMInfo1, $iGreen, 0)
				Case False
					GUICtrlSetData($lOEMInfo1, "OEM Info: others")
					GUICtrlSetImage($iOEMInfo1, $iRed, 0)
			EndSwitch
		Case "Show"
			_OCP("Microsoft.System", "")
		Case "Apply"
			; Read from ini file and write to registry
			Switch $IsOEMApplied
				Case True
					SetError(1)
					;MsgBox(8192, "Info", "Custom OEM Info was already applied")
				Case False
					_WriteOEM(_ReadOEM("INI"), "REG")
					GUICtrlSetData($lOEMInfo1, "OEM Info: applied")
					GUICtrlSetImage($iOEMInfo1, $iGreen, 0)
					MsgBox(64 + 8192, "Info", "Done")
			EndSwitch
	EndSwitch

EndFunc   ;==>_OEMInfo

Func _CheckAll()

	Local $sTotalTask = 8

	ProgressOn("Post Format Checkup", "Checking ", "", -1, -1, 1)
	ProgressSet(1 / $sTotalTask * 100, "Checking computer's name")
	_ComputerName("Check")
	Sleep(50)

	ProgressSet(2 / $sTotalTask * 100, "Checking drive's name")
	_DriveName("Check")
	Sleep(50)

	ProgressSet(3 / $sTotalTask * 100, "Checking antivirus")
	_Antivirus("Check")
	Sleep(50)

	ProgressSet(4 / $sTotalTask * 100, "Checking device manager")
	_Device("Check")
	Sleep(50)

	ProgressSet(5 / $sTotalTask * 100, "Checking shell folders")
	_ShellFolders("Check")
	Sleep(50)

	ProgressSet(6 / $sTotalTask * 100, "Checking OEM Information")
	_OEMInfo("Check")
	Sleep(50)

	ProgressSet(7 / $sTotalTask * 100, "Checking windows update")
	_Update("Check")
	Sleep(50)

	ProgressSet(8 / $sTotalTask * 100, "Checking product activation")
	_Activation("Check")
	Sleep(50)
	ProgressOff()

EndFunc   ;==>_CheckAll

Func _SplashBar($sPro, $sTxt)

	Local $iTotalTask = 8
	Local $iPercent = 100 / $iTotalTask

	GUICtrlSetData($sPro, 1 * $iPercent)
	GUICtrlSetData($sTxt, "Checking computer name")
	_ComputerName("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 2 * $iPercent)
	GUICtrlSetData($sTxt, "Checking root drive name")
	_DriveName("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 3 * $iPercent)
	GUICtrlSetData($sTxt, "Checking antivirus")
	_Antivirus("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 4 * $iPercent)
	GUICtrlSetData($sTxt, "Checking device driver")
	_Device("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 5 * $iPercent)
	GUICtrlSetData($sTxt, "Checking folders allocation")
	_ShellFolders("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 6 * $iPercent)
	GUICtrlSetData($sTxt, "Checking OEM Information")
	_OEMInfo("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 7 * $iPercent)
	GUICtrlSetData($sTxt, "Checking Windows update settings")
	_Update("Check")
	Sleep(50)

	GUICtrlSetData($sPro, 8 * $iPercent)
	GUICtrlSetData($sTxt, "Checking activation")
	_Activation("Check")
	Sleep(50)

EndFunc   ;==>_SplashBar

Func _ApplyAll()

	_ComputerName("Apply")
	_DriveName("Apply")
	_OEMInfo("Apply")
	_Update("Apply")

EndFunc   ;==>_ApplyAll

Func _About()

	MsgBox($MB_TOPMOST, "About", "Post Format Checkup (Beta)" & @LF & _
			@LF & _
			"© Angah ICT, 2016." & @LF & _
			@LF & _
			"Program created by Abd Halim." & @LF & _
			"Icons by Sublink Interactive [sublink.ca]." & @LF & _
			@LF & _
			"Hopes this tool helps you!")

EndFunc   ;==>_About

Func QUIT()
	Exit
EndFunc   ;==>QUIT
