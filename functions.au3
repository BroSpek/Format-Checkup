#Region ;*** GENERAL ***
; Open Control Panel
Func _OCP($sCanonical, $sPage)

  Local Const $sCLSID = "{06622D85-6856-4460-8DE1-A81921B41C4B}"
  Local Const $sIID   = "{D11AD862-66DE-4DF4-BF6C-1F5621996AF1}"
  Local Const $sTag   = "Open hresult(wstr;wstr;ptr);GetPath hresult(wstr;wstr;uint);GetCurrentView hresult(int*)"
  Local $oOCP = ObjCreateInterface($sCLSID, $sIID,  $sTag)
  $oOCP.Open($sCanonical, $sPage, Null)

EndFunc   ;==>_OCP

; Get computer model
Func _GetModel()

  Local $wbemFlagReturnImmediately = 0x10
  Local $wbemFlagForwardOnly = 0x20

	$objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	If IsObj($colItems) then
		For $objItem In $colItems
			Return $objItem.Name
		Next
	Else
		Return '_GetModel error: No WMI Objects Found'
	Endif

EndFunc   ;==>_GetModel

; Compare between two arrays
Func _ArrayCompare($array1, $array2)

  Local $i, $arr1 = 0, $arr2 = 0, $array, $result, $count = 0

  ; Determine if an array has count index
  If $array1[0] = Ubound($array1) - 1 Then $arr1 = 1 ; if it has count index
  If $array2[0] = Ubound($array2) - 1 Then $arr2 = 1

  ; Determine which array has the lowest count
  $array = Ubound($array1) - Ubound($array2)
  If $array < 0 Then
    $array = Ubound($array1) - $arr1 ; if it has count index, reduce the array size count (minus $arr1)
  Else
    $array = Ubound($array2) - $arr2
  EndIf

  ; One by one comparison
  For $i = 0 To $array - 1
    $result = StringCompare($array1[$i + $arr1], $array2[$i + $arr2]) ; ($i + $arr) is to jump to array content if it has count index
    If $result = 0 Then ; if similar
      $count += 1 ; 1 point, then added to the next
    EndIf
  Next

  Return Round($count / $array * 100) ; total point added / total lowest array count * 100(percent)

EndFunc

Func _GetImageDimension($sPassed_File_Name)

  Local $sDir_Name = StringRegExpReplace($sPassed_File_Name, "(^.*\\)(.*)", "\1")
  Local $sFile_Name = StringRegExpReplace($sPassed_File_Name, "^.*\\", "")
  Local $sDOS_Dir = FileGetShortName($sDir_Name, 1)

  Local $oShellApp = ObjCreate("shell.application")
  If IsObj($oShellApp) Then
    Local $oDir = $oShellApp.NameSpace($sDOS_Dir)
    If IsObj($oDir) Then
      Local $oFile = $oDir.Parsename($sFile_Name)
      If IsObj($oFile) Then
        Return $oDir.GetDetailsOf($oFile, 31)
      Else
        SetError(3)
      EndIf
    Else
      SetError(2)
    EndIf
  Else
    SetError(1)
  EndIf

EndFunc

; User's COM error function. Will be called if COM error occurs
; Local $objErrorHandler = ObjEvent("AutoIt.Error", "_ErrHandler")    ; Initialize a COM error handler
Func _ErrHandler($oError)
    MsgBox($MB_TOPMOST, "", @ScriptName & " (Line: " & $oError.scriptline & ")" & @CRLF & _
      "Error 0x" & Hex($oError.number) & @CRLF & _
      $oError.windescription)
    ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
      @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
      @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
      @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
      @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
      @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
      @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
      @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
      @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
      @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_ErrHandler
#EndRegion ;*** GENERAL ***
#Region ;*** ACTIVATION ***

Func _IsOfficeActivated() ; Return True or False

    $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
    If IsObj($objWMIService) Then
        $colItems = $objWMIService.ExecQuery("SELECT Description, LicenseStatus, GracePeriodRemaining FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey <> null")
        If IsObj($colItems) Then
            For $objItem In $colItems
                Switch $objItem.LicenseStatus
                    Case 0
                        ConsoleWrite(" ---UNLICENSED--- " & @CRLF)
                        Return False
                    Case 1
                        ConsoleWrite(" ---LICENSED--- " & @CRLF)
                            ;If licSr = 0 Then
                            ;    WScript.Echo MSG_ERRCODE & licSr & " as licensed"
                            ;End If
                        Return True
                    Case 2
                        ConsoleWrite("---OOB_GRACE--- Initial grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 3
                        ConsoleWrite("---OOT_GRACE--- Initial grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 4
                        ConsoleWrite("---NON_GENUINE_GRACE--- Grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 5
                        ConsoleWrite("---NOTIFICATIONS--- Grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 6
                        ConsoleWrite("---EXTENDED GRACE--- Extended grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case Else
                        ConsoleWrite("---UNKNOWN---" & @CRLF)
                        Return False
                EndSwitch
            Next
        EndIf
    EndIf

EndFunc   ;==>_IsOfficeActivated

Func _IsActivated($product) ; return: true;false, para: windows;office

    $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
    If IsObj($objWMIService) Then
        $colItems = $objWMIService.ExecQuery("SELECT Description, LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE PartialProductKey <> null AND Description LIKE '" & $product & "%'")
        If IsObj($colItems) Then
            For $objItem In $colItems
                Switch $objItem.LicenseStatus
                    Case 0
                        ConsoleWrite("Unlicensed" & @CRLF)
                        Return False
                    Case 1
                        If $objItem.GracePeriodRemaining Then
                            If StringInStr($objItem.Description, "TIMEBASED_") Then
                                ConsoleWrite("Timebased activation will expire in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                                Return True
                            Else
                                ConsoleWrite("Volume activation will expire in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                                Return True
                            EndIf
                        Else
                            ConsoleWrite("The machine is permanently activated." & @CRLF)
                            Return True
                        EndIf
                    Case 2
                        ConsoleWrite("Initial grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 3
                        ConsoleWrite("Additional grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 4
                        ConsoleWrite("Non-genuine grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                    Case 5
                        ConsoleWrite("Windows is in Notification mode" & @CRLF)
                        Return False
                    Case 6
                        ConsoleWrite("Extended grace period ends in " & $objItem.GracePeriodRemaining & " minutes" & @CRLF)
                        Return False
                EndSwitch
            Next
        EndIf
    EndIf

EndFunc   ;==>_IsWindowsActivated

#EndRegion ;*** ACTIVATION ***
#Region ;*** OEM ***
Func _CloseOEM()
  GUISetState(@SW_HIDE, $fOEMInfo)
EndFunc

Func _ClearOEM()

  GUICtrlSetImage($pOEMLogo,     $sFilesPath & "default.gif")
  GUICtrlSetData($iManufacturer, "")
  GUICtrlSetData($iModel,        "")
  GUICtrlSetData($iHours,        "")
  GUICtrlSetData($iPhone,        "")
  GUICtrlSetData($iURL,          "")
  $sOEMLogoTmp = ""

EndFunc

Func _BrowseOEM()

  Local $sImageDim

  $sOEMLogoTmp = FileOpenDialog("Browse", @UserProfileDir & "\", "Image (*.bmp)", $FD_FILEMUSTEXIST)
  If @error Then
    MsgBox(8192+64, "Info", "No image were selected.")
    FileChangeDir(@ScriptDir)
  Else
    $sImageDim = _GetImageDimension($sOEMLogoTmp)
    If @error Then
      MsgBox(8192+64, "Info", "An error occured!")
    Else
      If $sImageDim <> "120 x 120" Then
        MsgBox(8192+64, "Info", "Image is not 120 x 120 pixels.")
      Else
        GUICtrlSetImage($pOEMLogo, $sOEMLogoTmp)
      Endif
    Endif
  EndIf

EndFunc

Func _ReadOEM($sFrom) ; $sFrom: Registry or Ini File: "REG" or "INI" or "INP"

  Local $aOEM[7]

  Switch $sFrom
    Case "REG"
      ; Read from registry
      $aOEM[0] = 6
      $aOEM[1] = RegRead($sOEMKey, "Logo")
      $aOEM[2] = RegRead($sOEMKey, "Manufacturer")
      $aOEM[3] = RegRead($sOEMKey, "Model")
      $aOEM[4] = RegRead($sOEMKey, "SupportHours")
      $aOEM[5] = RegRead($sOEMKey, "SupportPhone")
      $aOEM[6] = RegRead($sOEMKey, "SupportURL")
      ;Return $aOEM
    Case "INI"
      ; Read from ini file
      $aOEM[0] = 6
      $aOEM[1] = IniRead($sSettingsIni, "General", "Logo", "")
      $aOEM[2] = IniRead($sSettingsIni, "General", "Manufacturer", "")
      $aOEM[3] = IniRead($sSettingsIni, "General", "Model", _GetModel())
      $aOEM[4] = IniRead($sSettingsIni, "Support Information", "SupportHours", "")
      $aOEM[5] = IniRead($sSettingsIni, "Support Information", "SupportPhone", "")
      $aOEM[6] = IniRead($sSettingsIni, "Support Information", "SupportURL", "")
      ;Return $aOEM
    Case "INP"
      ; Read from input
      $aOEM[0] = 6
      $aOEM[1] = $sOEMLogoTmp
      $aOEM[2] = GUICtrlRead($iManufacturer)
      $aOEM[3] = GUICtrlRead($iModel)
      $aOEM[4] = GUICtrlRead($iHours)
      $aOEM[5] = GUICtrlRead($iPhone)
      $aOEM[6] = GUICtrlRead($iURL)
      ;Return $aOEM
  EndSwitch

  $sOEMLogoTmp = $aOEM[1]

  If $aOEM[1] = "" Then $aOEM[1] = $sFilesPath & "default.gif"
  GUICtrlSetImage($pOEMLogo,     $aOEM[1])
  GUICtrlSetData($iManufacturer, $aOEM[2])
  GUICtrlSetData($iModel,        $aOEM[3])
  GUICtrlSetData($iHours,        $aOEM[4])
  GUICtrlSetData($iPhone,        $aOEM[5])
  GUICtrlSetData($iURL,          $aOEM[6])


  Return $aOEM

EndFunc   ;==>_ReadOEM

Func _WriteOEM($aOEM, $sTo) ; $aOEM: Input arrays of OEM's, $sTo: Write to, "REG" or "INI"

  Switch $sTo
    Case "REG"
      FileCopy($aOEM[1] , $sWinOEMLogo, 1)
      If @error Then MsgBox(8192+64, "Error", "Cannot copy image!")

      If $sOEMLogoTmp = "" Then
        $aOEM[1] = ""
      Else
        $aOEM[1] = $sWinOEMLogo
      EndIf

      RegWrite($sOEMKey, "Logo",         "REG_SZ", $aOEM[1])
      RegWrite($sOEMKey, "Manufacturer", "REG_SZ", $aOEM[2])
      RegWrite($sOEMKey, "Model",        "REG_SZ", $aOEM[3])
      RegWrite($sOEMKey, "SupportHours", "REG_SZ", $aOEM[4])
      RegWrite($sOEMKey, "SupportPhone", "REG_SZ", $aOEM[5])
      RegWrite($sOEMKey, "SupportURL",   "REG_SZ", $aOEM[6])
    Case "INI"
      FileCopy($aOEM[1] , $sCustomOEMLogo, 1)
      If @error Then MsgBox(8192+64, "Error", "Cannot copy image!")

      If $sOEMLogoTmp = "" Then
        $aOEM[1] = ""
      Else
        $aOEM[1] = StringReplace($sCustomOEMLogo, @ScriptDir & "\", "")
      EndIf

      IniWrite($sSettingsIni, "General" ,            "Logo",         $aOEM[1])
      IniWrite($sSettingsIni, "General" ,            "Manufacturer", $aOEM[2])
      ;IniWrite($sSettingsIni, "General" ,            "Model",        $aOEM[3])
      IniWrite($sSettingsIni, "Support Information", "SupportHours", $aOEM[4])
      IniWrite($sSettingsIni, "Support Information", "SupportPhone", $aOEM[5])
      IniWrite($sSettingsIni, "Support Information", "SupportURL",   $aOEM[6])
  EndSwitch
EndFunc   ;==>_WriteOEM
#EndRegion ;*** OEM ***
#Region ;*** DEVICE MANAGER ***
Func _IsDeviceOK() ; Return array of error code

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
  Local $colItems = $objWMIService.ExecQuery("Select * from Win32_PnPEntity " & "WHERE ConfigManagerErrorCode <> 0")
  Local $sErrorCode

  For $objItem in $colItems
		$sErrorCode &= $objItem.ConfigManagerErrorcode & ","
	Next
  $sErrorCode = StringTrimRight($sErrorCode, 1)
  Return StringSplit($sErrorCode, ",")

EndFunc   ;==>_IsDeviceOK
#EndRegion ;*** DEVICE MANAGER ***
#Region ;*** COMPUTER NAME ***
#cs *******************************************************************************
Function:        _RenameComputer( $sCompName , $sUserName = "" , $sPassword = "" )

Description:   Renames the local computer

Parameter(s):  $sCompName:    The new computer name

                Required Only if PC is joined to Domain:
                    $sUserName:    Username in DOMAIN\UserNamefFormat
                    $sPassword:    Password of the specified account

Returns:        1 - Succeded (Reboot to take effect)

                0 - Invalid parameters
                    @error 2 - Computername contains invalid characters.
                    @error 3 - Failed to create COM Object

                Returns error code returned by WMI
                    Sets @error 1

Author(s):     Kenneth Morrissey (ken82m)
#ce *******************************************************************************

Func _RenameComputer($sCompName, $sUserName = "", $sPassword = "")
    ;ConsoleWrite("$sCompName = '" & $sCompName & "' $sUserName = '" & $sUserName & "' $sPassword = '" & $sPassword & "'" & @CR)

    Local $Check = StringSplit($sCompName, "`~!@#$%^&*()=+_[]{}\|;:.'"",<>/? ")
    ;ConsoleWrite("$Check[0] = " & $Check[0] & "$Check[1] = " & $Check[1] & @CR)

    If $Check[0] > 1 Then
        SetError(2)
        Return 0
    EndIf

    $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
    If @error Then
        SetError(3)
        Return 0
    EndIf

    For $objComputer In $objWMIService.InstancesOf("Win32_ComputerSystem")
        $oReturn = $objComputer.Rename($sCompName,$sPassword,$sUserName)
        If $oReturn <> 0 Then
            SetError(1)
            Return $oReturn
        Else
            Return 1
        EndIf
    Next
EndFunc

Func _ReadComputerName()

    $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
    If @error Then
      SetError(3)
      Return 0
    EndIf

    For $objComputer In $objWMIService.InstancesOf("Win32_ComputerSystem")
      $colComName = $objComputer.Name
        If $colComName = 0 Then
            Return $colComName
        Else
            SetError(1)
        EndIf
    Next

EndFunc   ;==>_ReadComputerName

; Sets computer name without restart!
Func _SetComputerName($sCmpName)
    Local $sLogonKey = "HKLMSOFTWAREMicrosoftWindows NTCurrentVersionWinlogon"
    Local $sCtrlKey = "HKLMSYSTEMCurrentControlSet"
    Local $aRet

    ; RagsRevenge -> http://www.autoitscript.com/forum/index.php?showtopic=54091&view=findpost&p=821901
    If StringRegExp($sCmpName, '|/|:|*|?|"|<|>|.|,|~|!|@|#|$|%|^|&|(|)|;|{|}|_|=|+|[|]|x60' & "|'", 0) = 1 Then Return 0

    ; 5 = ComputerNamePhysicalDnsHostname
    $aRet = DllCall("Kernel32.dll", "BOOL", "SetComputerNameEx", "int", 5, "str", $sCmpName)
    If $aRet[0] = 0 Then Return SetError(1, 0, 0)
    RegWrite($sCtrlKey & "ControlComputernameActiveComputername", "ComputerName", "REG_SZ", $sCmpName)
    RegWrite($sCtrlKey & "ControlComputernameComputername", "ComputerName", "REG_SZ", $sCmpName)
    RegWrite($sCtrlKey & "ServicesTcpipParameters", "Hostname", "REG_SZ", $sCmpName)
    RegWrite($sCtrlKey & "ServicesTcpipParameters", "NV Hostname", "REG_SZ", $sCmpName)
    RegWrite($sLogonKey, "AltDefaultDomainName", "REG_SZ", $sCmpName)
    RegWrite($sLogonKey, "DefaultDomainName", "REG_SZ", $sCmpName)
    RegWrite("HKEY_USERS.DefaultSoftwareMicrosoftWindows MediaWMSDKGeneral", "Computername", "REG_SZ", $sCmpName)

    ; Set Global Environment Variable
    RegWrite($sCtrlKey & "ControlSession ManagerEnvironment", "Computername", "REG_SZ", $sCmpName)
    ; http://msdn.microsoft.com/en-us/library/ms686206%28VS.85%29.aspx
    $aRet = DllCall("Kernel32.dll", "BOOL", "SetEnvironmentVariable", "str", "Computername", "str", $sCmpName)
    If $aRet[0] = 0 Then Return SetError(2, 0, 0)
    ; http://msdn.microsoft.com/en-us/library/ms644952%28VS.85%29.aspx
    $iRet2 = DllCall("user32.dll", "lresult", "SendMessageTimeoutW", "hwnd", 0xffff, "dword", 0x001A, "ptr", 0, _
            "wstr", "Environment", "dword", 0x0002, "dword", 5000, "dword_ptr*", 0)
    If $iRet2[0] = 0 Then Return SetError(2, 0, 0)

    Return 1
EndFunc   ;==>_SetComputerName
#EndRegion ;*** COMPUTER NAME ***
#Region ;*** DRIVE NAME ***
Func _ReadDriveName($sDriveLetter)

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & @ComputerName & "\root\cimv2")
  Local $colItems = $objWMIService.ExecQuery("SELECT VolumeName FROM Win32_LogicalDisk WHERE Name = '" & $sDriveLetter & "'", "WQL", 0x30)

  For $objItem In $colItems
    $sVolName = $objItem.VolumeName
    If $sVolName = "" Then
      Return "Local Disk"
    Else
      Return $sVolName
    EndIf
    ExitLoop
  Next

EndFunc  ;==>_ReadDriveName

Func _RenameDrive($sDriveLetter, $sNewName)

  Local $objShell = ObjCreate("Shell.Application")
  $objShell.NameSpace($sDriveLetter).Self.Name = $sNewName

EndFunc  ;==>_RenameDrive

Func _DateFormat($Date, $style)

    Local $hGui = GUICreate("My GUI get date", 200, 200, 800, 200)
    Local $idDate = GUICtrlCreateDate($Date, 10, 10, 185, 20)
    GUICtrlSendMsg($idDate, 0x1032, 0, $style)
    Local $sReturn = GUICtrlRead($idDate)
    GUIDelete($hGui)
    Return $sReturn

EndFunc  ;==>_DateFormat

Func _WinVer()

	Switch @OSVersion
		Case "WIN_7"
			Return "W7"
		Case "WIN_8"
			Return "W8"
		Case "WIN_81"
			Return "W81"
		Case "WIN_10"
			Return "W10"
	EndSwitch

EndFunc  ;==>_WinVer

Func _WinArch()

	Switch @OSArch
		Case "X64"
			Return "64"
		Case "IA64"
			Return "64"
		Case "X86"
			Return "32"
	EndSwitch

EndFunc  ;==>_WinArch

Func _WinEd()

    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
    Local $colSettings = $objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

    For $objOperatingSystem In $colSettings
		Switch $objOperatingSystem.Caption
			Case "Microsoft Windows 7 Starter "
				Return "S"
			Case "Microsoft Windows 7 Home Basic "
				Return "HB"
			Case "Microsoft Windows 7 Home Premium "
				Return "HP"
			Case "Microsoft Windows 7 Professional "
				Return "P"
			Case "Microsoft Windows 7 Ultimate "
				Return "U"
			Case "Microsoft Windows 8 (Core) "
				Return "C"
			Case "Microsoft Windows 8 Single Language "
				Return "SL"
			Case "Microsoft Windows 8 Pro "
				Return "P"
			Case "Microsoft Windows 8 Enterprise "
				Return "P"
			Case "Microsoft Windows 8.1 (Core) "
				Return "C"
			Case "Microsoft Windows 8.1 Single Language "
				Return "SL"
			Case "Microsoft Windows 8.1 Pro "
				Return "P"
			Case "Microsoft Windows 8.1 Enterprise "
				Return "P"
			Case "Microsoft Windows 10 Home "
				Return "H"
			Case "Microsoft Windows 10 Pro "
				Return "P"
		EndSwitch
	Next

EndFunc  ;==>_WinEd
#EndRegion ;*** DRIVE NAME ***
#Region ;*** SECURITY ***
$sSecurityCenter = "SecurityCenter2"

Func _IsAntivirus() ; Return array of antivirus product

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\" & $sSecurityCenter)
	Local	$colItems = $objWMIService.ExecQuery("Select * from AntiVirusProduct")
  Local $sAVName

	For $objAntiVirusProduct In $colItems
		$sAVName &= $objAntiVirusProduct.displayName & ","
	Next
  $sAVName = StringTrimRight($sAVNAme, 1)
  Return StringSplit($sAVName, ",")

EndFunc   ;==>_IsAntivirus

Func _IsFirewall() ; Return array of firewall product

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\" & $sSecurityCenter)
  Local $colItems = $objWMIService.ExecQuery("Select * from FirewallProduct")
  Local $sFWName

  For $objFireWallProduct In $colItems
    $sFWName &= $objFireWallProduct.displayName & ","
  Next
  $sFWName = StringTrimRight($sFWName, 1)
  Return StringSplit($sFWName, ",")

EndFunc   ;==>_IsFirewall
#EndRegion ;*** SECURITY ***
#Region ;*** SHELL FOLDERS ***
Func _ReadShellFolders() ; Return an array of shell folders

  Local $sShellFolders = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
  Local $aShellFolders[12]

  $aShellFolders[0]  = 11
  $aShellFolders[1]  = RegRead($sShellFolders, "Desktop")
  $aShellFolders[2]  = RegRead($sShellFolders, "Favorites")
  $aShellFolders[3]  = RegRead($sShellFolders, "My Video")
  $aShellFolders[4]  = RegRead($sShellFolders, "My Pictures")
  $aShellFolders[5]  = RegRead($sShellFolders, "My Music")
  $aShellFolders[6]  = RegRead($sShellFolders, "Personal")
  $aShellFolders[7]  = RegRead($sShellFolders, "{374DE290-123F-4565-9164-39C4925E467B}")
  $aShellFolders[8]  = RegRead($sShellFolders, "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}")
  $aShellFolders[9]  = RegRead($sShellFolders, "{56784854-C6CB-462B-8169-88E350ACB882}")
  $aShellFolders[10] = RegRead($sShellFolders, "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}")
  $aShellFolders[11] = RegRead($sShellFolders, "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}")

  Return $aShellFolders

EndFunc   ;==>_ReadShellFolders
#EndRegion ;*** SHELL FOLDERS ***
#Region ;*** WINDOWS UPDATE: NOTIFICATION PAGE ***
Func _AUWrite()

  RegWrite($sAUKey, 'AUOptions', 'REG_DWORD', '1')
  RegWrite($sAUKey, 'IncludeRecommendedUpdates', 'REG_DWORD', '0')
  RegWrite($sAUKey, 'ElevateNonAdmins', 'REG_DWORD', '0')

EndFunc

Func _AURead()

  Global $sAU1, $sAU2, $sAU3

  $sAU1 = RegRead($sAUKey, 'AUOptions')
  $sAU2 = RegRead($sAUKey, 'IncludeRecommendedUpdates')
  $sAU3 = RegRead($sAUKey, 'ElevateNonAdmins')

EndFunc
#EndRegion ;*** WINDOWS UPDATE: NOTIFICATION PAGE ***
#Region ;*** WINDOWS UPDATE: SERVICES PAGE ***
Func _Service_State($strServiceName)

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
	If Not IsObj($objWMIService) then
		MsgBox(64, 'Error!', 'Non object variable')
		Exit
	Endif

	For $colServices in $objWMIService.InstancesOf("Win32_Service")
		If $colServices.Name = $strServiceName Then
			Return $colServices.State
		EndIf
	Next

EndFunc

Func _Service_ChangeState($strServiceName, $strState) ; State: Start, Stop, Pause, Resume

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
	If Not IsObj($objWMIService) then
		MsgBox(64, 'Error!', 'Non object variable')
		Exit
	Endif

	$colServices = $objWMIService.Get('Win32_Service.Name="' & $strServiceName & '"')
	$colState = $colServices.State
	Switch $strState
		Case 'Start'
			If $colState = 'Stopped' Then
				$return = $colServices.StartService()
				Return __ServiceDesc($return)
			Endif
			Return 'No action taken.'
		Case 'Stop'
			If $colState = 'Running' Then
				$return = $colServices.StopService()
				Return __ServiceDesc($return)
			Endif
			Return 'No action taken.'
		Case 'Resume'
			If $colState = 'Paused' Then
				$return = $colServices.ResumeService()
				Return __ServiceDesc($return)
			Endif
			Return 'No action taken.'
		Case 'Pause'
			If $colState = 'Running' Or $colState = 'Stopped' Then
				$return = $colServices.PauseService()
				Return __ServiceDesc($return)
			Endif
			Return 'No action taken.'
	EndSwitch

EndFunc

Func _Service_StartMode($strServiceName)

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
	If Not IsObj($objWMIService) then
		MsgBox(64, 'Error!', 'Non object variable')
		Return 1
	Endif

	$colStartType = $objWMIService.Get("Win32_Service.Name='" & $strServiceName & "'").StartMode
	If Not $colStartType Then
		Return 2
	Else
		Return $colStartType
	EndIf

EndFunc

Func _Service_ChangeStartMode($strServiceName, $strStartMode) ; StartMode: Auto, Manual, Disabled

  Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2")
	If Not IsObj($objWMIService) then
		MsgBox(64, 'Error!', 'Non object variable')
		Exit
	Endif

	$colServices = $objWMIService.Get('Win32_Service.Name="' & $strServiceName & '"')
	$colStartMode = $colServices.StartMode
	If $colStartMode <> $strStartMode Then
		$return = $colServices.ChangeStartMode($strStartMode)
		Return __ServiceDesc($return)
	Endif
	Return 'No action taken'

EndFunc
#EndRegion ;*** WINDOWS UPDATE: SERVICES PAGE ***
#Region ;*** WINDOWS UPDATE: INTERNAL ***
#cs
*********************************************************************************
*             Internal use only, DO NOT CALL this function                      *
*********************************************************************************
#ce
Func __ServiceDesc($rtn) ; $rtn: return code of operation

	Local $strSummary, $strDescrip
	Switch $rtn
		Case 0
			$strSummary = 'Success'
			$strDescrip = 'The request was accepted.'
		Case 1
			$strSummary = 'Not Supported'
			$strDescrip = 'The request is not supported.'
		Case 2
			$strSummary = 'Access Denied'
			$strDescrip = 'The user did not have the necessary access.'
		Case 3
			$strSummary = 'Dependent Services Running'
			$strDescrip = 'The service cannot be stopped because other services that are running are dependent on it.'
		Case 4
			$strSummary = 'Invalid Service Control'
			$strDescrip = 'The requested control code is not valid, or it is unacceptable to the service.'
		Case 5
			$strSummary = 'Service Cannot Accept Control'
			$strDescrip = 'The requested control code cannot be sent to the service because the state of the service (Win32_BaseService.State property) is equal to 0, 1, or 2.'
		Case 6
			$strSummary = 'Service Not Active'
			$strDescrip = 'The service has not been started.'
		Case 7
			$strSummary = 'Service Request Timeout'
			$strDescrip = 'The service did not respond to the start request in a timely fashion.'
		Case 8
			$strSummary = 'Unknown Failure'
			$strDescrip = 'Unknown failure when starting the service.'
		Case 9
			$strSummary = 'Path Not Found'
			$strDescrip = 'The directory path to the service executable file was not found.'
		Case 10
			$strSummary = 'Service Already Running'
			$strDescrip = 'The service is already running.'
		Case 11
			$strSummary = 'Service Database Locked'
			$strDescrip = 'The database to add a new service is locked.'
		Case 12
			$strSummary = 'Service Dependency Deleted'
			$strDescrip = 'A dependency this service relies on has been removed from the system.'
		Case 13
			$strSummary = 'Service Dependency Failure'
			$strDescrip = 'The service failed to find the service needed from a dependent service.'
		Case 14
			$strSummary = 'Service Disabled'
			$strDescrip = 'The service has been disabled from the system.'
		Case 15
			$strSummary = 'Service Logon Failed'
			$strDescrip = 'The service does not have the correct authentication to run on the system.'
		Case 16
			$strSummary = 'Service Marked For Deletion'
			$strDescrip = 'This service is being removed from the system.'
		Case 17
			$strSummary = 'Service No Thread'
			$strDescrip = 'The service has no execution thread.'
		Case 18
			$strSummary = 'Status Circular Dependency'
			$strDescrip = 'The service has circular dependencies when it starts.'
		Case 19
			$strSummary = 'Status Duplicate Name'
			$strDescrip = 'A service is running under the same name.'
		Case 20
			$strSummary = 'Status Invalid Name'
			$strDescrip = 'The service name has invalid characters.'
		Case 21
			$strSummary = 'Status Invalid Parameter'
			$strDescrip = 'Invalid parameters have been passed to the service.'
		Case 22
			$strSummary = 'Status Invalid Service Account'
			$strDescrip = 'The account under which this service runs is either invalid or lacks the permissions to run the service.'
		Case 23
			$strSummary = 'Status Service Exists'
			$strDescrip = 'The service exists in the database of services available from the system.'
		Case 24
			$strSummary = 'Service Already Paused'
			$strDescrip = 'The service is currently paused in the system.'
		Case Else
			$strSummary = 'Other'
			$strDescrip = 'Unknown error'
	EndSwitch
	Return $strSummary ; $strSummary Or $strDescrip

EndFunc   ;==>__ServiceDesc
#EndRegion ;*** WINDOWS UPDATE: INTERNAL ***
