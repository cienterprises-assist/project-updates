Option Explicit

' Registry Paths
Const RegPathChrome = "HKLM\SOFTWARE\Policies\Google\Chrome\"
Const RegPathURLAllowlist = "HKLM\SOFTWARE\Policies\Google\Chrome\URLAllowlist\"
Const RegPathURLBlocklist = "HKLM\SOFTWARE\Policies\Google\Chrome\URLBlocklist\"

' Check for administrative privileges
If Not IsAdmin Then
    WScript.Echo "This script must be run as Administrator! Please run as Administrator.", 0, "Error", 16
    WScript.Quit 1
End If

' Set up logging
Dim FSO, LogFile, TempFolder, LogPath
Set FSO = CreateObject("Scripting.FileSystemObject")
TempFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
LogPath = TempFolder & "\ChromePolicyUpdate.log"
On Error Resume Next
Set LogFile = FSO.CreateTextFile(LogPath, True)
If Err.Number <> 0 Then
    WScript.Echo "Error creating log file: " & Err.Description, 0, "Error", 16
    WScript.Quit 1
End If
On Error GoTo 0
LogFile.WriteLine "Chrome Policy Update Log - " & Now

' DWORD Value Names and Data for Chrome Settings
Dim ChromeValueNames(8), ChromeValueData(8)
ChromeValueNames(0) = "BrowserSignin"
ChromeValueData(0) = 0
ChromeValueNames(1) = "SyncDisabled"
ChromeValueData(1) = 1
ChromeValueNames(2) = "IncognitoModeAvailability"
ChromeValueData(2) = 1
ChromeValueNames(3) = "BrowserGuestModeEnabled"
ChromeValueData(3) = 0
ChromeValueNames(4) = "NewTabPageQuickLinksEnabled"
ChromeValueData(4) = 0
ChromeValueNames(5) = "NewTabPageContentEnabled"
ChromeValueData(5) = 0
ChromeValueNames(6) = "BrowserAddProfileEnabled"
ChromeValueData(6) = 0
ChromeValueNames(7) = "EnableMediaRouter"
ChromeValueData(7) = 0
ChromeValueNames(8) = "SearchSuggestEnabled"
ChromeValueData(8) = 0

' String Value Names and Data for URLAllowlist
Dim AllowlistValueNames(62), AllowlistValueData(62)
AllowlistValueNames(0) = "1"
AllowlistValueData(0) = "privacy-sandbox-dialog"
AllowlistValueNames(1) = "2"
AllowlistValueData(1) = "privacy-notice-sandbox-dialog"
AllowlistValueNames(2) = "3"
AllowlistValueData(2) = "privacy-notice-sandbox-app"
AllowlistValueNames(3) = "4"
AllowlistValueData(3) = "privacy-sandbox-app"
AllowlistValueNames(4) = "5"
AllowlistValueData(4) = "intro"
AllowlistValueNames(5) = "6"
AllowlistValueData(5) = "customize-chrome-side-panel.top-chrome"
AllowlistValueNames(6) = "7"
AllowlistValueData(6) = "password-manager"
AllowlistValueNames(7) = "8"
AllowlistValueData(7) = "settings"
AllowlistValueNames(8) = "9"
AllowlistValueData(8) = "downloads"
AllowlistValueNames(9) = "10"
AllowlistValueData(9) = "history "
AllowlistValueNames(10) = "11"
AllowlistValueData(10) = "bookmarks "
AllowlistValueNames(11) = "12"
AllowlistValueData(11) = "nurturing"
AllowlistValueNames(12) = "13"
AllowlistValueData(12) = "page-info"
AllowlistValueNames(13) = "14"
AllowlistValueData(13) = "permission-request-dialog"
AllowlistValueNames(14) = "15"
AllowlistValueData(14) = "collected-cookies-dialog"
AllowlistValueNames(15) = "16"
AllowlistValueData(15) = "about:blank"
AllowlistValueNames(16) = "17"
AllowlistValueData(16) = "127.0.0.1"
AllowlistValueNames(17) = "18"
AllowlistValueData(17) = "localhost"
AllowlistValueNames(18) = "19"
AllowlistValueData(18) = "print"
AllowlistValueNames(19) = "20"
AllowlistValueData(19) = "file:///C:/"
AllowlistValueNames(20) = "21"
AllowlistValueData(20) = "file:///D:/"
AllowlistValueNames(21) = "22"
AllowlistValueData(21) = "file:///E:/"
AllowlistValueNames(22) = "23"
AllowlistValueData(22) = "file:///F:/"
AllowlistValueNames(23) = "24"
AllowlistValueData(23) = "blob://*"
AllowlistValueNames(24) = "25"
AllowlistValueData(24) = "blob:*"
AllowlistValueNames(25) = "26"
AllowlistValueData(25) = "DevTools"
AllowlistValueNames(26) = "27"
AllowlistValueData(26) = "chrome-urls"
AllowlistValueNames(27) = "28"
AllowlistValueData(27) = "blob:internals"
AllowlistValueNames(28) = "29"
AllowlistValueData(28) = "version"
AllowlistValueNames(29) = "30"
AllowlistValueData(29) = "certificate-manager"
AllowlistValueNames(30) = "31"
AllowlistValueData(30) = "connection-monitoring-detected"
AllowlistValueNames(31) = "32"
AllowlistValueData(31) = "connection-help"
AllowlistValueNames(32) = "33"
AllowlistValueData(32) = "dino"
AllowlistValueNames(33) = "34"
AllowlistValueData(33) = "history-clusters-side-panel.top-chrome"
AllowlistValueNames(34) = "35"
AllowlistValueData(34) = "javascript"
AllowlistValueNames(35) = "36"
AllowlistValueData(35) = "whats-new"
AllowlistValueNames(36) = "37"
AllowlistValueData(36) = "view-source"
AllowlistValueNames(37) = "38"
AllowlistValueData(37) = "webrtc-internals"
AllowlistValueNames(38) = "39"
AllowlistValueData(38) = "view-cert"
AllowlistValueNames(39) = "40"
AllowlistValueData(39) = "tab-search.top-chrome"
AllowlistValueNames(40) = "41"
AllowlistValueData(40) = "system"
AllowlistValueNames(41) = "42"
AllowlistValueData(41) = "signin-email.confirmation"
AllowlistValueNames(42) = "43"
AllowlistValueData(42) = "googletagmanager"
AllowlistValueNames(43) = "44"
AllowlistValueData(43) = "reset-password"
AllowlistValueNames(44) = "45"
AllowlistValueData(44) = "net-internals"
AllowlistValueNames(45) = "46"
AllowlistValueData(45) = "net-export"
AllowlistValueNames(46) = "47"
AllowlistValueData(46) = "netlog-viewer.appspot.com"
AllowlistValueNames(47) = "48"
AllowlistValueData(47) = "management"
AllowlistValueNames(48) = "49"
AllowlistValueData(48) = "managed-user-profile-notice"
AllowlistValueNames(49) = "50"
AllowlistValueData(49) = "internals"
AllowlistValueNames(50) = "51"
AllowlistValueData(50) = "inspect"
AllowlistValueNames(51) = "52"
AllowlistValueData(51) = "history-side-panel.top-chrome"
AllowlistValueNames(52) = "53"
AllowlistValueData(52) = "https://googleapis.com/"
AllowlistValueNames(53) = "54"
AllowlistValueData(53) = "javascript:void(0)"
AllowlistValueNames(54) = "55"
AllowlistValueData(54) = "javascript:void0"
AllowlistValueNames(55) = "56"
AllowlistValueData(55) = "javascript:void()"
AllowlistValueNames(56) = "57"
AllowlistValueData(56) = "javascript:void"
AllowlistValueNames(57) = "58"
AllowlistValueData(57) = "javascript:"
AllowlistValueNames(58) = "59"
AllowlistValueData(58) = "googleusercontent.com"
AllowlistValueNames(59) = "301"
AllowlistValueData(59) = "mail.zoho.in"
AllowlistValueNames(60) = "302"
AllowlistValueData(60) = "accounts.zoho.in"
AllowlistValueNames(61) = "303"
AllowlistValueData(61) = "zmdownload-accl.zoho.in"
AllowlistValueNames(62) = "304"
AllowlistValueData(62) = "zmdownloadfree-accl.zoho.in"

' String Value Names and Data for URLBlocklist
Dim BlocklistValueNames(0), BlocklistValueData(0)
BlocklistValueNames(0) = "1"
BlocklistValueData(0) = "*"

' Function to write DWORD to the registry
Sub WriteRegistryDWORD(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath, ExistingValue
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & ValueName
    ExistingValue = ""
    On Error Resume Next
    ExistingValue = WSHShell.RegRead(RegKeyPath)
    If Err.Number = 0 Then
        If ExistingValue <> ValueData Then
            WSHShell.RegWrite RegKeyPath, ValueData, "REG_DWORD"
            LogFile.WriteLine "Overwrote different DWORD value: " & ValueName & " = " & ValueData & " (was: " & ExistingValue & ")"
        Else
            LogFile.WriteLine "DWORD value unchanged: " & ValueName & " = " & ValueData
        End If
    Else
        WSHShell.RegWrite RegKeyPath, ValueData, "REG_DWORD"
        LogFile.WriteLine "Wrote new DWORD value: " & ValueName & " = " & ValueData
    End If
    If Err.Number <> 0 Then
        WScript.Echo "Error writing DWORD to registry: " & Err.Description, 0, "Error", 16
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Function to write string to the registry
Sub WriteRegistryString(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath, ExistingValue
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & ValueName
    ExistingValue = ""
    On Error Resume Next
    ExistingValue = WSHShell.RegRead(RegKeyPath)
    If Err.Number = 0 Then
        If ExistingValue <> ValueData Then
            If ValueData = "*" And ExistingValue <> "*" Then
                WSHShell.RegWrite RegKeyPath, "*", "REG_SZ"
                LogFile.WriteLine "Overwrote different string value with *: " & ValueName & " = * (was: " & ExistingValue & ")"
            ElseIf ValueData <> "*" And ExistingValue = "*" Then
                LogFile.WriteLine "String value unchanged: " & ValueName & " = *"
            Else
                WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
                LogFile.WriteLine "Overwrote different string value: " & ValueName & " = " & ValueData & " (was: " & ExistingValue & ")"
            End If
        Else
            LogFile.WriteLine "String value unchanged: " & ValueName & " = " & ValueData
        End If
    Else
        WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
        LogFile.WriteLine "Wrote new string value: " & ValueName & " = " & ValueData
    End If
    If Err.Number <> 0 Then
        WScript.Echo "Error writing string to registry: " & Err.Description, 0, "Error", 16
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim i, Success
Success = True

' Write Chrome settings
For i = 0 To UBound(ChromeValueNames)
    WriteRegistryDWORD RegPathChrome, ChromeValueNames(i), ChromeValueData(i)
    If Err.Number <> 0 Then Success = False
Next

' Write URLAllowlist
For i = 0 To UBound(AllowlistValueNames)
    WriteRegistryString RegPathURLAllowlist, AllowlistValueNames(i), AllowlistValueData(i)
    If Err.Number <> 0 Then Success = False
Next

' Write URLBlocklist
For i = 0 To UBound(BlocklistValueNames)
    WriteRegistryString RegPathURLBlocklist, BlocklistValueNames(i), BlocklistValueData(i)
    If Err.Number <> 0 Then Success = False
Next

' Display result
If Success Then
    WScript.Echo "GC_BASE_URLA_ADDED" & vbCrLf & _
                 "@Cyber Informatics 2025" & vbCrLf & _
                 "(@Ci Enterprises 2021)", 0, "Success", 64
Else
    WScript.Echo "An error occurred while applying Chrome policies. Check log at: " & LogPath, 0, "Error", 16
End If

LogFile.Close
WScript.Quit

' Function to check if running as administrator
Function IsAdmin
    Dim WSHShell, TestKey, Result
    Set WSHShell = CreateObject("WScript.Shell")
    TestKey = "HKLM\Software\"
    On Error Resume Next
    WSHShell.RegRead TestKey
    Result = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    Set WSHShell = Nothing
    IsAdmin = Result
End Function
