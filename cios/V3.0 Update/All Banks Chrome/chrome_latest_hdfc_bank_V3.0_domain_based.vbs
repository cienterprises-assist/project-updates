Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Google\Chrome\URLAllowlist\"

' Check for administrative privileges
If Not IsAdmin Then
    WScript.Echo "This script must be run as Administrator! Please run as Administrator.", 0, "Error", 16
    WScript.Quit 1
End If

' Set up logging
Dim FSO, LogFile, TempFolder, LogPath
Set FSO = CreateObject("Scripting.FileSystemObject")
TempFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
LogPath = TempFolder & "\ChromeHDFCUpdate.log"
On Error Resume Next
Set LogFile = FSO.CreateTextFile(LogPath, True)
If Err.Number <> 0 Then
    WScript.Echo "Error creating log file: " & Err.Description, 0, "Error", 16
    WScript.Quit 1
End If
On Error GoTo 0
LogFile.WriteLine "HDFC Bank URLAllowlist Update Log - " & Now & " (IST)"

' String Value Names and Data
Dim ValueNames(38), ValueData(38)
ValueNames(0) = "501"
ValueData(0) = "appaccesss.hdfcbank.com"
ValueNames(1) = "502"
ValueData(1) = "collections.hdfcbank.com"
ValueNames(2) = "701"
ValueData(2) = "hdfcbwc.lms.getvymo.com"
ValueNames(3) = "702"
ValueData(3) = "hdfc.lms.getvymo.com"
ValueNames(4) = "703"
ValueData(4) = "recordings.mum1.exotel.com"
ValueNames(5) = "704"
ValueData(5) = "exotel.com"
ValueNames(6) = "901"
ValueData(6) = "drm.hdfcbank.com"
ValueNames(7) = "1001"
ValueData(7) = "emsigner.com"
ValueNames(8) = "1002"
ValueData(8) = "in.emsigner.com"
ValueNames(9) = "1101"
ValueData(9) = "collxnidms.com"
ValueNames(10) = "1301"
ValueData(10) = "cdn.forms.office.net"
ValueNames(11) = "1302"
ValueData(11) = "forms.office.net"
ValueNames(12) = "1303"
ValueData(12) = "office.net"
ValueNames(13) = "1304"
ValueData(13) = "cdn.forms.office.com"
ValueNames(14) = "1305"
ValueData(14) = "forms.office.com"
ValueNames(15) = "1306"
ValueData(15) = "office.com"
ValueNames(16) = "1307"
ValueData(16) = "cdn.forms.office.in"
ValueNames(17) = "1308"
ValueData(17) = "forms.office.in"
ValueNames(18) = "1309"
ValueData(18) = "office.in"
ValueNames(19) = "1501"
ValueData(19) = "teams.microsoft.com"
ValueNames(20) = "1701"
ValueData(20) = "accounts.google.com"
ValueNames(21) = "1702"
ValueData(21) = "docs.google.com"
ValueNames(22) = "1703"
ValueData(22) = "docst.google.com"
ValueNames(23) = "1901"
ValueData(23) = "cartradeexchange.com"
ValueNames(24) = "1902"
ValueData(24) = "hdfc.cartradeexchange.com"
ValueNames(25) = "1903"
ValueData(25) = "crtra.de"
ValueNames(26) = "2101"
ValueData(26) = "osend.in"
ValueNames(27) = "2102"
ValueData(27) = "tracker.osend.in"
ValueNames(28) = "2103"
ValueData(28) = "trackerc.osend.in"
ValueNames(29) = "2301"
ValueData(29) = "phonon.io"
ValueNames(30) = "2302"
ValueData(30) = "chatcalling.phonon.io"
ValueNames(31) = "2501"
ValueData(31) = "www.hdfcbank.com"
ValueNames(32) = "2502"
ValueData(32) = "hdfcbank.com"
ValueNames(33) = "2701"
ValueData(33) = "signdesk.in"
ValueNames(34) = "2702"
ValueData(34) = "api.signdesk.in"
ValueNames(35) = "2703"
ValueData(35) = "in.signdesk.in"
ValueNames(36) = "2704"
ValueData(36) = "signdesk.com"
ValueNames(37) = "2705"
ValueData(37) = "api.signdesk.com"
ValueNames(38) = "2706"
ValueData(38) = "in.signdesk.com"

' Function to write to the registry
Sub WriteRegistryValue(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath, ExistingValue
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & ValueName
    ExistingValue = ""
    On Error Resume Next
    ExistingValue = WSHShell.RegRead(RegKeyPath)
    If Err.Number = 0 Then
        If ExistingValue = ValueData Then
            WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
            LogFile.WriteLine "Overwrote existing string value: " & ValueName & " = " & ValueData
        Else
            WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
            LogFile.WriteLine "Overwrote different string value: " & ValueName & " = " & ValueData & " (was: " & ExistingValue & ")"
        End If
    Else
        WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
        LogFile.WriteLine "Wrote new string value: " & ValueName & " = " & ValueData
    End If
    If Err.Number <> 0 Then
        LogFile.WriteLine "Error writing to registry for " & ValueName & ": " & Err.Description & " (Error Code: " & Err.Number & ")"
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim i, Success
Success = True

' Write URLAllowlist values
For i = 0 To UBound(ValueNames)
    WriteRegistryValue RegPath, ValueNames(i), ValueData(i)
    If Err.Number <> 0 Then Success = False
Next

' Display result
If Success Then
    WScript.Echo "Your allowed websites are added to chrome enterprise" & vbCrLf & _
                 "@Cyber Informatics 2025" & vbCrLf & _
                 "(@Ci Enterprises 2021)", 0, "Success", 64
Else
    WScript.Echo "An error occurred while applying HDFC Bank URLAllowlist policies. Check log at: " & LogPath, 0, "Error", 16
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