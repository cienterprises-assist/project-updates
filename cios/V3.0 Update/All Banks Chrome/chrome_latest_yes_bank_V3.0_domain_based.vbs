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
LogPath = TempFolder & "\ChromeYesUpdate.log"
On Error Resume Next
Set LogFile = FSO.CreateTextFile(LogPath, True)
If Err.Number <> 0 Then
    WScript.Echo "Error creating log file: " & Err.Description, 0, "Error", 16
    WScript.Quit 1
End If
On Error GoTo 0
LogFile.WriteLine "Yes Bank URLAllowlist Update Log - " & Now & " (IST)"

' String Value Names and Data
Dim ValueNames(2), ValueData(2)
ValueNames(0) = "501"
ValueData(0) = "yescollect.yesbank.in"
ValueNames(1) = "601"
ValueData(1) = "retailloans.yes.in"
ValueNames(2) = "602"
ValueData(2) = "retailloans.yesbank.in"

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
    WScript.Echo "An error occurred while applying Yes Bank URLAllowlist policies. Check log at: " & LogPath, 0, "Error", 16
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