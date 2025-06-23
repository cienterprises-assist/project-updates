Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Microsoft\Edge\URLAllowlist"

' String Value Names and Data
Dim ValueNames(3), ValueData(3)
ValueNames(0) = "10605"
ValueData(0) = "Test"


' Function to write to the registry
Sub WriteRegistryValue(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & "\" & ValueName
    WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
    If Err.Number <> 0 Then
        WScript.Echo "Error writing to registry: " & Err.Description
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim i
For i = 0 to UBound(ValueNames)
    WriteRegistryValue RegPath, ValueNames(i), ValueData(i)
Next

WScript.Echo "Cyber Update Successfully" & vbCrLf & _
             "@Cyber Informatics 2025" & vbCrLf & _
             "(@Ci Enterprises 2021)"