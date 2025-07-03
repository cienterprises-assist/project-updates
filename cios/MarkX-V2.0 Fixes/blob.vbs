Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Microsoft\Edge\URLAllowlist"

' String Value Names and Data
Dim ValueNames(4), ValueData(4)
ValueNames(0) = "15605"
ValueData(0) = "blob://*"
ValueNames(1) = "15606"
ValueData(1) = "blob:*"
ValueNames(2) = "15607"
ValueData(2) = "zmdownloadfree-accl.zoho.in"
ValueNames(3) = "15608"
ValueData(3) = "zmdownload-accl.zoho.in"


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

WScript.Echo "Your download issue has been solved" & vbCrLf & _
             "@Cyber Informatics 2025" & vbCrLf & _
             "(@Ci Enterprises 2021)"