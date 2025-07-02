Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Google\Chrome\URLAllowlist"

' String Value Names and Data
Dim ValueNames(4), ValueData(4)
ValueNames(0) = "6300"
ValueData(0) = "f4email.rediff.com"

ValueNames(1) = "6301"
ValueData(1) = "w3.org"

ValueNames(2) = "6302"
ValueData(2) = "in.rediff.com"

ValueNames(3) = "6303"
ValueData(3) = "mail.rediff.com"

' Function to write to the registry
Sub WriteRegistryValue(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & "\" & ValueName
    WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
    If Err.Number <> 0 Then
        WScript.Echo "Error on permitting login and logout to zmail: " & Err.Description
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim i, success
success = True
For i = 0 to UBound(ValueNames)
    WriteRegistryValue RegPath, ValueNames(i), ValueData(i)
    If Err.Number <> 0 Then
             success = False
    End If
Next
If success Then
WScript.Echo "Rediffmail accounts login & logout permitted" & vbCrLf & _
             "@ Ci ENTERPRISES" & vbCrLf & _
             "(@ CiEVERSE by Ci ENTERPRISES)"
End If