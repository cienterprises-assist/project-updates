Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Google\Chrome\URLAllowlist"

' String Value Names and Data
Dim ValueNames(0), ValueData(0)
ValueNames(0) = "302"
ValueData(0) = "accounts.zoho.in"

' Function to delete to the registry
Sub DeleteRegistryValue(Path, ValueName)
    On Error Resume Next
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.RegDelete Path & "\" & ValueName
    If Err.Number <> 0 Then
        WScript.Echo "Error restricting login and logout to zmail: " & Err.Description
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim i, success
success = True
For i = 0 to UBound(ValueNames)
    DeleteRegistryValue RegPath, ValueNames(i)
    If Err.Number <> 0 Then
             success = False
    End If
Next
If success Then
WScript.Echo "Zmail accounts login & logout restricted" & vbCrLf & _
             "@ Ci ENTERPRISES" & vbCrLf & _
             "(@ CiEVERSE by Ci ENTERPRISES)"
End If