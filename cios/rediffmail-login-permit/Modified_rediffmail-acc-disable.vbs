Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Google\Chrome\URLAllowlist"

' String Value Name and Data
Const ValueName = "6303"
Const ValueData = "mail.rediff.com"

' Function to delete a specific registry value
Sub DeleteRegistryValue(Path, ValName)
    On Error Resume Next
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    ' Check if the value exists before attempting deletion
    Dim currentValue
    currentValue = WSHShell.RegRead(Path & "\" & ValName)
    
    ' Only delete if the value exists and matches the expected data
    If Err.Number = 0 And currentValue = ValueData Then
        WSHShell.RegDelete Path & "\" & ValName
        If Err.Number <> 0 Then
            WScript.Echo "Error deleting registry value " & ValName & ": " & Err.Description
            Exit Sub
        End If
    ElseIf Err.Number <> 0 Then
        WScript.Echo "Registry value " & ValName & " does not exist or is inaccessible."
        Err.Clear
        Exit Sub
    Else
        WScript.Echo "Registry value " & ValName & " does not match expected data."
        Exit Sub
    End If
    
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim success
success = True

' Attempt to delete the specific registry value
DeleteRegistryValue RegPath, ValueName

' Check if any error occurred during deletion
If Err.Number <> 0 Then
    success = False
End If

' Display success message if deletion was successful
If success Then
    WScript.Echo "Rediffmail accounts login & logout restricted" & vbCrLf & _
                 "@ Ci ENTERPRISES" & vbCrLf & _
                 "(@ CiEVERSE by Ci ENTERPRISES)"
End If