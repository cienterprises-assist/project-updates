Option Explicit

' Define constants
Const HKEY_USERS = &H80000003
Const REG_SZ = 1

' User details
Dim strUsername, strFullName, strSID
strUsername = "User"
strFullName = "System Operator (SysOp)"
strSID = "S-1-5-21-2296551787-2341494431-3366209023-1002"

' Create Shell and Registry objects
Dim objShell, objRegistry
Set objShell = CreateObject("WScript.Shell")
Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' Check if SID is loaded in HKEY_USERS
Dim arrSubKeys, bSIDLoaded
objRegistry.EnumKey HKEY_USERS, "", arrSubKeys
bSIDLoaded = False
If Not IsNull(arrSubKeys) Then
    For Each subKey In arrSubKeys
        If InStr(subKey, strSID) > 0 Then
            bSIDLoaded = True
            Exit For
        End If
    Next
End If

' Load SID hive if not loaded
If Not bSIDLoaded Then
    Dim strHivePath
    strHivePath = "C:/Users/" & strUsername & "\NTUSER.DAT"
    objShell.Run "reg load HKU\" & strSID & " " & Chr(34) & strHivePath & Chr(34), 0, True
    If Err.Number <> 0 Then
        WScript.Echo "Failed to load SID " & strSID & " hive. Error: " & Err.Description, 0, "Error", 16
        WScript.Quit 1
    End If
End If

' Update full name in loaded hive
Dim strKeyPath
strKeyPath = "HKEY_USERS\" & strSID
objRegistry.SetStringValue HKEY_USERS, strSID, "FullName", strFullName
WScript.Echo "Updated full name for " & strUsername & " (SID: " & strSID & ") to " & strFullName & ".", 0, "Success", 64

' Add additional message
WScript.Echo "Your Full Name of user is add " & strFullName & ".", 0, "Info", 64

' Unload the hive
objShell.Run "reg unload HKU\" & strSID, 0, True
If Err.Number <> 0 Then
    WScript.Echo "Failed to unload SID " & strSID & " hive. Error: " & Err.Description, 0, "Warning", 48
End If

' Cleanup
Set objRegistry = Nothing
Set objShell = Nothing
WScript.Quit
