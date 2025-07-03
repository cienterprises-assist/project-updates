Option Explicit

Dim WShell, FSO, LogFile, TempFolder, LogPath, Key, Keys, KeyExists, AllDeleted, ErrorCount
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Check for administrative privileges
If Not IsAdmin Then
    WShell.Popup "This script must be run as Administrator!", 0, "Error", 16
    WScript.Quit
End If

' Set up logging
TempFolder = WShell.ExpandEnvironmentStrings("%TEMP%")
LogPath = TempFolder & "\DeleteBrowserPolicies.log"
Set LogFile = FSO.CreateTextFile(LogPath, True)

LogFile.WriteLine "Delete Browser Policies Log - " & Now

' List of registry keys to delete (only URLAllowlist)
Keys = Array( _
    "HKLM\Software\Policies\Microsoft\Edge\URLAllowlist\", _
    "HKLM\Software\Policies\Google\Chrome\URLAllowlist\" _
)

AllDeleted = True
ErrorCount = 0

' Process each key
For Each Key In Keys
    On Error Resume Next
    ' Check if key exists
    KeyExists = False
    WShell.RegRead Key
    If Err.Number = 0 Then
        KeyExists = True
    Else
        LogFile.WriteLine "Key not found: " & Key
        AllDeleted = False
    End If
    Err.Clear

    ' Attempt to delete key if it exists
    If KeyExists Then
        WShell.RegDelete Key
        If Err.Number <> 0 Then
            LogFile.WriteLine "Error deleting " & Key & ": " & Err.Description
            AllDeleted = False
            ErrorCount = ErrorCount + 1
            Err.Clear
        Else
            LogFile.WriteLine "Successfully deleted: " & Key
        End If
    End If
    On Error Goto 0
Next

LogFile.Close

' Display result
If AllDeleted And ErrorCount = 0 Then
    WShell.Popup "All URLAllowlist registry keys have been successfully deleted.", 0, "Success", 64
Else
    WShell.Popup "Operation completed with issues. Check log at: " & LogPath, 0, "Warning", 48
End If

WScript.Quit

' Function to check if running as administrator
Function IsAdmin
    Dim TestKey, Result
    TestKey = "HKLM\Software\"
    On Error Resume Next
    WShell.RegRead TestKey
    Result = (Err.Number = 0)
    Err.Clear
    On Error Goto 0
    IsAdmin = Result
End Function