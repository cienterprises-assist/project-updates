Option Explicit

Dim WShell, FSO, LogFile, TempFolder, LogPath, Key, Keys, KeyExists
Dim ChromeStatus, EdgeStatus, PopupMessage
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Check for administrative privileges
If Not IsAdmin Then
    WShell.Popup "This script must be run as Administrator!", 0, "Error", 16
    WScript.Quit 1
End If

' Set up logging
TempFolder = WShell.ExpandEnvironmentStrings("%TEMP%")
LogPath = TempFolder & "\DeleteBrowserPolicies.log"
On Error Resume Next
Set LogFile = FSO.CreateTextFile(LogPath, True)
If Err.Number <> 0 Then
    WShell.Popup "Error creating log file: " & Err.Description, 0, "Error", 16
    WScript.Quit 1
End If
On Error GoTo 0

LogFile.WriteLine "Delete Browser Policies Log - " & Now

' List of registry keys to delete (only URLAllowlist)
Keys = Array( _
    "HKLM\Software\Policies\Microsoft\Edge\URLAllowlist\", _
    "HKLM\Software\Policies\Google\Chrome\URLAllowlist\" _
)

' Initialize status flags
ChromeStatus = "NotFound" ' Possible values: "NotFound", "Deleted"
EdgeStatus = "NotFound"   ' Possible values: "NotFound", "Deleted"

' Process each key
Dim i
For i = 0 To UBound(Keys)
    Key = Keys(i)
    On Error Resume Next
    ' Check if key exists
    KeyExists = False
    WShell.RegRead Key
    If Err.Number = 0 Then
        KeyExists = True
    Else
        LogFile.WriteLine "Key not found: " & Key
    End If
    Err.Clear

    ' Attempt to delete key if it exists
    If KeyExists Then
        DeleteRegistryKey Key
        If Err.Number <> 0 Then
            LogFile.WriteLine "Error deleting " & Key & ": " & Err.Description
            WShell.Popup "Error deleting " & Key & ": " & Err.Description, 0, "Error", 16
            LogFile.Close
            WScript.Quit 1
        Else
            LogFile.WriteLine "Successfully deleted: " & Key
            If InStr(Key, "Microsoft\Edge") > 0 Then
                EdgeStatus = "Deleted"
            ElseIf InStr(Key, "Google\Chrome") > 0 Then
                ChromeStatus = "Deleted"
            End If
        End If
    Else
        If InStr(Key, "Microsoft\Edge") > 0 Then
            EdgeStatus = "NotFound"
        ElseIf InStr(Key, "Google\Chrome") > 0 Then
            ChromeStatus = "NotFound"
        End If
    End If
    On Error GoTo 0
Next

LogFile.Close

' Determine popup message based on status
PopupMessage = ""
If ChromeStatus = "Deleted" And EdgeStatus = "Deleted" Then
    PopupMessage = "Your script run successfully" & vbCrLf & _
                   "@Cyber Informatics 2025" & vbCrLf & _
                   "(@Ci Enterprises 2021)"
Else
    If ChromeStatus = "NotFound" Then
        PopupMessage = PopupMessage & "ch URLA Key not found"
    ElseIf ChromeStatus = "Deleted" Then
        PopupMessage = PopupMessage & "Your script run successfully for Chrome"
    End If
    If EdgeStatus = "NotFound" Then
        If PopupMessage <> "" Then PopupMessage = PopupMessage & vbCrLf
        PopupMessage = PopupMessage & "mse URLA Key not found"
    ElseIf EdgeStatus = "Deleted" Then
        If PopupMessage <> "" Then PopupMessage = PopupMessage & vbCrLf
        PopupMessage = PopupMessage & "Your script run successfully for Edge"
    End If
End If

' Display popup
WShell.Popup PopupMessage, 0, "Result", 64

WScript.Quit

' Function to check if running as administrator
Function IsAdmin
    Dim TestKey, Result
    TestKey = "HKLM\Software\"
    On Error Resume Next
    WShell.RegRead TestKey
    Result = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    IsAdmin = Result
End Function

' Function to recursively delete registry key and subkeys
Sub DeleteRegistryKey(KeyPath)
    On Error Resume Next
    Dim SubKey, i
    ' Attempt to delete numeric subkeys (e.g., "1", "2", etc.)
    For i = 1 To 1000 ' Arbitrary limit for subkeys
        SubKey = KeyPath & CStr(i)
        WShell.RegRead SubKey
        If Err.Number = 0 Then
            WShell.RegDelete SubKey
        Else
            Err.Clear
            Exit For
        End If
    Next
    
    ' Delete the main key
    WShell.RegDelete KeyPath
    If Err.Number <> 0 Then
        Err.Clear ' Pass error to caller
    End If
    On Error GoTo 0
End Sub