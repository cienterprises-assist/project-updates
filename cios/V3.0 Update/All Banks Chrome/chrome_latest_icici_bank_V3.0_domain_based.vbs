Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Microsoft\Edge\URLAllowlist"

' String Value Names and Data
Dim ValueNames(9), ValueData(9)
ValueNames(0) = "501"
ValueData(0) = "neocaps.icicibank.com"
ValueNames(1) = "502"
ValueData(1) = "neocollections.icicibank.com"
ValueNames(2) = "503"
ValueData(2) = "omnidocs.icicibank.com"
ValueNames(3) = "601"
ValueData(3) = "dsmg-optimus.com"
ValueNames(4) = "701"
ValueData(4) = "ibox-vsts.icicibank.com"
ValueNames(5) = "801"
ValueData(5) = "icicibankltd.lms.getvymo.com"
ValueNames(6) = "901"
ValueData(6) = "filesecure.icicibank.com"
ValueNames(7) = "1001"
ValueData(7) = "emsigner.com"
ValueNames(8) = "1002"
ValueData(8) = "in.emsigner.com"


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

WScript.Echo "Your allowed websites are added" & vbCrLf & _
             "@Cyber Informatics 2025" & vbCrLf & _
             "(@Ci Enterprises 2021)"