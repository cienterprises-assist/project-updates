Option Explicit

' Registry Path
Const RegPath = "HKLM\Software\Policies\Google\Chrome\URLAllowlist"

' String Value Names and Data
Dim ValueNames(13), ValueData(13)
ValueNames(0) = "201"
ValueData(0) = "javascript:void(0)"
ValueNames(1) = "202"
ValueData(1) = "javascript:void0"
ValueNames(2) = "203"
ValueData(2) = "https://pcccollections.axisbank.co.in/"
ValueNames(3) = "204"
ValueData(3) = "pcccollections.axisbank.co.in/"
ValueNames(4) = "205"
ValueData(4) = "https://login.microsoftonline.com/"
ValueNames(5) = "206"
ValueData(5) = "https://login.live.com/"
ValueNames(6) = "207"
ValueData(6) = "https://adfs.axisbank.com/"
ValueNames(7) = "208"
ValueData(7) = "https://adselfservice.axisbank.com/"
ValueNames(8) = "209"
ValueData(8) = "https://mysignins.microsoft.com/"
ValueNames(9) = "301"
ValueData(9) = "https://axisallocation.axisbank.co.in/"
ValueNames(10) = "302"
ValueData(10) = "axisallocation.axisbank.co.in/"
ValueNames(11) = "401"
ValueData(11) = "https://tap.tcsapps.com/"
ValueNames(12) = "402"
ValueData(12) = "tap.tcsapps.com/"

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