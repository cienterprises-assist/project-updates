Option Explicit

' Registry Paths
Const RegPathEdge = "HKLM\SOFTWARE\Policies\Microsoft\Edge\"
Const RegPathURLAllowlist = "HKLM\SOFTWARE\Policies\Microsoft\Edge\URLAllowlist"
Const RegPathURLBlocklist = "HKLM\SOFTWARE\Policies\Microsoft\Edge\URLBlocklist"

' DWORD Value Names and Data for Edge Settings
Dim EdgeValueNames(15), EdgeValueData(15)
EdgeValueNames(0) = "InPrivateModeAvailability"
EdgeValueData(0) = 1
EdgeValueNames(1) = "EnableMediaRouter"
EdgeValueData(1) = 0
EdgeValueNames(2) = "FamilySafetySettinsEnabled"
EdgeValueData(2) = 0
EdgeValueNames(3) = "ImportBrowserSettings"
EdgeValueData(3) = 0
EdgeValueNames(4) = "BrowserAddProfileEnabled"
EdgeValueData(4) = 0
EdgeValueNames(5) = "BrowserGuestModeEnabled"
EdgeValueData(5) = 0
EdgeValueNames(6) = "BrowserSignin"
EdgeValueData(6) = 0
EdgeValueNames(7) = "BrowserKidsModeEnabled"
EdgeValueData(7) = 0
EdgeValueNames(8) = "CopilotPageContext"
EdgeValueData(8) = 0
EdgeValueNames(9) = "NewTabPageContentEnabled"
EdgeValueData(9) = 0
EdgeValueNames(10) = "ShowMicrosoftRewards"
EdgeValueData(10) = 0
EdgeValueNames(11) = "NewTabPageQuickLinksEnabled"
EdgeValueData(11) = 0
EdgeValueNames(12) = "SearchSuggestEnabled"
EdgeValueData(12) = 0
EdgeValueNames(13) = "EdgeShoppingAssistantEnabled"
EdgeValueData(13) = 0
EdgeValueNames(14) = "SpotlightExperiencesAndRecommendationsEnabled"
EdgeValueData(14) = 0
EdgeValueNames(15) = "SyncDisabled"
EdgeValueData(15) = 1

' String Value Names and Data for URLAllowlist
Dim AllowlistValueNames(47), AllowlistValueData(47)
AllowlistValueNames(0) = "1"
AllowlistValueData(0) = "nurturing"
AllowlistValueNames(1) = "2"
AllowlistValueData(1) = "page-info"
AllowlistValueNames(2) = "3"
AllowlistValueData(2) = "permission-request-dialog"
AllowlistValueNames(3) = "4"
AllowlistValueData(3) = "collected-cookies-dialog"
AllowlistValueNames(4) = "5"
AllowlistValueData(4) = "wallet"
AllowlistValueNames(5) = "6"
AllowlistValueData(5) = "net-internals"
AllowlistValueNames(6) = "7"
AllowlistValueData(6) = "netlog-viewer.appspot.com"
AllowlistValueNames(7) = "8"
AllowlistValueData(7) = "127.0.0.1"
AllowlistValueNames(8) = "9"
AllowlistValueData(8) = "localhost"
AllowlistValueNames(9) = "10"
AllowlistValueData(9) = "favorites"
AllowlistValueNames(10) = "11"
AllowlistValueData(10) = "downloads"
AllowlistValueNames(11) = "12"
AllowlistValueData(11) = "history"
AllowlistValueNames(12) = "13"
AllowlistValueData(12) = "settings"
AllowlistValueNames(13) = "14"
AllowlistValueData(13) = "print"
AllowlistValueNames(14) = "15"
AllowlistValueData(14) = "file:///C:/"
AllowlistValueNames(15) = "16"
AllowlistValueData(15) = "file:///D:/"
AllowlistValueNames(16) = "17"
AllowlistValueData(16) = "file:///E:/"
AllowlistValueNames(17) = "18"
AllowlistValueData(17) = "file:///F:/"
AllowlistValueNames(18) = "19"
AllowlistValueData(18) = "about:blank"
AllowlistValueNames(19) = "20"
AllowlistValueData(19) = "net-export"
AllowlistValueNames(20) = "21"
AllowlistValueData(20) = "collections"
AllowlistValueNames(21) = "22"
AllowlistValueData(21) = "devtools"
AllowlistValueNames(22) = "23"
AllowlistValueData(22) = "blob://*"
AllowlistValueNames(23) = "24"
AllowlistValueData(23) = "blob:*"
AllowlistValueNames(24) = "25"
AllowlistValueData(24) = "performance-center"
AllowlistValueNames(25) = "26"
AllowlistValueData(25) = "version"
AllowlistValueNames(26) = "27"
AllowlistValueData(26) = "flags"
AllowlistValueNames(27) = "28"
AllowlistValueData(27) = "edge-urls"
AllowlistValueNames(28) = "29"
AllowlistValueData(28) = "about"
AllowlistValueNames(29) = "30"
AllowlistValueData(29) = "apps"
AllowlistValueNames(30) = "31"
AllowlistValueData(30) = "blob-internals"
AllowlistValueNames(31) = "32"
AllowlistValueData(31) = "crashes"
AllowlistValueNames(32) = "33"
AllowlistValueData(32) = "help"
AllowlistValueNames(33) = "34"
AllowlistValueData(33) = "network-errors"
AllowlistValueNames(34) = "35"
AllowlistValueData(34) = "pre-launch-fre"
AllowlistValueNames(35) = "36"
AllowlistValueData(35) = "webrtc-internals"
AllowlistValueNames(36) = "37"
AllowlistValueData(36) = "webrtc-logs"
AllowlistValueNames(37) = "501"
AllowlistValueData(37) = "cieverse.com"
AllowlistValueNames(38) = "502"
AllowlistValueData(38) = "cienterprises.org.in"
AllowlistValueNames(39) = "1001"
AllowlistValueData(39) = "assist.zoho.in"
AllowlistValueNames(40) = "1002"
AllowlistValueData(40) = "join.zoho.in"
AllowlistValueNames(41) = "1101"
AllowlistValueData(41) = "github.io"
AllowlistValueNames(42) = "1102"
AllowlistValueData(42) = "githubusercontent.com"
AllowlistValueNames(43) = "701"
AllowlistValueData(43) = "cienterprises-my.sharepoint.com"
AllowlistValueNames(44) = "702"
AllowlistValueData(44) = "cienterprisesorg-my.sharepoint.com"
AllowlistValueNames(45) = "1502"
AllowlistValueData(45) = "v.gd"
AllowlistValueNames(46) = "1503"
AllowlistValueData(46) = "is.gd"
AllowlistValueNames(47) = "1504"
AllowlistValueData(47) = "tinyurl.com"
AllowlistValueNames(48) = "1505"
AllowlistValueData(48) = "shorturl.at"
AllowlistValueNames(49) = "38"
AllowlistValueData(49) = "googleapis.com"
AllowlistValueNames(50) = "39"
AllowlistValueData(50) = "msftstatic.com"
AllowlistValueNames(51) = "40"
AllowlistValueData(51) = "gstatic.com"
AllowlistValueNames(52) = "41"
AllowlistValueData(52) = "javascript:void(0)"
AllowlistValueNames(53) = "42"
AllowlistValueData(53) = "javascript:void0"
AllowlistValueNames(54) = "43"
AllowlistValueData(54) = "javascript:void()"
AllowlistValueNames(55) = "44"
AllowlistValueData(55) = "javascript:void"
AllowlistValueNames(56) = "45"
AllowlistValueData(56) = "javascript:"
AllowlistValueNames(57) = "46"
AllowlistValueData(57) = "javascript"
AllowlistValueNames(58) = "47"
AllowlistValueData(58) = "googleusercontent.com"
AllowlistValueNames(59) = "48"
AllowlistValueData(59) = "k7computing.com"

' String Value Names and Data for URLBlocklist
Dim BlocklistValueNames(0), BlocklistValueData(0)
BlocklistValueNames(0) = "1"
BlocklistValueData(0) = "*"

' Function to write DWORD to the registry
Sub WriteRegistryDWORD(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & ValueName
    WSHShell.RegWrite RegKeyPath, ValueData, "REG_DWORD"
    If Err.Number <> 0 Then
        WScript.Echo "Error writing DWORD to registry: " & Err.Description
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Function to write string to the registry
Sub WriteRegistryString(Path, ValueName, ValueData)
    On Error Resume Next
    Dim WSHShell, RegKeyPath
    Set WSHShell = CreateObject("WScript.Shell")
    RegKeyPath = Path & ValueName
    WSHShell.RegWrite RegKeyPath, ValueData, "REG_SZ"
    If Err.Number <> 0 Then
        WScript.Echo "Error writing string to registry: " & Err.Description
    End If
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

' Main script execution
Dim i

' Write Edge settings
For i = 0 to UBound(EdgeValueNames)
    WriteRegistryDWORD RegPathEdge, EdgeValueNames(i), EdgeValueData(i)
Next

' Write URLAllowlist
For i = 0 to UBound(AllowlistValueNames)
    WriteRegistryString RegPathURLAllowlist, AllowlistValueNames(i), AllowlistValueData(i)
Next

' Write URLBlocklist
For i = 0 to UBound(BlocklistValueNames)
    WriteRegistryString RegPathURLBlocklist, BlocklistValueNames(i), BlocklistValueData(i)
Next

WScript.Echo "Your Base of Edge is added" & vbCrLf & _
             "@Cyber Informatics 2025" & vbCrLf & _
             "(@Ci Enterprises 2021)"