Attribute VB_Name = "WV2Globals"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'WV2Globals
'-generally used to store state/collected data during browser operations
'make sure anything written to in here is cleaned up when the browser is closed,
'so the garbage collector can free up the memory used, and to prevent bugs between browsing sessions
'***

'store information as we browse
Public dicData As New Dictionary
Public dicRequests As New Dictionary
Public dicResponses As New Dictionary
Public dicContentHandlers As New Dictionary

'browser environment/control/display specific globals
Public g_wv2Env As wv2Environment
Public g_Env As ICoreWebView2Environment
Public g_webFrame As Control
Public g_webHostHwnd As Long
Public g_wv2() As wv2
Public g_selectedTabIndex As Long

Public Sub CleanUp()

Dim i As Integer

Debug.Print "cleaning up"

For i = 0 To UBound(g_wv2)
    Set g_wv2(i) = Nothing
Next i

Erase g_wv2
Set g_wv2Env = Nothing
Set g_Env = Nothing
Set g_webFrame = Nothing

'Unload WebView
'Unload WebviewMethodDc
'Set WebviewMethodDc = Nothing

If dicData.Count > 0 Then
    For i = 0 To dicData.Count - 1
'        Unload dicData(i)
        Set dicData(i) = Nothing
    Next i
End If
If dicRequests.Count > 0 Then
    For i = 0 To dicRequests.Count - 1
'        Unload dicRequests(i)
        Set dicRequests(i) = Nothing
    Next i
End If
If dicResponses.Count > 0 Then
    For i = 0 To dicResponses.Count - 1
'        Unload dicResponses(i)
        Set dicResponses(i) = Nothing
    Next i
End If
If dicContentHandlers.Count > 0 Then
    For i = 0 To dicContentHandlers.Count - 1
        Unload dicContentHandlers(i)
        Set dicContentHandlers(i) = Nothing
    Next i
End If
'Unload dicData
'Unload dicRequests
'Unload dicResponses
'Unload dicContentHandlers
Set dicData = Nothing
Set dicRequests = Nothing
Set dicResponses = Nothing
Set dicContentHandlers = Nothing

PluginManager.Kill


Debug.Print "cleaned up"
End Sub

'return the wv2 object of the currently selected tab
Public Property Get ActiveBrowserTab() As wv2
    Set ActiveBrowserTab = g_wv2(g_selectedTabIndex) 'g_selectedTabIndex is set by UserForm1.browserTabs_Change()
End Property


