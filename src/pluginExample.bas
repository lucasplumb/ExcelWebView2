Attribute VB_Name = "pluginExample"
Option Explicit
Private m_scriptComplete As Boolean
Private m_lastExecutedScriptResponse As String
Private m_lastWebContentResponse As String
Private m_responseReceived As Boolean
Public Property Let LastWebContentResponse(val As String)
    m_lastWebContentResponse = val
End Property
Public Property Let LastExecutedScriptResponse(val As String)
    m_lastExecutedScriptResponse = val
End Property
Public Property Get ScriptComplete() As Boolean
    ScriptComplete = m_scriptComplete
End Property
Public Property Let ScriptComplete(val As Boolean)
    m_scriptComplete = val
End Property
Public Property Get ResponseReceived() As Boolean
    ResponseReceived = m_responseReceived
End Property
Public Property Let ResponseReceived(val As Boolean)
    m_responseReceived = val
End Property

Public Sub Search1()
    Debug.Print "---SEARCH EXAMPLE #1---"
    ActiveBrowserTab.OpenUrl "en.wikipedia.org"
    'the next line will fail, because the page navigation hasnt completed yet
    ActiveBrowserTab.ExecuteScript "document.getElementById('searchInput').value = 'WebView2'"
    Debug.Print "observe that the search input value was not set, because we tried to set it before the page was done loading..."
    Debug.Print ""
    Stop
End Sub

Public Sub Search2()
    Debug.Print "---SEARCH EXAMPLE #2---"
    ActiveBrowserTab.OpenUrl "en.wikipedia.org"
    'since the browser is event driven, we need to wait for navigation to complete
    WaitForNavigation
    ActiveBrowserTab.ExecuteScript "document.getElementById('searchInput').value = 'WebView2'"
    ActiveBrowserTab.ExecuteScript "document.getElementById('search-form').submit();"
    WaitForNavigation
    ActiveBrowserTab.ExecuteScript "document.documentElement.outerHTML;"
    ScriptComplete = False
    Debug.Print "source - " & m_lastExecutedScriptResponse 'fails because the script hasnt completed yet - see pluginExampleCls - m_lastExecutedScriptResponse is set when the script event handler completes
    'Stop 'pause to observe that the page HTML has not printed
    Debug.Print "observe that the source HTML is not printed correctly - that is because we need to wait for the script to complete"
    'lets try waiting for it to finish now
    WaitForScriptComplete
    Debug.Print "source - " & Left$(m_lastExecutedScriptResponse, 1000) & "..."
    Debug.Print "as you can see, the source has now been printed"
    Debug.Print ""
    Stop
End Sub

Public Sub Search3()
    Debug.Print "---SEARCH EXAMPLE #3---"
    'instead of navigating, lets send a POST (GET, in this case) request
    Dim source As String
    Dim webReq As clsWebResourceBuilder
    Set webReq = New clsWebResourceBuilder
    With webReq
        .Method = HTTP_GET
        .uri = "https://en.wikipedia.org/wiki/Special:Search?search=WebView2&go=Go"
        .SetRequestHeader "Referer", "https://www.wikipedia.org"
        .PostRequest ActiveBrowserTab
    End With
    WaitForNavigation
    ActiveBrowserTab.ExecuteScript "document.documentElement.innerHTML;"
    ScriptComplete = False
    Debug.Print "source - " & m_lastExecutedScriptResponse 'fails because the script hasnt completed yet - see pluginExampleCls - m_lastExecutedScriptResponse is set when the script event handler completes
    'Stop 'pause to observe that the page HTML has not printed
    Debug.Print "observe that the source HTML is not printed correctly - that is because we need to wait for the script to complete"
    'lets try waiting for it to finish now
    WaitForScriptComplete
    Debug.Print "source - " & Left$(m_lastExecutedScriptResponse, 1000) & "..."
    Debug.Print "as you can see, the source has now been printed"
    Debug.Print ""
    Stop
End Sub

Public Sub Search4()
    Debug.Print "---SEARCH EXAMPLE #4---"
    'instead of navigating, lets send a POST (GET, in this case) request
    Dim source As String
    Dim webReq As clsWebResourceBuilder
    Set webReq = New clsWebResourceBuilder
    With webReq
        .Method = HTTP_GET
        .uri = "https://en.wikipedia.org/wiki/Special:Search?search=WebView2&go=Go"
        .SetRequestHeader "Referer", "https://www.wikipedia.org"
        .PostRequest ActiveBrowserTab
    End With
    WaitForResponse 'instead of waiting for navigation, wait for a particular response... see 'm_ContentEvent_WebResourceResponseViewGetContentCompleted' in the pluginExampleCls
    'instead of getting the html via JS, lets get the web server's response content this time
    Debug.Print "web response - " & Left$(m_lastWebContentResponse, 1000) & "..."
    Debug.Print "as you can see, the source has now been printed - and its actually formatted a little more nicely for us"
End Sub
Private Sub WaitForNavigation()
    Do Until ActiveBrowserTab.NavigationComplete
        DoEvents
    Loop
End Sub
Private Sub WaitForResponse()
    Do Until ResponseReceived
        DoEvents
    Loop
End Sub
Private Sub WaitForScriptComplete()
    Do Until ScriptComplete
        DoEvents
    Loop
End Sub
