VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'wv2
'individual webview control object
'environment setup is performed automatically on first instantiation of this class
'simply spawn as many wv2's as needed for tabs etc
'***


Private m_WebViewController As ICoreWebView2Controller
Private m_WebViewCore As ICoreWebView2_2

Private WithEvents m_wv2env As wv2Environment
Attribute m_wv2env.VB_VarHelpID = -1
Private WithEvents m_webViewHandlers As clsWebViewEventHandlers
Attribute m_webViewHandlers.VB_VarHelpID = -1
Private WithEvents m_scriptHandler As clsWebViewScriptCompleteHandler
Attribute m_scriptHandler.VB_VarHelpID = -1

Private m_tab As MSForms.Tab
Attribute m_tab.VB_VarHelpID = -1

Private resDict As Dictionary

Private myIndex As Long

Private m_pageSource As String
Private m_navigationComplete As Boolean

'///////////////////////////////
'WEB CONTROL FUNCTIONS
'///////////////////////////////
Function NavigateWithResource(res As ICoreWebView2WebResourceRequest) 'can be used to POST data
    m_WebViewCore.NavigateWithWebResourceRequest res
    NavigationComplete = False
End Function
Function GoBack()
    m_WebViewCore.GoBack
End Function
Function GoForward()
    m_WebViewCore.GoForward
End Function
Function StopLoading()
    m_WebViewCore.Stop
End Function
Function Reload()
    m_WebViewCore.Reload
End Function
Function OpenUrl(ByVal url As String)
    If Left$(LCase$(url), 4) <> "http" And Left$(LCase$(url), 4) <> "file" Then url = "http://" & url
    m_WebViewCore.Navigate url
    NavigationComplete = False
End Function
Function OpenDevTools()
    m_WebViewCore.OpenDevToolsWindow
End Function
Function Focus()
    Dim i As Integer
    If (Not Not g_wv2) <> 0 Then
        'hide all controllers
        For i = LBound(g_wv2) To UBound(g_wv2)
            If Not g_wv2(i).controller Is Nothing Then
                g_wv2(i).controller.IsVisible = False
            End If
        Next i
        'set myself to visible
        g_wv2(myIndex).controller.IsVisible = True
    End If
End Function
Function ExecuteScript(javaScript As String, Optional PropLet As String = vbNullString)
    Dim scriptHandler As clsWebViewScriptCompleteHandler 'create an instance of the scriptCompleteHandler class and set its parent, which will then call wv2.ExecuteScriptCompletedHandler back to us in this wv2 instance
    Set scriptHandler = New clsWebViewScriptCompleteHandler
    Set scriptHandler.Parent = Me
    scriptHandler.PropLet = PropLet 'if we want the result of the script to set some variable when it completes, we can use this argument
    m_WebViewCore.ExecuteScript javaScript, scriptHandler 'm_scriptHandler
End Function
'//////////////////////////////




'///////////////////////////////
'PROPERTIES
'///////////////////////////////
Public Property Get index() As Long
    index = myIndex
End Property
Public Property Get controller() As ICoreWebView2Controller
    Set controller = m_WebViewController
End Property
Public Property Get handler() As clsWebViewEventHandlers
    If m_webViewHandlers Is Nothing Then
        Set m_webViewHandlers = New clsWebViewEventHandlers
    End If
    Set handler = m_webViewHandlers
End Property
Public Property Set handler(m_handler As clsWebViewEventHandlers)
    Set m_webViewHandlers = m_handler
End Property
Public Property Get browserTab() As MSForms.Tab
    Set browserTab = m_tab
End Property
Public Property Get resourceHistory(key As Long) As clsWebResData
    Set resourceHistory = resDict(key)
End Property
Public Property Get pageSource() As String
    ExecuteScript "document.documentElement.outerHTML;", "pageSource" 'use the PropLet argument to set the m_pageSource variable when the script completes
    pageSource = m_pageSource 'note this property will not be ready immediately, we need to wait for the script handler to return first
End Property
Public Property Let pageSource(val As String)
    m_pageSource = val
End Property
Public Property Get NavigationComplete() As Boolean
    NavigationComplete = m_navigationComplete
End Property
Public Property Let NavigationComplete(val As Boolean)
    m_navigationComplete = val
End Property
'////////////////////////////////









'///////////////////////////////
'CLASS CREATION FUNCTIONS
'///////////////////////////////
Private Sub WebViewReady() 'called each time a new wv2controller is ready (every new tab)
    Dim token As EventRegistrationToken 'just pass the same token pointer around, at this point we dont really care to ever remove these handlers, maybe <TODO> in the future
    
    'initialize web event handlers
    m_WebViewCore.AddWebResourceRequestedFilter "*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL 'you MUST add a WebResource filter to receive the events at all, we want to see all events so use the * wildcard
    m_WebViewCore.add_WebResourceRequested m_webViewHandlers, token
    m_WebViewCore.add_WebResourceResponseReceived m_webViewHandlers, token
    m_WebViewCore.add_DOMContentLoaded m_webViewHandlers, token
    m_WebViewCore.add_NavigationCompleted m_webViewHandlers, token
    m_WebViewCore.add_NavigationStarting m_webViewHandlers, token
    m_WebViewCore.add_DocumentTitleChanged m_webViewHandlers, token
    
    'navigate to homepage
    Me.OpenUrl homePageUrl
End Sub

Private Function GetWebView2Env(ByRef m_tab As MSForms.Tab) As wv2Environment
    Set m_tab = UserForm1.browserTabs.Tabs.Add("tab" & myIndex, "New Tab", myIndex)
    
    If g_wv2Env Is Nothing Then
        If g_webFrame Is Nothing Then
            Set g_webFrame = UserForm1.Controls.Add("Forms.Frame.1", "tab_Frame" & myIndex, False)
            With g_webFrame
                .Top = UserForm1.browserTabs.Top + 14
                .Left = UserForm1.browserTabs.Left
                .Height = UserForm1.browserTabs.Height - 14
                .Width = UserForm1.browserTabs.Width
                .TabStop = False
                .Visible = True
            End With
        End If
    
        g_webHostHwnd = g_webFrame.[_GethWnd]
    
        Set m_wv2env = New wv2Environment
        m_wv2env.Init
    End If
    
    Set m_wv2env = g_wv2Env
    Set GetWebView2Env = m_wv2env
    'Debug.Print myIndex
    'Set UserForm1.browserTabs.SelectedItem = UserForm1.browserTabs.Tabs.Item(myIndex + 1)
End Function


Private Sub Class_Initialize()
    Dim newCount As Integer
    Set resDict = New Dictionary
    
    'keep a reference to this instance in global
    If (Not Not g_wv2) = 0 Then 'if the g_wv2 array is uninitialized, this is the first instance
        
        'cleanup/initialize plugins on first instance creation
        PluginManager.Kill
        pluginLoader.LoadPlugins
        
        newCount = 0
    Else
        newCount = UBound(g_wv2) + 1
    End If
    myIndex = newCount
    
    If g_wv2Env Is Nothing Then
        'create webview2 environment if it doesnt exist yet
        Set m_wv2env = GetWebView2Env(m_tab)
        'controller will be created automatically when environment is initialized
        'see m_wv2env_wv2EnvReady
    Else
        Set m_wv2env = GetWebView2Env(m_tab)
        'environment already exists, so we need to create a new controller instead (this is called when this class is initialized more than once, for instance creating a new browser tab)
        g_wv2Env.this.CreateCoreWebView2Controller g_webHostHwnd, handler
    End If
    

    
    ReDim Preserve g_wv2(newCount)
    Set g_wv2(newCount) = Me
End Sub

Private Sub Class_Terminate()

    'Set handler = Nothing
    If Not m_WebViewCore Is Nothing Then
        m_WebViewCore.Stop
    End If
    Set m_webViewHandlers = Nothing
    Set m_WebViewCore = Nothing
    Set m_WebViewController = Nothing
    If Not resDict Is Nothing Then
        Dim v As Variant
        For Each v In resDict.Items
            Set v = Nothing
        Next v
        For Each v In resDict.Keys
            Set v = Nothing
        Next v
        resDict.RemoveAll
    End If
    Set resDict = Nothing
    
End Sub






















'///////////////////////////////
'EVENT HANDLERS
'///////////////////////////////
Private Sub m_scriptHandler_wv2ScriptComplete(ByVal sender As wv2, resultObjectAsJson As String)

End Sub
Public Sub ExecuteScriptCompletedHandler(errorCode As Long, resultObjectAsJson As String)

End Sub

Private Sub m_webViewHandlers_DocumentTitleChanged(sender As WebView2_edit.ICoreWebView2, args As Long)
    Dim docTitle As String
    docTitle = StrFromPtr(m_WebViewCore.DocumentTitle)
    If Len(docTitle) > 10 Then
        docTitle = Left(docTitle, 7) & "..."
    End If
    m_tab.Caption = docTitle
End Sub

Private Sub m_webViewHandlers_NavigationCompleted(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationCompletedEventArgs)
    Dim docTitle As String
    UserForm1.cmdStopReload.Caption = "R"
    docTitle = StrFromPtr(m_WebViewCore.DocumentTitle)
    If Len(docTitle) > 10 Then
        docTitle = Left(docTitle, 7) & "..."
    End If
    m_tab.Caption = docTitle
    NavigationComplete = True
End Sub

Private Sub m_webViewHandlers_NavigationStarting(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationStartingEventArgs)
    Dim uri As String
    UserForm1.cmdStopReload.Caption = "X"
    uri = StrFromPtr(args.uri)
    If Len(uri) > 10 Then
        uri = Left(uri, 7) & "..."
    End If
    m_tab.Caption = uri
    NavigationComplete = False
End Sub

Private Sub m_webViewHandlers_wv2ControllerReady(ByRef createdController As WebView2_edit.ICoreWebView2Controller)
    Dim wc2 As ICoreWebView2Controller2
    Set wc2 = createdController
    
    wc2.DefaultBackgroundColor = &H808080FF
    
    'sets the viewing area for the web view controller
    Dim RECT1 As RECT
    GetClientRect g_webHostHwnd, RECT1
    createdController.set_Bounds RECT1.Left, RECT1.Top, RECT1.Right, RECT1.Bottom
    
    Set m_WebViewController = createdController
    Set m_WebViewCore = createdController.CoreWebView2
    WebViewReady
End Sub

Private Sub m_webViewHandlers_wv2EnvironmentReady(ByRef createdEnvironment As WebView2_edit.ICoreWebView2Environment)

End Sub

Private Sub m_webViewHandlers_DOMContentLoaded(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2DOMContentLoadedEventArgs)
    '<TODO> figure out how to ensure this event is raised here prior to any plugins, for some reason plugin callbacks are going first
    'm_pageSource = pageSource 'when page loads, automatically request the source with javascript so that the pageSource property is populated when other code requests it
End Sub

Private Sub m_webViewHandlers_WebResourceRequested(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceRequestedEventArgs)

End Sub

Private Sub m_webViewHandlers_WebResourceResponseReceived(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceResponseReceivedEventArgs)
    Dim req As clsWebResData
    Set req = New clsWebResData
    
    Set req.sender = sender
    
    req.metaTitle = "RESPONSE"
    
    If Not args.Request Is Nothing Then
        'content/body/post data
        If Not args.Request.Content Is Nothing Then
            req.reqContent = IStreamToString(args.Request.Content)
        End If
        
        'uri/url
        req.uri = StrFromPtr(args.Request.uri)
        
        'GET/POST or whatever
        req.Method = StrFromPtr(args.Request.Method)
        
        
        
        'headers
        If Not args.Request.Headers Is Nothing Then
            If IterateHeaders(args.Request.Headers.GetIterator, req, "REQUEST") = False Then
                'error!
            End If
        End If
    End If

    If Not args.Response Is Nothing Then
    
        'reason
        req.ReasonPhrase = args.Response.ReasonPhrase
        
        'headers
        If Not args.Response.Headers Is Nothing Then
            If IterateHeaders(args.Response.Headers.GetIterator, req, "RESPONSE") = False Then
                'error!
            End If
        End If
        
        'status
        req.StatusCode = args.Response.StatusCode
        
        logResource req, resDict
        
        args.Response.GetContent req.contentHandler
    End If
    
    
End Sub

Private Sub m_wv2env_wv2CtrlReady(ByRef createdController As WebView2_edit.ICoreWebView2Controller)

End Sub

Private Sub m_wv2env_wv2EnvReady(ByRef createdEnvironment As WebView2_edit.ICoreWebView2Environment)
    g_wv2Env.this.CreateCoreWebView2Controller g_webHostHwnd, handler
End Sub

Private Sub m_wv2env_wv2Ready(env As wv2Environment)

End Sub





'///////////////////////////////
'HELPER FUNCTIONS
'///////////////////////////////

Private Function IterateHeaders(ByRef iterator As ICoreWebView2HttpHeadersCollectionIterator, ByRef webData As clsWebResData, Optional hmetaTitle As String) As Boolean
    'this function probably belongs in WV2Tools because it specifically outputs a more readable format for clsWebResData
    Dim hName As Long, hVal As Long, sName As String, sVal As String
    If Not iterator Is Nothing Then
        If webData.Headers <> "" Then
            webData.Headers = webData.Headers & vbTab & "[" & hmetaTitle & "]" & vbCrLf & vbTab & "---------------" & vbCrLf
        Else
            webData.Headers = "[" & hmetaTitle & "]" & vbCrLf & vbTab & "---------------" & vbCrLf
        End If
        Do While iterator.HasCurrentHeader
            iterator.GetCurrentHeader hName, hVal
            If hName <> 0 Then sName = StrFromPtr(hName)
            If hVal <> 0 Then sVal = StrFromPtr(hVal)
            If sName <> "" Or sVal <> "" Then
                If webData.Headers <> "" Then
                    webData.Headers = webData.Headers & vbTab & sName & ":" & sVal & vbCrLf
                Else
                    webData.Headers = sName & ":" & sVal & vbCrLf
                End If
            End If
            iterator.MoveNext
        Loop
        webData.Headers = webData.Headers & vbTab & "---------------" & vbCrLf
        IterateHeaders = True
        Exit Function
    End If
    IterateHeaders = False
End Function
