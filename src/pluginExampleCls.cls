VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pluginExampleCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'pluginBase - example base class for a plugin to implement your own custom code on top of the normal browser functions
'the code in this class should essentially only *handle* events, it should not store any kind of state information or variables unless modified to do so
'the reason for this is due to the nature of handling Events in VBA, WithEvents can only be used with a single object - therefore,
'we create multiple "shadow copies" of this class via PluginManager every time a separate class i.e. 'clsWebViewEventHandlers' is created
'recommend creating a standard module which this class interacts with if you need to store any sort of state/information for your plugin to use
'***

Implements pluginInterface

Private WithEvents m_WebView2Event As clsWebViewEventHandlers
Attribute m_WebView2Event.VB_VarHelpID = -1
Private WithEvents m_ContentEvent As clsWebViewContentHandler
Attribute m_ContentEvent.VB_VarHelpID = -1
Private WithEvents m_ScriptEvent As clsWebViewScriptCompleteHandler
Attribute m_ScriptEvent.VB_VarHelpID = -1

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    'clean up any data your plugin created
End Sub

Private Sub m_ContentEvent_WebResourceResponseViewGetContentCompleted(res As clsWebResData, Content As WebView2_edit.IStream)
    'Debug.Print res.resContent 'server response content
    If res.Method = HTTPMethodToStr(HTTP_GET) And res.uri = "https://en.wikipedia.org/wiki/WebView2" Then
        pluginExample.LastWebContentResponse = res.resContent
        pluginExample.ResponseReceived = True
    End If
End Sub

Private Sub m_ScriptEvent_wv2ScriptComplete(ByVal sender As wv2, resultObjectAsJson As String)
    pluginExample.ScriptComplete = True
    pluginExample.LastExecutedScriptResponse = resultObjectAsJson
End Sub

Private Sub m_WebView2Event_DocumentTitleChanged(sender As WebView2_edit.ICoreWebView2, args As Long)
    
End Sub

Private Sub m_WebView2Event_DOMContentLoaded(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2DOMContentLoadedEventArgs)
    pluginExample.LastExecutedScriptResponse = "" 'reset script response on each page load
End Sub

Private Sub m_WebView2Event_NavigationCompleted(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationCompletedEventArgs)
    
End Sub

Private Sub m_WebView2Event_NavigationStarting(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationStartingEventArgs)
End Sub

Private Sub m_WebView2Event_WebResourceRequested(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceRequestedEventArgs)
    'Debug.Print StrFromPtr(sender.source)
    pluginExample.ResponseReceived = False
End Sub

Private Sub m_WebView2Event_WebResourceResponseReceived(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceResponseReceivedEventArgs)
    
End Sub

Private Sub m_WebView2Event_wv2ControllerReady(createdController As WebView2_edit.ICoreWebView2Controller)

End Sub

Private Sub m_WebView2Event_wv2EnvironmentReady(createdEnvironment As WebView2_edit.ICoreWebView2Environment)

End Sub


'boilerplate
Private Property Get pluginInterface_NewInstance() As pluginInterface
    Set pluginInterface_NewInstance = New pluginExampleCls 'always set this to the class of your custom plugin
End Property
Private Property Get pluginInterface_ContentEvent() As clsWebViewContentHandler
    Set pluginInterface_ContentEvent = m_ContentEvent
End Property

Private Property Set pluginInterface_ContentEvent(ByVal RHS As clsWebViewContentHandler)
    Set m_ContentEvent = RHS
End Property

Private Property Set pluginInterface_ScriptEvent(ByVal RHS As clsWebViewScriptCompleteHandler)
    Set m_ScriptEvent = RHS
End Property

Private Property Get pluginInterface_ScriptEvent() As clsWebViewScriptCompleteHandler
    Set pluginInterface_ScriptEvent = m_ScriptEvent
End Property

Private Property Set pluginInterface_WebView2Event(ByVal RHS As clsWebViewEventHandlers)
    Set m_WebView2Event = RHS
End Property
Private Property Get pluginInterface_WebView2Event() As clsWebViewEventHandlers
    Set pluginInterface_WebView2Event = m_WebView2Event
End Property
'/boilerplate

