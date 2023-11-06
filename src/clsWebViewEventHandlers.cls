VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebViewEventHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'clsWebViewEventHandlers
'this class is instantiated along with each wv2 instance and passes most common
'browser events to other classes for processing
'***

Implements ICoreWebView2WebResourceRequestedEventHandler
Implements ICoreWebView2WebResourceResponseReceivedEventHandler
Implements ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
Implements ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
Implements ICoreWebView2DOMContentLoadedEventHandler
Implements ICoreWebView2NavigationCompletedEventHandler
Implements ICoreWebView2NavigationStartingEventHandler
Implements ICoreWebView2DocumentTitleChangedEventHandler

Public Event wv2EnvironmentReady(createdEnvironment As WebView2_edit.ICoreWebView2Environment)
Public Event wv2ControllerReady(createdController As WebView2_edit.ICoreWebView2Controller)

Public Event WebResourceResponseReceived(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceResponseReceivedEventArgs)
Public Event WebResourceRequested(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceRequestedEventArgs)
Public Event DOMContentLoaded(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2DOMContentLoadedEventArgs)
Public Event NavigationCompleted(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationCompletedEventArgs)
Public Event NavigationStarting(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationStartingEventArgs)
Public Event DocumentTitleChanged(sender As WebView2_edit.ICoreWebView2, args As Long)



Private Sub Class_Initialize()
    PluginManager.AddBroadcaster Me
End Sub

Private Sub Class_Terminate()
    'unregister this class to all plugins?
    PluginManager.RemoveBroadcaster Me
End Sub

Public Sub ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_Invoke(ByVal errorCode As Long, ByVal createdController As WebView2_edit.ICoreWebView2Controller)
    RaiseEvent wv2ControllerReady(createdController)
End Sub

Private Sub ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_Invoke(ByVal errorCode As Long, ByVal createdEnvironment As WebView2_edit.ICoreWebView2Environment)
    RaiseEvent wv2EnvironmentReady(createdEnvironment)
End Sub

Private Sub ICoreWebView2DOMContentLoadedEventHandler_Invoke(ByVal sender As WebView2_edit.ICoreWebView2, ByVal args As WebView2_edit.ICoreWebView2DOMContentLoadedEventArgs)
    RaiseEvent DOMContentLoaded(sender, args)
End Sub

Private Sub ICoreWebView2NavigationCompletedEventHandler_Invoke(ByVal sender As WebView2_edit.ICoreWebView2, ByVal args As WebView2_edit.ICoreWebView2NavigationCompletedEventArgs)
    RaiseEvent NavigationCompleted(sender, args)
End Sub

Private Sub ICoreWebView2NavigationStartingEventHandler_Invoke(ByVal sender As WebView2_edit.ICoreWebView2, ByVal args As WebView2_edit.ICoreWebView2NavigationStartingEventArgs)
    RaiseEvent NavigationStarting(sender, args)
End Sub

Private Sub ICoreWebView2WebResourceRequestedEventHandler_Invoke(ByVal sender As WebView2_edit.ICoreWebView2, ByVal args As WebView2_edit.ICoreWebView2WebResourceRequestedEventArgs)
    RaiseEvent WebResourceRequested(sender, args)
End Sub

Private Sub ICoreWebView2WebResourceResponseReceivedEventHandler_Invoke(ByVal sender As WebView2_edit.ICoreWebView2, ByVal args As WebView2_edit.ICoreWebView2WebResourceResponseReceivedEventArgs)
    RaiseEvent WebResourceResponseReceived(sender, args)
End Sub

Private Sub ICoreWebView2DocumentTitleChangedEventHandler_Invoke(ByVal sender As WebView2_edit.ICoreWebView2, args As Long)
    'args is null
    RaiseEvent DocumentTitleChanged(sender, args)
End Sub
