VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebViewScriptCompleteHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'individual event handler to let us know the results of calling ExecuteScript

Implements ICoreWebView2ExecuteScriptCompletedHandler
Public Event wv2ScriptComplete(ByVal sender As wv2, ByRef resultObjectAsJson As String) 'used if we want to declare this object using WithEvents
Public Parent As wv2 'used if we want to create new instances of this object from a local scope and call back to its parent
Public PropLet As String

Private selfRef As clsWebViewScriptCompleteHandler

Private Sub Class_Initialize()
    Set selfRef = Me 'self reference so we dont get destroyed going out of scope, so we can "fire and forget" this handler until it completes, or times out
    PluginManager.AddBroadcaster Me
End Sub

Private Sub Class_Terminate()
    'Set parent = Nothing
    'Set selfRef = Nothing
End Sub

Private Sub ICoreWebView2ExecuteScriptCompletedHandler_Invoke(ByVal errorCode As Long, ByVal resultObjectAsJson As Long)
    Dim jsonResponse As String
    jsonResponse = StrFromPtr(resultObjectAsJson)
    RaiseEvent wv2ScriptComplete(Parent, jsonResponse)
    PluginManager.RemoveBroadcaster Me
    If PropLet <> vbNullString And Not Parent Is Nothing Then CallByName Parent, PropLet, VbLet, jsonResponse
    If Not Parent Is Nothing Then Parent.ExecuteScriptCompletedHandler errorCode, jsonResponse
    Set selfRef = Nothing 'now that script is complete, remove the self reference and this instance will automatically clean itself up
End Sub
