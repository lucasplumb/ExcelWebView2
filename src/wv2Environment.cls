VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wv2Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ExcelWebView2 by Lucas Plumb @ 2023
'WebView2Environment object, do NOT instantiate this manually!
'simply create a new wv2 class object using:
'    Dim wv As wv2
'    Set wv = New wv2
'the wv2 object itself will handle everything needed for setup
'this class object will automatically be instantiated if needed by the wv2 object

Private m_WebView2Environment As ICoreWebView2Environment
Private WithEvents m_webViewHandlers As clsWebViewEventHandlers
Attribute m_webViewHandlers.VB_VarHelpID = -1

Public Event wv2CtrlReady(ByRef createdController As WebView2_edit.ICoreWebView2Controller)
Public Event wv2EnvReady(ByRef createdEnvironment As WebView2_edit.ICoreWebView2Environment)
Public Event wv2Ready(ByRef env As wv2Environment)


Public Property Get this() As ICoreWebView2Environment
    If m_WebView2Environment Is Nothing Then
        If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, handler) <> S_OK Then
            'MessageBox 0, "Failed to create environment", "Error", 0
            Unload Me
        Else
            'we could raise an "environment creation successful" event here, but we do it in "handler" instead
        End If
    Else
        Set this = m_WebView2Environment
    End If
End Property

Public Sub Init()
    If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, handler) <> S_OK Then
        'MessageBox 0, "Failed to create environment", "Error", 0
        Unload Me
    Else
        'we could raise an "environment creation successful" event here, but we do it in "handler" instead
    End If
    RaiseEvent wv2Ready(Me)
End Sub

Public Property Get handler() As clsWebViewEventHandlers
    If m_webViewHandlers Is Nothing Then
        Set m_webViewHandlers = New clsWebViewEventHandlers
    End If
    Set handler = m_webViewHandlers
End Property

Public Property Set handler(m_handler As clsWebViewEventHandlers)
    Set m_webViewHandlers = m_handler
End Property

Private Sub Class_Initialize()
    If Not g_wv2Env Is Nothing Then
        MsgBox "wv2Environment error - class created when environment already exists", vbCritical, "Error"
    End If
    Set g_wv2Env = Me
End Sub

Private Sub Class_Terminate()
    Set m_WebView2Environment = Nothing
    Set m_webViewHandlers = Nothing
End Sub

Private Sub m_webViewHandlers_wv2ControllerReady(ByRef createdController As WebView2_edit.ICoreWebView2Controller)
    Debug.Print "controller ready in env"
    RaiseEvent wv2CtrlReady(createdController)
End Sub

Private Sub m_webViewHandlers_wv2EnvironmentReady(ByRef createdEnvironment As WebView2_edit.ICoreWebView2Environment)
    Set m_WebView2Environment = createdEnvironment
    Set g_Env = createdEnvironment
    RaiseEvent wv2EnvReady(g_Env)
    Debug.Print "environment ready in env"
End Sub

