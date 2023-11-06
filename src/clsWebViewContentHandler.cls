VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebViewContentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'a simple event handler that fires an event back to the clsWebResData object who spawned it when data is received

Implements ICoreWebView2WebResourceResponseViewGetContentCompletedHandler

Public Event WebResourceResponseViewGetContentCompleted(ByRef res As clsWebResData, ByRef Content As IStream)

Private m_resource As clsWebResData 'the parent clsWebResData who created this event handler

Public Property Set Resource(ByRef res As clsWebResData)
    Set m_resource = res 'parent clsWebResData sets this to itself
End Property

Private Sub Class_Initialize()
    'PluginManager.AddBroadcaster Me
End Sub

Private Sub Class_Terminate()
    'unregister this class to all plugins?
    PluginManager.RemoveBroadcaster Me
End Sub

Private Sub ICoreWebView2WebResourceResponseViewGetContentCompletedHandler_Invoke(ByVal errorCode As Long, ByVal Content As IStream)
    RaiseEvent WebResourceResponseViewGetContentCompleted(m_resource, Content) 'fire an event back to my parent class with the content that was received
    PluginManager.RemoveBroadcaster Me
End Sub
