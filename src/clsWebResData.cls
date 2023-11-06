VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebResData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'clsWebResData - a separate event handler class for WebResource content,
'created on initial WebResource request/response event, stores information about the event,
'spawns a clsWebViewContentHandler if one doesn't already exist for this resource, which
'then fires an event back to us here when content is available in the stream -
'this way the content can be linked back to the initial event when the data finally becomes available
'NOTE: this class is also used for historical data to view later, from frmTools or somewhere else - so instances of this class should not be destroyed,
'unless you don't care to keep track of past information
'***

Public metaTitle As String
Public reqContent As String
Public resContent As String
Public uri As String
Public Method As String
Public Headers As String
Public Context As COREWEBVIEW2_WEB_RESOURCE_CONTEXT
Public StatusCode As Long
Public ReasonPhrase As Long
Public sender As WebView2_edit.ICoreWebView2
Private WithEvents c_Handler As clsWebViewContentHandler
Attribute c_Handler.VB_VarHelpID = -1

Public Property Get contentHandler() As clsWebViewContentHandler
    'spawn a handler for when content is received - set this instance of clsWebResData as the parent via contentHandler.resource
    If c_Handler Is Nothing Then
        Set c_Handler = New clsWebViewContentHandler
    End If
    Set contentHandler = c_Handler
    Set contentHandler.Resource = Me
    PluginManager.AddBroadcaster contentHandler
End Property

Private Sub c_Handler_WebResourceResponseViewGetContentCompleted(ByRef res As clsWebResData, Content As IStream)
    Dim strStream As String
    If Not Content Is Nothing Then
        strStream = IStreamToString(Content) 'open and read the IStream data into a string
        Me.resContent = Me.resContent & strStream 'append received data in case more becomes available later
    End If
End Sub

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    PluginManager.RemoveBroadcaster c_Handler
    Set c_Handler = Nothing
End Sub
