VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pluginInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'pluginInterface - interface for creating plugin classes
' - this is essentially a template - DO NOT ADD CODE HERE!
' - instead, create a new class and put "Implements pluginInterface" at the top, then add your code to those methods
' - see pluginBase class module for an example of how to create your own plugin
'***

Private WithEvents m_WebView2Event As clsWebViewEventHandlers
Attribute m_WebView2Event.VB_VarHelpID = -1
Private WithEvents m_ContentEvent As clsWebViewContentHandler
Attribute m_ContentEvent.VB_VarHelpID = -1
Private WithEvents m_ScriptEvent As clsWebViewScriptCompleteHandler
Attribute m_ScriptEvent.VB_VarHelpID = -1
Public Property Get WebView2Event() As clsWebViewEventHandlers

End Property
Public Property Set WebView2Event(ByVal eventHandler As clsWebViewEventHandlers)

End Property
Public Property Get ContentEvent() As clsWebViewContentHandler

End Property
Public Property Set ContentEvent(ByVal eventHandler As clsWebViewContentHandler)

End Property
Public Property Get ScriptEvent() As clsWebViewScriptCompleteHandler

End Property
Public Property Set ScriptEvent(ByVal eventHandler As clsWebViewScriptCompleteHandler)

End Property
Public Property Get NewInstance() As pluginInterface

End Property
