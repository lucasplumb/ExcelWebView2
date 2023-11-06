VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pluginContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'pluginContainer - contains duplicate instances of some type of pluginInterface, acting as a sort of delegate
'for more information on how plugins are managed, see the pluginManagerSingleton class
'***

'this is our "container" - these should act as essentially "shadow copies" of our plugin, whose only job is to handle events
'keep in mind, however, that they are *literally* copies of the plugin class, except they will have their "WithEvents" members actually set to an object which raises events
'it is important that they do not keep any state information or execute functions - they should raise everything up to the "parent" to actually act on the events

Private m_pluginInstance As Dictionary

'this is the "parent" - it should not have any of its "WithEvents" or "listener" members assigned to - rather, the "shadow copies" in the container
'will forward all their events to this parent which will execute code to act on the events and manipulate data
Private m_pluginTemplate As pluginInterface

Public Property Set Template(newTemplate As pluginInterface)
    Set m_pluginTemplate = newTemplate
End Property

Public Property Get NewInstance() As pluginInterface
    Dim newInterface As pluginInterface
    Set newInterface = m_pluginTemplate.NewInstance
    m_pluginInstance.Add ObjPtr(newInterface), newInterface
    Set NewInstance = newInterface
End Property

Public Sub RemoveInstance(instancePtr As Long)
    If m_pluginInstance.Exists(instancePtr) Then
        Set m_pluginInstance(instancePtr) = Nothing
        m_pluginInstance.Remove instancePtr
    End If
End Sub

Private Sub Class_Initialize()
    Set m_pluginInstance = New Dictionary
End Sub

Private Sub Class_Terminate()
    m_pluginInstance.RemoveAll
End Sub
