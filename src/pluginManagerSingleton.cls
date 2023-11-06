VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pluginManagerSingleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'clsWebResourceBuilder - helper to create/read WebResourceRequest/WebResourceResponse
'***

Private m_plugins() As TPlugin
Private m_broadcasters As New Dictionary
Private m_listenerContainer As New Dictionary
Private m_listeners As New Dictionary

'here, for each item in the plugins list, we need to create a pluginContainer
'pluginContainer will contain an array of pluginInterfaces which are all the same type as the initial plugin loaded...
'then, every time we want to add a listener, we duplicate the initial plugin and store it in pluginContainer
'the end result looks something like this:
'pluginLoader creates a single instance of our plugin, and passes it to the PluginManager
'PluginManager creates a pluginContainer for each unique plugin instance loaded
'when some "event handler" needs to broadcast its events to this swarm of plugin instances,
'-the handler will call PluginManager.AddBroadcaster(Me), where 'Me' is the target/broadcaster - which is the thing raising the events
'PluginManager.AddBroadcaster will iterate all pluginContainers and insert a new instance of the original plugin instance loaded
'for each new instance loaded, we then set that instances "WebView2Event" property, or whatever, based on the type of event handler passed as 'Me' in AddBroadcaster
'wow

Public Sub AddBroadcaster(eventBroadcaster As Object)
    If m_broadcasters.Exists(ObjPtr(eventBroadcaster)) = False Then
        m_broadcasters.Add ObjPtr(eventBroadcaster), eventBroadcaster
        NewBroadcasterForListeners eventBroadcaster
    Else
        'error - might want to raise an error here, as we probably shouldnt be trying to add multiple broadcasters to the same class instance
        Debug.Print "broadcaster already exists"
    End If
End Sub
Public Sub RemoveBroadcaster(eventBroadcaster As Object)
    Dim broadcasterPlugins() As Long, i As Long
    If m_broadcasters.Exists(ObjPtr(eventBroadcaster)) = True Then
        
        broadcasterPlugins = m_listeners(ObjPtr(eventBroadcaster))
        For i = 0 To UBound(broadcasterPlugins)
            m_listenerContainer(m_listeners(ObjPtr(eventBroadcaster))).RemoveInstance broadcasterPlugins(i)
        Next i
        m_listenerContainer.Remove m_listeners(ObjPtr(eventBroadcaster))
        m_listeners.Remove ObjPtr(eventBroadcaster)
        m_broadcasters.Remove ObjPtr(eventBroadcaster)
    Else
        'error - dont want to actually raise an error here - maybe our classes are just being overzealous with cleanup
        'thats fine - we really want to avoid circular references at all costs to avoid leaking memory
        Debug.Print "broadcaster doesnt exist"
    End If
End Sub

'when we add a broadcaster (i.e., an object instance which has events we want to raise to our plugin), we need to create a new instance of that plugin and set a property based on the type of broadcaster class
'<TODO> this could potentially be improved by searching through existing instances which do not have the "event type" ie, WebView2Event, ContentEvent, etc already set, and reusing those instead of making new instances
Private Sub NewBroadcasterForListeners(eventBroadcaster As Object)
    On Error GoTo e
    Dim plugins() As TPlugin, plugin As pluginInterface, i As Long
    Dim interfaces() As pluginInterface
    Dim broadcasterPlugins() As Long
    plugins = PluginManager.AllPlugins
    For i = 0 To UBound(plugins)
        Select Case TypeName(eventBroadcaster)
            Case "clsWebViewEventHandlers"
                Set plugin = plugins(i).container.NewInstance
                Set plugin.WebView2Event = eventBroadcaster
            Case "clsWebViewContentHandler"
                Set plugin = plugins(i).container.NewInstance
                Set plugin.ContentEvent = eventBroadcaster
            Case "clsWebViewScriptCompleteHandler"
                Set plugin = plugins(i).container.NewInstance
                Set plugin.ScriptEvent = eventBroadcaster
            Case Else
                'RemoveBroadcaster eventBroadcaster 'remove broadcaster created in AddBroadcaster
                err.Raise 1, "pluginManagerSingleton.NewBroadcasterForListeners", "Broadcaster class name must match one of the names in the Select statement."
        End Select
        m_listenerContainer.Add ObjPtr(plugin), plugins(i).container
        'handle multiple plugins being added to the same broadcaster...
        If m_listeners.Exists(ObjPtr(eventBroadcaster)) Then
            broadcasterPlugins = m_listeners(ObjPtr(eventBroadcaster))
            ReDim Preserve broadcasterPlugins(UBound(broadcasterPlugins) + 1)
            broadcasterPlugins(UBound(broadcasterPlugins)) = ObjPtr(plugin)
            m_listeners(ObjPtr(eventBroadcaster)) = broadcasterPlugins
        Else
            ReDim broadcasterPlugins(0)
            broadcasterPlugins(0) = ObjPtr(plugin)
            m_listeners.Add ObjPtr(eventBroadcaster), broadcasterPlugins
        End If
    Next i
    Exit Sub
e:
    MsgBox err.Description, vbCritical, err.source: Exit Sub
End Sub

Private Function CreateContainer(plugin As pluginInterface) As pluginContainer
    Dim pContainer As pluginContainer
    If (Not Not pluginContainers) = 0 Then
        ReDim pluginContainers(0)
    Else
        ReDim Preserve pluginContainers(UBound(pluginContainers) + 1)
    End If
    Set pContainer = New pluginContainer
    Set pluginContainers(UBound(pluginContainers)) = pContainer
    Set pContainer.Template = plugin
    Set CreateContainer = pContainer
End Function

Public Sub LoadPlugin(plugin As pluginInterface)
    Dim pContainer As pluginContainer
    If (Not Not m_plugins) = 0 Then
        ReDim m_plugins(0)
    Else
        ReDim Preserve m_plugins(UBound(m_plugins) + 1)
    End If
    Set m_plugins(UBound(m_plugins)).plugin = plugin

    Set m_plugins(UBound(m_plugins)).container = New pluginContainer
    Set pContainer = m_plugins(UBound(m_plugins)).container
    Set pContainer.Template = plugin
    
End Sub

Public Property Get Count()
    If (Not Not plugins) = 0 Then
        Count = 0
    Else
        Count = UBound(plugins) + 1
    End If
End Property

Public Property Get AllPlugins() As TPlugin()
    AllPlugins = m_plugins
End Property

Private Sub Class_Initialize()

End Sub
Public Sub Kill()
    Class_Terminate
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    
    If (Not Not m_plugins) <> 0 Then
        For i = 0 To UBound(m_plugins)
            Set m_plugins(i).container = Nothing
            Set m_plugins(i).plugin = Nothing
        Next i
    End If
    Erase m_plugins
    
    m_broadcasters.RemoveAll
    m_listeners.RemoveAll
    m_listenerContainer.RemoveAll
    Set m_broadcasters = Nothing
    Set m_listeners = Nothing
    Set m_listenerContainer = Nothing
End Sub

