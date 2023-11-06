Attribute VB_Name = "pluginLoader"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'pluginLoader - load and register all plugins with a single call
'***

Private plugins() As pluginInterface
Private m_plugins As pluginManagerSingleton

'get or spawn a plugin manager instance
Public Property Get PluginManager() As pluginManagerSingleton 'Public Function
    If m_plugins Is Nothing Then
        Set m_plugins = New pluginManagerSingleton
    End If
    Set PluginManager = m_plugins
End Property
Private Sub AddPlugin(plugin As pluginInterface)
    If (Not Not plugins) = 0 Then
        ReDim plugins(0)
    Else
        ReDim Preserve plugins(UBound(plugins) + 1)
    End If
    Set plugins(UBound(plugins)) = plugin
End Sub
Public Sub LoadPlugins()
    'add your plugins here
    'AddPlugin New pluginBase
    AddPlugin New pluginExampleCls
    'add more plugins as desired above, just call AddPlugin(New myPluginClass)
    
    'tell plugin manager to load plugins then clear this modules references to them so PluginManager controls their life time
    Dim i As Integer
    For i = 0 To UBound(plugins)
        PluginManager.LoadPlugin plugins(i)
        Set plugins(i) = Nothing
    Next i
    
    Erase plugins
End Sub
