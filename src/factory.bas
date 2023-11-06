Attribute VB_Name = "factory"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'factory
'this modules job is just to create objects
'***



'the wv2 class is the "core" browser object - it is "the browser" view
'simply create a new instance of it and the class will handle setting itself up
Public Function NewTab() As wv2
    Set NewTab = New wv2
End Function

