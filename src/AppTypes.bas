Attribute VB_Name = "AppTypes"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'AppTypes - custom types used for the browser
'***

'my own version of 64bit number handling
'used with Currency data type and my functions
'LargeIntToCurrency, CurrencyToLargeInt, BytesToCurrency - in byteConversion module
'(work in progress)
Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type


'BROWSER/WEBVIEW2
Public Type COREWEBVIEW2_CREATION_PROPERTIES
    browserExecutableFolder As Long
    userDataFolder As Long
    Language As Long
End Type

Public Type COREWEBVIEW2_ENVIRONMENT_OPTIONS
    AdditionalBrowserArguments As Long
    Language As Long
    ExperimentalFeaturesEnabled As Long
End Type

'PLUGINS
Public Type TPlugin
    plugin As pluginInterface 'parent
    container As pluginContainer 'container of "listeners"
End Type
