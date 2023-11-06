Attribute VB_Name = "Constants"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'Constants
'***

'STATUS/RESULTS
Public Const E_FAIL As Long = &H80004005
Public Const S_OK As Long = &H0
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_InsertS = &H200

'SIZING/WINDOWS/SHAPES
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SW_SHOWMAXIMIZED = 3

'STRINGS/TEXT
Public Const MF_STRING = &H0&: Const MF_POPUP = &H10&
Public Const MF_SEPARATOR = &H800&:  Const MF_GRAYED = &H1&
Public Const CP_ACP = 0 ' default to ANSI code page
Public Const CP_UTF8 = 65001 ' default to UTF-8 code page

'COM
Public Const E_NOINTERFACE As Long = &H80004002

'MEMORY
Public Const VT_BY_REF = &H4000&
Public Const OFFSET_4 = 4294967296#
Public Const MAXINT_4 = 2147483647
