Attribute VB_Name = "Types"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'Types - generic win32 and other types
'***

'WINDOWS/DRAWING/SIZING
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        Y As Long
End Type

Public Type GUID
      Data1 As Long
      Data2 As Integer
      Data3 As Integer
      Data4(0 To 7) As Byte
End Type
