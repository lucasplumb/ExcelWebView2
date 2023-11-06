Attribute VB_Name = "APIFunctions"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'APIFunctions
'***

'---WEBVIEW2LOADER
Public Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal browserExecutableFolder As Long, ByVal userDataFolder As Long, ByVal environmentOptions As Long, ByVal createdEnvironmentCallback As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler) As Long
'---

'---USER32 API
'WINDOWS/FORMS/MENUS
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
'---

'---OLE32 API
Public Declare Function CoInitialize Lib "ole32" (ByRef pvReserved As Any) As Long
Public Declare Function CoInitializeEx Lib "ole32" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
Public Declare Sub CoUninitialize Lib "ole32" ()
'MEMORY
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'STREAMS
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Public Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
'---

'---OLEAUT32 API
'AUTOMATION/COM
Public Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgvt As Any, prgpvarg As Any, pvargResult As Variant) As Long
'---

'---KERNEL32 API
'STRINGS/BYTES
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
'MEMORY
Public Declare Function GetAddrOf Lib "kernel32" Alias "MulDiv" (nNumber As Any, Optional ByVal nNumerator As Long = 1, Optional ByVal nDenominator As Long = 1) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
'ERRORS
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
'---

'---WINMM API
'TIME/TIMING
Public Declare Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetTime Lib "winmm" () As Long
'---

'---GDIPLUS
'DRAWING
Public Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long 'note the parameter change to IUnknown
Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As Long) As Long 'note the parameter change to IUnknown
'---

'---SHLWAPI
'EVENT SINKS
Public Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As Long) As Long
'---
