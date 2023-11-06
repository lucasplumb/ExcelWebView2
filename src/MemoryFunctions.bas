Attribute VB_Name = "MemoryFunctions"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'MemoryFunctions
'***

'a little hack similar to GetAddress and AddressOf
Public Function GetBaseAddress(vb_array As Variant) As Long
    Dim vType As Integer
    'First 2 bytes are the VARENUM.
    CopyMemory vType, vb_array, 2
    Dim lp As Long
    'Get the data pointer.
    CopyMemory lp, ByVal VarPtr(vb_array) + 8, 4
    'Make sure the VARENUM is a pointer.
    If (vType And VT_BY_REF) <> 0 Then
        'Dereference it for the variant data address.
        CopyMemory lp, ByVal lp, 4
        'Read the SAFEARRAY data pointer.
        Dim address As Long
        CopyMemory address, ByVal lp, 16
        GetBaseAddress = address
    End If
End Function

'conversion for LPWSTR* frequently used in the WebView2.tlb, which we have converted to LONG* as VBA automation doesnt support LPWSTR*
Public Function StrFromPtr(ByVal lpStr As Long) As String
 Dim bStr() As Byte
 Dim cChars As Long
 On Error Resume Next
 ' Get the number of characters in the buffer
 cChars = lstrlen(lpStr) * 2
 If cChars > 0 Then
  ' Resize the byte array
  ReDim bStr(0 To cChars - 1) As Byte
  ' Grab the ANSI buffer
  Call CopyMemory(bStr(0), ByVal lpStr, cChars)
 End If
 ' Now convert to a VB Unicode string
 StrFromPtr = bStr
End Function

'unused at the moment
Private Function GetStrFromPtrW(ByVal Ptr As Long) As String
    SysReAllocString VarPtr(GetStrFromPtrW), Ptr
End Function

'bitwise shift right
Public Function shr(ByVal value As Long, ByVal Shift As Byte) As Long
    Dim i As Byte
    shr = value
    If Shift > 0 Then
        shr = Int(shr / (2 ^ Shift))
    End If
End Function
'bitwise shift left
Public Function shl(ByVal value As Long, ByVal Shift As Byte) As Long
    shl = value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To Shift
            m = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If
End Function

'(WIP) - HRESULT to WIN32 error code
'#define FACILITY_WIN32 0x0007
'#define __HRESULT_FROM_WIN32(x) ((HRESULT)(x) <= 0 ? ((HRESULT)(x)) : ((HRESULT) (((x) & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000)))
Function hresToWin32(e As Long) As Long
    If e <= 0 Then hresToWin32 = e: Exit Function
    Dim s1 As Long, s2 As Long, s3 As Long
    s1 = e And &HFFFF&
    s2 = shl(&H7&, 16)
    s3 = &H80000000
    hresToWin32 = s1 Or s2 Or s3
End Function
