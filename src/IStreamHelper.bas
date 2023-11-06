Attribute VB_Name = "IStreamHelper"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'IStream helper functions - used mainly for receiving/sending WebResource request/response content
'***

Public Function IStreamToString(ByVal istr As IStream) As String
    'read bytes from an IStream into a unicode string
    On Error GoTo err
    Dim id As Long
    Dim stats As STATSTG, sb() As Byte, cSize As Currency, bRead As Currency
    Dim needSize As Currency
    Dim totalRead As Currency
    If Not istr Is Nothing Then
        istr.Stat stats, STATFLAG_DEFAULT
        cSize = stats.cbSize * 10000@
        needSize = cSize
        If cSize > 0 Then
            If cSize > MAXINT_4 Then cSize = MAXINT_4 'prevent istream.read overflow, because it returns LONG - <TODO> change IStream in .tlb to use Currency?
            ReDim sb(0 To needSize - 1) 'size our read buffer to match the stream size
            'IStream.Read returns the number of bytes read, or 0 at end of stream
            'since we can potentially receive a stream with cbSize > max 32bit Long value,
            'we need to keep track separately of total bytes read as Currency
            'this way we keep reading until we get the full stream even if IStream.Read reads 0 bytes but there is more in the stream to read
            Do While bRead < needSize Or (bRead = 0 And totalRead < needSize)
                bRead = bRead + istr.Read(sb(bRead), CLng(cSize))
                totalRead = totalRead + bRead
            Loop
            IStreamToString = StrConv(sb, vbUnicode)
            Exit Function
        End If
    End If
    IStreamToString = ""
Exit Function
err:
    MsgBox "IStreamToString, ID=" & err.Number & ", ERR:" & err.Description
End Function

Public Function IStreamFromArray(ByVal arrayPtr As Long, ByVal length As Long) As stdole.IUnknown
    'create an IUnknown interface with a byte array which can then be passed to anything expecting an IStream
    'pass arrayPtr like VarPtr(myArray(0))
    'length = bytes to be read from ArrayPtr
    On Error GoTo err
    Dim o_hMem As Long
    Dim o_lpMem  As Long
    'allocate memory and create stream from passed byte array
    If arrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, IStreamFromArray
    ElseIf length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal arrayPtr, length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, IStreamFromArray)
            End If
        End If
    End If
    Exit Function
err:
    MsgBox "IStreamFromArray, ID=" & err.Number & ", ERR:" & err.Description
End Function
