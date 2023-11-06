Attribute VB_Name = "byteConversion"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'byteConversion - helper functions for manipulating bytes/byte arrays
'***


'UTF8 string to byte array
Public Function EncodeToBytes(ByVal sdata As String) As Byte() 'Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sdata), -1, 0, 0, 0, 0) - 1
    If nSize > 0 Then
        ReDim aRetn(0 To nSize - 1) As Byte
        WideCharToMultiByte CP_UTF8, 0, StrPtr(sdata), -1, VarPtr(aRetn(0)), nSize, 0, 0
        EncodeToBytes = aRetn
        Erase aRetn
    Else
        ReDim EncodeToBytes(-1 To -1)
    End If
End Function

'UTF8 byte array to string
Public Function DecodeToBytes(byteArr() As Byte) As Byte()   'Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(byteArr(0)), -1, 0, 0) - 1
    If nSize > 0 Then
        ReDim aRetn(0 To 2 * nSize - 1) As Byte
        MultiByteToWideChar CP_UTF8, 0, VarPtr(byteArr(0)), -1, VarPtr(aRetn(0)), nSize
        DecodeToBytes = aRetn
        Erase aRetn
    Else
        ReDim DecodeToBytes(-1 To -1)
    End If
End Function

'input number of bytes and output a more readable string describing the size
Public Function BytesToXB(value As Currency) As String
    'Dim Value As Currency
    'Value = LargeIntToCurrency(RawValue)
    Select Case value
    Case Is > (2 ^ 30)
        BytesToXB = Round(value / (2 ^ 30), 2) & " GB"
    Case Is > (2 ^ 20)
        BytesToXB = Round(value / (2 ^ 20), 2) & " MB"
    Case Is > (2 ^ 10)
        BytesToXB = Round(value / (2 ^ 10), 2) & " KB"
    Case Else
        BytesToXB = value & " B"
    End Select
End Function

'convert an array of bytes to hex, byte by byte (for readability, memory/stream debugging etc)
Public Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim l As Long, strRet As String
    
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & Hex$(ByteArray(l)) & " "
    Next l
    
    'remove last space at end.
    ByteArrayToHex = Left$(strRet, Len(strRet) - 1)
End Function

'insert 0x00 separators between each character of a string to convert to a wstr or wide string
Public Function StrToWStr(ByVal lpStr As String) As String
    Dim bStr() As Byte
    Dim abData() As Byte
    abData = StrConv(lpStr, vbFromUnicode)
    
    Dim i As Integer, x As Integer
    ReDim bStr(0 To (Len(lpStr) * 2) + 1) As Byte
    For i = 0 To (Len(lpStr) * 2) - 1 Step 2
       bStr(i) = abData(x)
       bStr(i + 1) = 0
       x = x + 1
    Next i
    bStr((Len(lpStr) * 2)) = 0
    
    StrToWStr = bStr
End Function

'convert UTF8 string to UTF16 string
Public Function UTF8to16(str As String) As String
    Dim position As Long, strConvert As String, codeReplace As Integer, strOut As String
    
    strOut = str
    position = InStr(strOut, Chr$(195))
    
    If position > 0 Then
        Do Until position = 0
            strConvert = Mid$(strOut, position, 2)
            codeReplace = Asc(Right$(strConvert, 1))
            If codeReplace < 255 Then
                strOut = Replace(strOut, strConvert, Chr$(codeReplace + 64))
            Else
                strOut = Replace(strOut, strConvert, Chr$(34))
            End If
            position = InStr(strOut, Chr$(195))
        Loop
    End If
    
    UTF8to16 = strOut

End Function





'***
'just experimenting with some "64bit" numbers using the Currency type - these functions are not currently used
'***

'copy LARGE_INTEGER struct into currency, then multiply by 10k - we lose some info in the HighPart, in turn getting a whole number instead of a decimal/"float" value
Public Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    'copy 8 bytes from the large integer to an empty currency
    CopyMemory LargeIntToCurrency, liInput, 8
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
    'Debug.Print "large " & LargeIntToCurrency
End Function
'split a currency value into 2 longs to work with LARGE_INTEGER
Public Function CurrencyToLargeInt(liInput As Currency) As LARGE_INTEGER
    'copy 8 bytes from the large integer to an empty currency
    Dim largeint As LARGE_INTEGER, i1 As Currency
    i1 = liInput
    i1 = i1 / 10000 'divide input by 10k to drop the decimals off
    CopyMemory largeint.LowPart, i1, 4
    CopyMemory largeint.HighPart, ByVal VarPtr(i1) + 4, 4
    CurrencyToLargeInt = largeint
End Function
'when a LARGE_INTEGER is fed in to this function as a byte array,
'we assign a currency value using power of notation for each numeral place in hexadecimal,
Public Function BytesToCurrency(ByRef B() As Byte) As Currency
    'currency max 922,337,203,685,477.5807 <TODO> see if we can do something better to deal with these decimals?
    'technically  922,337,203,685,477
    '256^6 is     281,474,976,710,656
    '256^6 * 3 =  844,424,930,131,968
    'we can kinda support 64bit ULONGLONGs... just nowhere near the max value of a real one
    
    If B(6) > 3 Then
        MsgBox "BytesToCurrency error - overflow"
        BytesToCurrency = -1
        Exit Function
    End If
    
    Dim i As Currency, c As Currency
    For i = 0 To 6
        c = c + (B(i)) * (256@ ^ i)
    Next i
    BytesToCurrency = c
End Function



