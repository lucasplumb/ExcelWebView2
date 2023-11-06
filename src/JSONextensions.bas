Attribute VB_Name = "JSONextensions"
Option Explicit
'WIP
'Lucas Plumb @ 2023
'functions to help interpret JSON requests/responses into a more digestible,
'VBA friendly format for replication/manipulation
Function beautifyJSON(ByVal sJSONString As String) As String
    Dim vJSON As Variant
    Dim vflat As Variant
    Dim sState As String
    Dim out As String, tmpStr As String, lines() As String
    Dim i As Integer, v As Variant, x As Variant
    Dim aheader() As Variant, adata() As Variant
    Dim linecounter As Integer
    
    JSon.Parse sJSONString, vJSON, sState
    
    Select Case True
    Case sState <> "Object"
        'parseData = -1
        Debug.Print "data not an object"
        Exit Function
    Case Else
        Debug.Print "got object"
        
        JSon.Flatten vJSON, vflat
'        For Each v In vFlat.Items
'            Debug.Print v
'        Next v
        out = out & "-----" & vbCrLf
        For Each v In vflat.Keys
            out = out & v & "-" & vflat(v) & "///" & vbCrLf
            'JSon.ToArray vFlat(v), adata, aheader
            'Debug.Print aheader(0) & "-" & adata(0) & vbCrLf
        Next v

        'For i = 0 To UBound(adata)
            'Debug.Print adata
        'Next i
        out = out & "-----" & vbCrLf
                tmpStr = JSon.Serialize(vJSON)
                tmpStr = Replace(tmpStr, """", """""")
                lines = Split(tmpStr, vbCrLf)
                For i = 0 To UBound(lines)
                    If linecounter >= 23 Then
                        linecounter = 0
                        If Right$(out, 6) = " & _" & vbCrLf Then
                            out = Left$(out, Len(out) - 6) & vbCrLf
                        End If
                        out = out & "req = req & _ " & vbCrLf
                    Else
                        linecounter = linecounter + 1
                    End If
                    lines(i) = Replace(lines(i), """", """""", 1, 1)
                    If i = 0 Then
                        lines(0) = """" & lines(0)
                    End If
                    lines(i) = lines(i) & """"
                    If i < UBound(lines) Then
                        lines(i) = lines(i) & " & _"
                        lines(i) = lines(i) & vbCrLf
                    End If
                    If Left$(Replace(lines(i), vbTab, ""), 1) <> """" Then
                        lines(i) = Replace(lines(i), Left$(Replace(lines(i), vbTab, ""), 1), """" & Left$(Replace(lines(i), vbTab, ""), 1), 1, 1)
                    End If
                    If i = 0 Then
                        lines(i) = "req = " & lines(i)
                    End If
                    out = out & lines(i)
                Next i
                
                
    End Select
    
    beautifyJSON = out
End Function
