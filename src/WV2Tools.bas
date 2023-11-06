Attribute VB_Name = "WV2Tools"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'WV2Tools
'contains helper functions for the frmTools UserForm, and general WebView2 specific helpers
'***

Public Sub logResource(ByRef res As clsWebResData, resDict As Dictionary)
    Dim strURIpart As String
    
    resDict.Add resDict.Count, res
    If Len(res.uri) > 25 Then
        strURIpart = Left$(res.uri, 22) & "..."
    Else
        strURIpart = res.uri
    End If
    
    If res.metaTitle = "RESPONSE" Then
        dicResponses.Add dicResponses.Count, res
        frmTools.lstDataResponses.AddItem strURIpart, 0
        frmTools.lstDataResponses.TopIndex = 0
    End If
End Sub

Public Function StringFromWebData(ByRef req As clsWebResData) As String
    Dim sdata As String
    sdata = req.metaTitle & vbCrLf & "{" & vbCrLf & vbTab & req.Method & "|" & req.uri & vbCrLf & vbTab & "Headers:" & vbCrLf & vbTab & req.Headers & vbCrLf
    If Len(req.reqContent) > 0 Then sdata = sdata & vbTab & "REQContent:" & vbCrLf & vbTab & req.reqContent & vbCrLf
    If Len(req.resContent) > 0 Then sdata = sdata & vbTab & "RESContent:" & vbCrLf & vbTab & req.resContent & vbCrLf
    sdata = sdata & "}" & vbCrLf & vbCrLf
    StringFromWebData = sdata
End Function

Public Function OutputRawData(ByRef req As clsWebResData, ByRef txtBox As MSForms.TextBox)
'    Dim sdata As String
'    sdata = req.metaTitle & vbCrLf & "{" & vbCrLf & vbTab & req.Method & "|" & req.URI & vbCrLf & vbTab & "Headers:" & vbCrLf & vbTab & req.Headers & vbCrLf
'    If Len(req.resContent) > 0 Then sdata = sdata & vbTab & "Content:" & vbCrLf & vbTab & req.resContent & vbCrLf
'    sdata = sdata & "}" & vbCrLf & vbCrLf
    'StringFromWebData
    Dim strURIpart As String
    If Len(req.uri) > 25 Then
        strURIpart = Left$(req.uri, 22) & "..."
    Else
        strURIpart = req.uri
    End If
    
    dicData.Add dicData.Count, req
    frmTools.lstDataSingle.AddItem strURIpart, 0
    frmTools.lstDataSingle.TopIndex = 0
    If req.metaTitle = "REQUEST" Then
        dicRequests.Add dicRequests.Count, req
        frmTools.lstDataRequests.AddItem strURIpart, 0
        frmTools.lstDataRequests.TopIndex = 0
    End If
    If req.metaTitle = "RESPONSE" Then
        dicResponses.Add dicResponses.Count, req
        frmTools.lstDataResponses.AddItem strURIpart, 0
        frmTools.lstDataResponses.TopIndex = 0
    End If
    
    If Not txtBox Is Nothing Then
        txtBox.Text = StringFromWebData(req) & txtBox.Text
    End If
    
End Function

Public Function HttpHeadersToString(ByRef iterator As ICoreWebView2HttpHeadersCollectionIterator) As String
    Dim sHeader As String
    Dim hName As Long, hVal As Long, sName As String, sVal As String
    If Not iterator Is Nothing Then
        Do While iterator.HasCurrentHeader
            iterator.GetCurrentHeader hName, hVal
            If hName <> 0 Then sName = StrFromPtr(hName)
            If hVal <> 0 Then sVal = StrFromPtr(hVal)
            If sName <> "" Or sVal <> "" Then
                sHeader = sHeader & sName & ":" & sVal & vbCrLf
            End If
            iterator.MoveNext
        Loop
        HttpHeadersToString = sHeader
        Exit Function
    End If
End Function
