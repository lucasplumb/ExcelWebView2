VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTools 
   Caption         =   "Tools"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   4470
   ClientWidth     =   9120
   OleObjectBlob   =   "frmTools.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub lstDataRequests_Change()
    txtDataRequests.Text = StringFromWebData(dicRequests(lstDataRequests.ListCount - (lstDataRequests.ListIndex + 1)))
End Sub

Private Sub lstDataResponses_Change()
    If dicResponses(lstDataResponses.ListCount - (lstDataResponses.ListIndex + 1)) Is Nothing Then Exit Sub
    txtDataResponses.Text = StringFromWebData(dicResponses(lstDataResponses.ListCount - (lstDataResponses.ListIndex + 1)))
End Sub

Private Sub lstDataResponses_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim sJSONString As String
    Dim vJSON As Variant
    Dim sState As String
    Dim vflat As Variant
    Dim strResult As Variant
    
    Dim req As clsWebResData
    
    Set req = dicResponses(lstDataResponses.ListCount - (lstDataResponses.ListIndex + 1))

    txtDataResponses.Text = beautifyJSON(req.resContent)
    
    txtDataResponses.Text = txtDataResponses.Text & vbCrLf & vbCrLf & beautifyJSON(req.reqContent)
    
End Sub

Private Sub lstDataSearchData_Change()
    On Error Resume Next
    Dim dataidx As Long
    dataidx = CLng(lstDataSearchData.column(1, lstDataSearchData.ListIndex))
    If dicData.Count > dataidx And lstDataSearchData.ListIndex > -1 Then
        txtDataSearchData.Text = StringFromWebData(dicData(CLng(lstDataSearchData.column(1, lstDataSearchData.ListIndex))))
    End If
End Sub

Private Sub lstDataSingle_Change()
    txtDataSingle.Text = StringFromWebData(dicData(lstDataSingle.ListCount - (lstDataSingle.ListIndex + 1)))
End Sub

Private Sub txtDataResponses_Change()

End Sub

Private Sub txtDataSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstDataSearchData.Clear
        lstDataSearchData.ColumnWidths = (";0cm")
        Dim i As Long, res As clsWebResData
        For i = 0 To dicData.Count - 1
            Set res = dicData.Item(i)
            If InStr(res.uri, txtDataSearch) > 0 Then
                lstDataSearchData.AddItem Left$(res.uri, 25), 0
                lstDataSearchData.List(0, 1) = i
                lstDataSearchData.TopIndex = 0
            End If
        Next i
        txtDataSearch.SelStart = 0
        txtDataSearch.SelLength = Len(txtDataSearch.Text)
    End If
End Sub

Private Sub txtDataSearch_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        txtDataSearch.SelStart = 0
        txtDataSearch.SelLength = Len(txtDataSearch.Text)
End Sub


Private Sub UserForm_Initialize()
MultiPage1.ForeColor = &H0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Unload Me
End Sub
