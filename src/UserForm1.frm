VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ExcelWebView2"
   ClientHeight    =   15000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22080
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'WIP
'Dim HostObj As New HostObjectClass
'
'Function AddHostObjectToScript_OleExp(ObjName As String, obj1 As Object) As Boolean
'On Error GoTo err
'    Dim ICoreWebView2A As ICoreWebView2
'    Dim NN As IUnknown
'    Set NN = webviewWindow
'    Set ICoreWebView2A = NN
'    ICoreWebView2A.AddHostObjectToScript StrPtr(ObjName), obj1
'    AddHostObjectToScript_OleExp = True
'Exit Function
'err:
'MsgBox "errhost:" & err.Number & "," & err.Description
'End Function
'
'Private Sub AddHostObject()
'    If AddHostObjectToScript_NEW("HostClass", HostObj) Then
'        ExecuteScript "const HostClassA=window.chrome.webview.hostObjects.HostClass;"
'        ExecuteScript "const HostClassA2=window.chrome.webview.hostObjects.sync.HostClass;"
'        'ExecuteScript "alert(HostClassA2.ClassAdd(33,44));"
'        'DoAddHostObjectToScript.Enabled = False
'    End If
'End Sub

Private Sub browserTabs_Change()
    If (Not Not g_wv2) <> 0 Then
        g_wv2(browserTabs.SelectedItem.index).Focus
        g_selectedTabIndex = browserTabs.SelectedItem.index
    End If
End Sub

Private Sub cmdBack_Click()
    ActiveBrowserTab.GoBack
End Sub

Private Sub cmdForward_Click()
    ActiveBrowserTab.GoForward
End Sub

Private Sub cmdNewTab_Click()
    factory.NewTab
End Sub

Private Sub cmdStopReload_Click()
    Dim i As Long
    If cmdStopReload.Caption = "X" Then
        For i = 0 To UBound(g_wv2)
            g_wv2(i).StopLoading
        Next i
    Else
        ActiveBrowserTab.Reload
    End If
End Sub

Private Sub CommandButton10_Click()
    ActiveBrowserTab.ExecuteScript "eval(1+1);"
    pluginExample.Search1
    pluginExample.Search2
    pluginExample.Search3
    pluginExample.Search4
End Sub

Private Sub CommandButton7_Click()
    g_wv2(browserTabs.SelectedItem.index).OpenDevTools
End Sub

Private Sub txtUrl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ActiveBrowserTab.OpenUrl txtUrl.Text
        txtUrl.SelStart = 0
        txtUrl.SelLength = Len(txtUrl.Text)
    End If
End Sub

Private Sub UserForm_Initialize()
    factory.NewTab
    frmTools.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    WV2Globals.CleanUp
    frmTools.Hide
    Unload Me
End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    
End Sub




