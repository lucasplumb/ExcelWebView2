VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HostObjectClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'WIP

Option Explicit
Dim Itemv As String

Public JsResult As String
Function Eval(js As String) As String
    JsResult = ""
    Dim Start As Long
    Start = timeGetTime
    ExecuteScript "HostClassA2.ReturnVal(" & js & ")"
    
    While JsResult = "" And timeGetTime - Start < 1000
    DoEvents:    Wend
    Eval = JsResult
End Function

Public Sub ReturnVal(s)
    JsResult = s
End Sub
'
'
'Function Eval(js As String) As String
'    JsResult = ""
'    Dim Start As Long, WaitStart As Long
'    Start = timeGetTime
'
'    ExecuteScript "HostClassA2.ReturnVal(" & js & ")"
'
'    WaitStart = timeGetTime
'
'    While JsResult = "" And timeGetTime - Start < 1000
'    DoEvents:    Wend
'    '返回结果用时
'    Debug.Print "Get Result Used Time：" & timeGetTime - Start & "MS" & ",Wait Used:" & timeGetTime - WaitStart & "MS"
'    Eval = JsResult
'End Function
'
'Public Sub ReturnVal(s)
'    'js返回值给VB宿主类 MsgBox "RunJS的结果：" & s
'    JsResult = s
'End Sub


'Function Eval(js As String) As String
'    Dim FirstValue As String
'    FirstValue = Now:    JsResult = FirstValue
'    ExecuteScript "HostClassA2.ReturnVal(" & js & ")"
'    While JsResult = FirstValue
'        DoEvents
'    Wend
'    Eval = JsResult
'End Function
Private Sub Class_Initialize()
Itemv = "Now:" & Now
End Sub

Public Function ClassAdd(a, B) As Long
    ClassAdd = a + B
End Function
 

Public Property Get Item() As String
    Item = Itemv
End Property

Public Property Let Item(ByVal vNewValue As String)
    Itemv = vNewValue
End Property
'function () { [native code] }
Public Sub test()
    MsgBox "Call EXE Class Test from JS"
End Sub
Public Sub SendData(s)
MsgBox "s=" & s
End Sub


Public Function GetV2(value1 As String) As Long
'    MsgBox "a=" & a
'    MsgBox "b=" & b
'    MsgBox a + b
    'ClassAdd = a + b
    'GetV2 = "{'value':77}"
    Dim buffer() As Byte
    value1 = EncodeToBytes("test abc")
    'value1="ddd"
    'GetV2 = buffer
    MsgBox "test"
 
    'CLng(a) + CLng(b)
End Function

Public Sub PutData(ParamArray a() As Variant)
MsgBox UBound(a)
 
End Sub


