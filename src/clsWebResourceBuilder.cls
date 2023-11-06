VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebResourceBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'clsWebResourceBuilder - helper to create/read WebResourceRequest/WebResourceResponse
'***

Private m_WV2Env_2 As ICoreWebView2Environment2

Private m_uri As String
Private m_method As String
Private m_postData As String
Private m_reqHeaders As ICoreWebView2HttpRequestHeaders
Private m_resHeaders As ICoreWebView2HttpResponseHeaders
Private m_content As String
Private m_reasonPhrase As String
Private m_statusCode As HttpStatusCode
Private m_postDataStream As IStream

Public Function Request() As ICoreWebView2WebResourceRequest
    Set Request = m_WV2Env_2.CreateWebResourceRequest( _
        m_uri, _
        m_method, _
        POSTData, _
        RequestHeaders)
End Function
Public Property Get Response() As ICoreWebView2WebResourceResponse
    Set Response = m_WV2Env_2.CreateWebResourceResponse( _
        Content, _
        StatusCode, _
        ReasonPhrase, _
        ResponseHeaders)
End Property

Public Property Get uri() As String
    uri = m_uri
End Property
Public Property Let uri(address As String)
    m_uri = address
End Property
Public Property Get ReasonPhrase() As String
    ReasonPhrase = m_reasonPhrase
End Property
Public Property Let ReasonPhrase(phrase As String)
    m_reasonPhrase = phrase
End Property
Public Property Get Method() As HTTPMethod
    Method = StrToHTTPMethod(m_method)
End Property
Public Property Let Method(requestMethod As HTTPMethod)
    m_method = HTTPMethodToStr(requestMethod)
End Property
Public Property Get POSTData() As IStream
Attribute POSTData.VB_Description = "Set this property using a String, retrieving it will return an IStream"
    'create an IStream from m_postData
    Dim postDataBytes() As Byte
    postDataBytes = EncodeToBytes(m_postData)
    If UBound(postDataBytes) > -1 Then
        Set m_postDataStream = IStreamFromArray(VarPtr(postDataBytes(0)), UBound(postDataBytes) + 1)
    Else 'need to handle an empty POST data...
        Set m_postDataStream = IStreamFromArray(0&, UBound(postDataBytes) + 1)
    End If
    Set POSTData = m_postDataStream
End Property
Public Property Let POSTData(str As Variant) 'using str as variant lets us assign the POSTData property as a string but get it as an IStream
    m_postData = str
End Property
Public Property Get Content() As IStream
Attribute Content.VB_Description = "Set this property using a String, retrieving it will return an IStream"
    'create an IStream from m_content
    Dim contentBytes() As Byte
    contentBytes = EncodeToBytes(m_content)
    Set Content = IStreamFromArray(VarPtr(contentBytes(0)), UBound(contentBytes) + 1)
End Property
Public Property Let Content(str As Variant) 'using str as variant lets us assign the Content property as a string but get it as an IStream
    m_content = str
End Property
Public Property Let StatusCode(responseStatusCode As HttpStatusCode)
    m_statusCode = responseStatusCode
End Property
Public Property Get StatusCode() As HttpStatusCode
    StatusCode = m_statusCode
End Property
Public Property Get RequestHeaders() As String
    If m_reqHeaders Is Nothing Then
        RequestHeaders = ""
    Else
        RequestHeaders = HttpHeadersToString(m_reqHeaders.GetIterator)
    End If
End Property
Public Function SetRequestHeader(Name As String, value As String)
    Set m_reqHeaders = Request.Headers
    m_reqHeaders.SetHeader Name, value
End Function
Public Function RemoveRequestHeader(Name As String)
    Set m_reqHeaders = Request.Headers
    m_reqHeaders.RemoveHeader Name
End Function
Public Property Get ResponseHeaders() As String
    If m_resHeaders Is Nothing Then
        ResponseHeaders = ""
    Else
        ResponseHeaders = HttpHeadersToString(m_resHeaders.GetIterator)
    End If
End Property
Public Function SetResponseHeader(Name As String, value As String)
    Set m_resHeaders = Response.Headers
    m_resHeaders.SetHeader Name, value
End Function
Public Function RemoveResponseHeader(Name As String)
    Set m_resHeaders = Response.Headers
    m_resHeaders.RemoveHeader Name
End Function

Private Sub Class_Initialize()
    Set m_WV2Env_2 = g_Env
End Sub

Public Sub PostRequest(wv2Object As wv2)
    wv2Object.NavigateWithResource Request
End Sub

