Attribute VB_Name = "IUnknownFake"
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'IUnknownFake - *NOTE*
'-these functions were previously used with a factory class object to impersonate some generic IUnknown,
'especially the WebView2 event handler interfaces. creating instances of a class object and assigning/instantiating an IUnknown
'variable can allow you to perform some COM trickery and overwrite the Invoke function with your own version.
'this was a quick way to manage the events WebView2 fires but are unsupported by VBA in the early development phase.
'these functions are no longer used at the moment, but i am including them for legacy/experimental/development purposes.
'if you wish to use them, you will need to create a class module with its own internal m_This/m_VTable/m_pVTable variables,
'declare fake_IUnknown As Object,
'then call Set fake_IUnknown = InitializeVTable(m_This, m_VTable, m_pVTable)
'***

Public Type IUnknownVtblObj
    pVTable As Long
End Type
Public Type IUnknownVtbl
    VTable(3) As Long
End Type
'Private m_This as IUnknownVtblObj '<- this should be allocated in a class object to be used as some generic IUnknown, then passed into the InitializeVTable function
'Private m_VTable As IUnknownVtbl '<- this should be allocated in a class object to be used as some generic IUnknown, then passed into the InitializeVTable function
'Private m_pVTable As Long '<- this should be allocated in a class object to be used as some generic IUnknown, then passed into the InitializeVTable function

Public Function InitializeVTable(ByRef this As IUnknownVtblObj, ByRef m_VTable As IUnknownVtbl, ByRef m_pVTable As Long) As IUnknown
    If m_pVTable = 0 Then
        With m_VTable
            .VTable(0) = GetAddress(AddressOf QueryInterface1)
            .VTable(1) = GetAddress(AddressOf AddRef1)
            .VTable(2) = GetAddress(AddressOf Release1)
            .VTable(3) = GetAddress(AddressOf Invoke1)

            m_pVTable = VarPtr(.VTable(0))
        End With
    End If
    
    With this
        .pVTable = m_pVTable
        CopyMemory InitializeVTable, VarPtr(.pVTable), 4
    End With
End Function
Public Function QueryInterface1(this As IUnknownVtblObj, riid As Long, pvObj As Long) As Long
    'not implemented
    pvObj = 0
    QueryInterface1 = E_NOINTERFACE
End Function

Public Function AddRef1(this As IUnknownVtblObj) As Long
   'not implemented
End Function

Public Function Release1(this As IUnknownVtblObj) As Long
   'not implemented
End Function

Public Function Invoke1(this As IUnknownVtblObj, Optional ByVal a1 As Long = 0, Optional ByVal a2 As Long = 0) As Long
    'not implemented
End Function

'feed some fake_IUnknown declared As Object in to get its IUnknown pointer back
'just a helper since ObjPtr won't work with a faked IUnknown
Public Function IUnkObjPtr(ByVal pObj As IUnknown) As Long
    IUnkObjPtr = VarPtr(pObj)
End Function
