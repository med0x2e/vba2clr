Option Explicit

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal arg1 As LongPtr, ByVal arg2 As LongPtr, ByVal arg3 As Integer, ByVal arg4 As Integer, ByVal arg5 As Long, ByRef arg6 As Integer, ByRef arg7 As LongPtr, ByRef arg8 As Variant) As Long

Private Declare PtrSafe Function CLRCreateInstance Lib "mscoree.dll" (ByRef arg1 As Any, ByRef arg2 As Any, ByRef arg3 As IUnknown) As Long

Private Function DispCallWrapper(ByVal IObj As IUnknown, ByVal vftOffset As Integer, ParamArray args() As Variant)
    
    Dim out As Variant
    Dim ret As Long
    Dim prgvt(0 To 2) As Integer
    Dim prgpvarg(0 To 2) As LongPtr
    Dim index As Long
    For index = 0 To 2
        prgvt(index) = VarType(args(index))
        prgpvarg(index) = VarPtr(args(index))
    Next index
    ret = DispCallFunc(ObjPtr(IObj), vftOffset * Len(prgpvarg(0)), 4, vbLong, 3, prgvt(0), prgpvarg(0), out)

End Function

Private Function Init(ByVal CLRVersion As String) As mscoree.CorRuntimeHost
    
    Dim ret As Long
    Dim pMetaHost As IUnknown, pRuntimeInfo As IUnknown, pCorRuntimeHost As IUnknown
    
     'CLSID_CLRMetaHost "9280188D-0E8E-4867-B30C-7FA83884E8DE"
    Dim GUID_CMH&(0 To 7)
    GUID_CMH(0) = &H9280188D: GUID_CMH(1) = &H48670E8E: GUID_CMH(2) = &HA87F0CB3: GUID_CMH(3) = &HDEE88438
    
    'CLSID IID_ICLRMetaHost "D332DB9E-B9B3-4125-8207-A14884F53216"
    Dim GUID_ICMH&(0 To 7)
    GUID_ICMH(0) = &HD332DB9E: GUID_ICMH(1) = &H4125B9B3: GUID_ICMH(2) = &H48A10782: GUID_ICMH(3) = &H1632F584
    
    ret = CLRCreateInstance(GUID_CMH(0), GUID_ICMH(0), pMetaHost)
    
    'CLSID IID_ICLRRuntimeInfo "BD39D1D2-BA2F-486A-89B0-B4B0CB466891"
    Dim GUID_ICRI&(0 To 7)
    GUID_ICRI(0) = &HBD39D1D2: GUID_ICRI(1) = &H486ABA2F: GUID_ICRI(2) = &HB0B4B089: GUID_ICRI(3) = &H916846CB

    Call DispCallWrapper(pMetaHost, 3, StrPtr(CLRVersion), VarPtr(GUID_ICRI(0)), VarPtr(pRuntimeInfo))
     
    'CLSID_CorRuntimeHost "CB2F6723-AB3A-11D2-9C40-00C04FA30A3E"
    Dim GUID_CRH&(0 To 7)
    GUID_CRH(0) = &HCB2F6723: GUID_CRH(1) = &H11D2AB3A: GUID_CRH(2) = &HC000409C: GUID_CRH(3) = &H3E0AA34F
    
    'CLSID IID_ICorRuntimeHost "CB2F6722-AB3A-11D2-9C40-00C04FA30A3E"
    Dim GUID_ICRH&(0 To 7)
    GUID_ICRH(0) = &HCB2F6722: GUID_ICRH(1) = &H11D2AB3A: GUID_ICRH(2) = &HC000409C: GUID_ICRH(3) = &H3E0AA34F

    ' ICLRRuntimeInfo::GetInterface(REFCLSID, REFIID, void**) [vftable index = 9]
    Call DispCallWrapper(pRuntimeInfo, 9, VarPtr(GUID_CRH(0)), VarPtr(GUID_ICRH(0)), VarPtr(pCorRuntimeHost))

    Set Init = pCorRuntimeHost

End Function

Sub Stg1()
    
    'Accepts both CLR 2.0 and CLR 4.0, Supports assemblies built with .NET 2.0 -> .NET 4.x
    ThisDocument.VBProject.References.AddFromFile "C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscoree.tlb"
    ThisDocument.VBProject.References.AddFromFile "C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
    
    Call Stg2

End Sub

Sub Stg2()

    Dim CLRVersion As String
    CLRVersion = "v4.0.30319"
        
    Dim ICRHost As mscoree.CorRuntimeHost
    Dim IDomain As mscorlib.AppDomain

    Set ICRHost = Init(CLRVersion)
    Call ICRHost.Start
    Call ICRHost.GetDefaultDomain(IDomain)
    
    Dim path As String
    path = "C:\HelloWorld.NET4.x.exe"

    IDomain.ExecuteAssembly_2 (path)
    
End Sub

