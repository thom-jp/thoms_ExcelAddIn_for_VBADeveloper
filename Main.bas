Attribute VB_Name = "Main"
'çÏé“: thom
'ÉuÉçÉO: http://thom.hateblo.jp/
'Github: https://github.com/thom-jp/thoms_ExcelAddIn_for_VBADeveloper

Private Menu As MenuCreator

Public Sub Auto_Open()
    Dim vbc As VBComponent: Set vbc _
    = ThisWorkbook.VBProject.VBComponents("MenuMacros")
    
    Set Menu = New MenuCreator
    Menu.Init "thom", "thom(&M)", vbc
    
    Dim arr() As Instruction: arr = GetInstructions(vbc.CodeModule)
    Dim i As Long
    
    For i = 0 To UBound(arr)
        Menu.AddSubMenu arr(i).name, arr(i).shortcut
    Next
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Menu.RemoveMenu
    On Error GoTo 0
    Set Menu = Nothing
End Sub

Private Function GetInstructions(cmod As CodeModule) As Instruction()
    Dim psl As Long, pbl As String
    Dim ret() As Instruction: ReDim ret(0)
    Dim i As Long
    For i = 1 To cmod.CountOfLines
        Dim pname As String
        pname = cmod.ProcOfLine(i, vbext_pk_Proc)
        If pname <> "" Then
            psl = cmod.ProcBodyLine(pname, vbext_pk_Proc)
            If i = psl Then
                pbl = cmod.Lines(psl, 1)
                Set ret(UBound(ret)) = New Instruction
                
                On Error Resume Next
                    ret(UBound(ret)).shortcut = Split(pbl, "'")(1)
                On Error GoTo 0
                
                ret(UBound(ret)).name = pname
                ReDim Preserve ret(UBound(ret) + 1)
            End If
        End If
    Next
    ReDim Preserve ret(UBound(ret) - 1)
    GetInstructions = ret
End Function

