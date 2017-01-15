Attribute VB_Name = "fnInsertCode"
Sub InsertCode(code As String)
    Dim sl As Long, sc As Long, el As Long, ec As Long
    Application.VBE.ActiveCodePane.GetSelection sl, sc, el, ec
    L = Application.VBE.ActiveCodePane.CodeModule.Lines(sl, 1)
    L2 = Left(L, sc - 1) & code & Mid(L, sc)
    Application.VBE.ActiveCodePane.CodeModule.ReplaceLine sl, L2
    Application.VBE.ActiveCodePane.SetSelection sl, sc + Len(code) + 1, sl, sc + Len(code) + 1
End Sub
