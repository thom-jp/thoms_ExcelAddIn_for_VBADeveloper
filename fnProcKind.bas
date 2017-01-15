Attribute VB_Name = "fnProcKind"
Function CurrentProcKind(procName As String, sl, el) As vbext_ProcKind
    Dim CM As CodeModule: Set CM _
        = Application.VBE.ActiveCodePane.CodeModule
    Dim ret As vbext_ProcKind
    Dim E As Long
    For ret = 0 To 3
        On Error Resume Next
        procBeginLine = CM.ProcStartLine(procName, ret)
        E = Err.Number
        On Error GoTo 0
        If E = 0 Then
            procEndLine = procBeginLine + CM.ProcCountLines(procName, ret)
            If sl >= procBeginLine And el <= procEndLine Then
                GoTo Quit
            End If
        End If
    Next
    ret = -1
Quit:
    CurrentProcKind = ret
End Function
