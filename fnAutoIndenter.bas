Attribute VB_Name = "fnAutoIndenter"
Public Enum CodeKind
    Normal = 1
    BlockIn = 2
    BlockOut = 3
    ProdIn = 4
    ProdOut = 5
    BlockOutIn = 6
    BlankLine = 7
    BlockInIn = 8
    BlockOutOut = 9
End Enum

Function ConvertString(ByVal x As String) As String
    ConvertString = x
    If InStr(1, x, """") > 0 Then
        If InStr(1, x, "'") > 0 Then
            If InStr(1, x, """") < InStr(1, x, "'") Then
                ConvertString = InnerConvertString(x)
            End If
        Else
            ConvertString = InnerConvertString(x)
        End If
    End If
End Function

Function GetCodeKind(x2 As String) As CodeKind
    Dim ret As CodeKind
    Select Case True
        Case Left(Trim(x2), 1) = "'": ret = Normal
        Case InStr(1, x2, "End Sub") > 0: ret = ProdOut
        Case InStr(1, x2, "Exit Sub") > 0: ret = Normal
        Case InStr(1, x2, "Go Sub") > 0: ret = Normal
        Case InStr(1, x2, "Sub ") > 0: ret = ProdIn
        Case InStr(1, x2, "End Function") > 0: ret = ProdOut
        Case InStr(1, x2, "Exit Function") > 0: ret = Normal
        Case InStr(1, x2, "Function ") > 0: ret = ProdIn
        Case InStr(1, x2, "End Property") > 0: ret = ProdOut
        Case InStr(1, x2, "Exit Property") > 0: ret = Normal
        Case InStr(1, x2, "Property ") > 0: ret = ProdIn
        Case InStr(1, x2, "End Select") > 0: ret = BlockOutOut
        Case InStr(1, x2, "Select Case") > 0: ret = BlockInIn
        Case InStr(1, x2, "Case ") > 0: ret = BlockOutIn
        Case InStr(1, x2, "For ") > 0: ret = BlockIn
        Case InStr(1, x2, "Next") > 0: ret = BlockOut
        Case InStr(1, x2, "Exit Do") > 0: ret = Normal
        Case InStr(1, x2, "Do ") > 0: ret = BlockIn
        Case Trim(x2) = "Do": ret = BlockIn
        Case InStr(1, x2, "End With") > 0: ret = BlockOut
        Case InStr(1, x2, "With") > 0: ret = BlockIn
        Case InStr(1, x2, "End If") > 0: ret = BlockOut
        Case InStr(1, x2, "Else") > 0: ret = BlockOutIn
        Case InStr(1, x2, "Then ") > 0: ret = Normal
        Case InStr(1, x2, "If ") > 0: ret = BlockIn
        Case InStr(1, x2, "Loop ") > 0: ret = BlockOut
        Case Trim(x2) = "Loop": ret = BlockOut
        Case Trim(x2) <> "": ret = BlankLine
    End Select
    GetCodeKind = ret
End Function

Function Indent(code As String)
    Dim Lines() As String: Lines = Split(code, vbNewLine)
    Dim SB As New StringBuilder
    For Each x In Lines
        Dim x2 As String: x2 = ConvertString(x)
        Select Case GetCodeKind(x2)
        Case CodeKind.BlankLine
            SB.AppendLine String(t, vbTab) & Trim(x)
        Case CodeKind.BlockIn
            SB.AppendLine String(t, vbTab) & Trim(x)
            t = t + 1
        Case CodeKind.BlockOut
            t = t - 1
            SB.AppendLine String(t, vbTab) & Trim(x)
        Case CodeKind.BlockOutIn
            t = t - 1
            SB.AppendLine String(t, vbTab) & Trim(x)
            t = t + 1
        Case CodeKind.Normal
            SB.AppendLine String(t, vbTab) & Trim(x)
        Case CodeKind.ProdIn
            t = 0
            SB.AppendLine String(t, vbTab) & Trim(x)
            t = 1
        Case CodeKind.ProdOut
            t = 0
            SB.AppendLine String(t, vbTab) & Trim(x)
        Case CodeKind.BlockInIn
            SB.AppendLine String(t, vbTab) & Trim(x)
            t = t + 2
        Case CodeKind.BlockOutOut
            t = t - 2
            SB.AppendLine String(t, vbTab) & Trim(x)
        End Select
    Next
    Indent = SB.Value
End Function

Private Function InnerConvertString(x As String) As String
    Dim newstr As String
    Dim IsString As Boolean
    For i = 1 To Len(x)
        IsString = (Mid(x, i, 1) = """") Xor IsString
        If Not IsString Then
            newstr = newstr & Mid(x, i, 1)
        End If
    Next
    InnerConvertString = newstr
End Function

