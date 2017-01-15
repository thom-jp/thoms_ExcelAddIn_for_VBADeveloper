VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Value As String
Public IndentLevel As Long
Public AppendTime As Long
Sub Append(str As String)
    Value = Value & str
End Sub
Sub AppendLine(Optional str = "", Optional trimspace = True)
    Value = Value & vbNewLine
    Value = Value & String(IndentLevel, vbab)
    Append CStr(str)
    AppendTime = AppendTime + 1
End Sub

Sub Indent()
    IndentLevel = IndentLevel + 1
End Sub

Sub UnIndent()
    If IndentLevel > 0 Then
        IndentLevel = IndentLevel - 1
    Else
        Debug.Print Value
        Err.Raise vbObjectError, "StringBuilder", "これ以上インデントを減らすことはできません。"
    End If
End Sub
