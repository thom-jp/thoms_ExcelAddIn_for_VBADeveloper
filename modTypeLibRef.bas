Attribute VB_Name = "modTypeLibRef"
Option Explicit
Sub 初期設定()
    shTypeLib.Cells.Clear
    Call Ref
    Call SheetSort
End Sub

Sub Ref()
    Dim レジストリ As Object: Set レジストリ = _
        CreateObject("WbemScripting.SWbemLocator") _
            .ConnectServer(, "root\default") _
                .Get("StdRegProv")

    Const HKCR = &H80000000
    Dim TypeLibの子, TypeLibの孫, 子, 孫, 値, 値2
    Dim arr() As String
    ReDim arr(1 To 2, 1 To 1)
    Dim i As Long
    i = 1
    レジストリ.EnumKey HKCR, "TypeLib", TypeLibの子
    For Each 子 In TypeLibの子
        レジストリ.EnumKey HKCR, "TypeLib\" & 子, TypeLibの孫
        If Not IsNull(TypeLibの孫) Then
            For Each 孫 In TypeLibの孫
                レジストリ.GetStringValue HKCR, "TypeLib\" & 子 & "\" & 孫 & "\0\win32", , 値
                レジストリ.GetStringValue HKCR, "TypeLib\" & 子 & "\" & 孫, , 値2
                If (Not IsNull(値2)) And (Not IsNull(値)) Then
                    shTypeLib.Cells(i, 1) = 値2
                    shTypeLib.Cells(i, 2) = 値
                    i = i + 1
                End If
            Next
        End If
    Next
End Sub

Sub SheetSort()
    With shTypeLib.Sort
        With .SortFields
            .Clear
            .Add Key:=shTypeLib.Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End With
        .SetRange shTypeLib.Cells(1, 1).CurrentRegion
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

