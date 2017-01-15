VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TypeLibRef 
   Caption         =   "参照設定（改）"
   ClientHeight    =   6624
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   11700
   OleObjectBlob   =   "TypeLibRef.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "TypeLibRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    TextBox1.Text = vbNullString
    Call ListReset
    Dim w As Workbook
    For Each w In Workbooks
        If Not w Is ThisWorkbook Then
            ComboBox1.AddItem w.name
        End If
    Next
End Sub

Private Sub ListReset()
    ListBox1.Clear
    Dim arr() As Variant
    arr = ReadData
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Me.ListBox1.AddItem arr(i, 1)
        Me.ListBox1.List(ListBox1.ListCount - 1, 1) = arr(i, 2)
    Next
End Sub

Function ReadData() As Variant
    If shTypeLib.Cells(1, 1) = "" Then
        MsgBox "データがありません。", vbExclamation, "確認"
        MsgBox "初期設定を実施します。", vbInformation, "確認"
        初期設定
    End If
    ReadData = shTypeLib.Cells(1, 1).CurrentRegion.Value
End Function

Private Sub cmdクリア_Click()
    UserForm_Initialize
End Sub

Private Sub cmd検索_Click()
    Call ListReset
    Dim i As Long
    Do While i < ListBox1.ListCount
        If InStr(1, ListBox1.List(i, 0), TextBox1.Text, vbTextCompare) = 0 Then
            ListBox1.RemoveItem (i)
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub cmd絞込み_Click()
    Dim i As Long
    Do While i < ListBox1.ListCount
        If InStr(1, ListBox1.List(i, 0), TextBox1.Text, vbTextCompare) = 0 Then
            ListBox1.RemoveItem (i)
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With ListBox1
        ListBox2.AddItem .List(.ListIndex, 0)
        ListBox2.List(ListBox2.ListCount - 1, 1) = .List(.ListIndex, 1)
    End With
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next '未選択でWクリックした場合のエラーを無視
        ListBox2.RemoveItem ListBox2.ListIndex
    On Error GoTo 0
End Sub

Private Sub cmdOK_Click()
    With ListBox2
        If .ListCount > 0 Then
            If ComboBox1.Text <> "" Then
                Dim i As Long, cnt As Long
                For i = 0 To .ListCount - 1
                    On Error Resume Next '参照不可エラーは面倒なのでスキップ
                    Workbooks(ComboBox1.Text).VBProject.References.AddFromFile .List(i, 1)
                    If Err.Number = 0 Then cnt = cnt + 1
                    On Error GoTo 0
                Next
                MsgBox "ワークブック「" & Workbooks(ComboBox1.Text).name & "」に" _
                    & cnt & "件の参照を追加しました。", vbInformation, "完了"
                Unload Me
            Else
                MsgBox "左上のコンボボックスで対象のブックを選択してください。", vbInformation, "エラー"
            End If
        End If
    End With
End Sub

Private Sub cmd初期設定_Click()
    MsgBox "初期設定では、レジストリからタイプライブラリの情報を読み込みます。", vbInformation, "初期設定について"
    MsgBox "初回起動時や、ソフトウェアのインストールを行った場合に実施してください。", vbInformation, "初期設定について"
    MsgBox "この操作は数秒～数十秒かかります。", vbInformation, "初期設定について"
    If vbYes = MsgBox("続行しますか。", vbInformation + vbYesNo, "確認") Then
        Call 初期設定
        Call UserForm_Initialize
        MsgBox "完了しました。", vbInformation, "完了"
    Else
        MsgBox "キャンセルしました。", vbInformation, "中止"
    End If
End Sub

