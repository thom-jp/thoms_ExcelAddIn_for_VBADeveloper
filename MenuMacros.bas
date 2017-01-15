Attribute VB_Name = "MenuMacros"
Sub 参照設定改() 'R
    TypeLibRef.Show
End Sub

Sub カレントプロジェクトをエクスポート() 'E
    MsgBox "出力先を選択してください。", vbInformation
    Dim savepath: savepath = OpenFolderDialog
    If savepath = "" Then GoTo Cancel_Exit
    
    If CreateObject("Scripting.FileSystemObject").GetFolder(savepath).Files.Count > 0 Then
        If vbYes <> MsgBox( _
            "フォルダ内にファイルが存在します。" _
            & "同名のファイルは上書きされますがこのまま続行しますか？" _
            , vbExclamation + vbYesNo, "確認") _
        Then
            GoTo Cancel_Exit
        End If
    End If
        
    Dim vbc As VBComponent
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbc.Type
            Case vbext_ct_StdModule
                ext = ".bas"
            Case vbext_ct_ClassModule
                ext = ".cls"
            Case vbext_ct_Document
                ext = ".obj.cls"
            Case vbext_ct_MSForm
                ext = ".frm"
            Case Else
                ext = ".unknown"
        End Select
        vbc.Export savepath & "\" & vbc.name & ext
    Next
    MsgBox "エクスポート完了しました。", vbInformation, "完了"
Exit Sub
Cancel_Exit:
    MsgBox "キャンセルしました。", vbInformation, "キャンセル"
End Sub

Sub カラーパレットを表示() 'O
    Application.Visible = False
    With Application.VBE.ActiveCodePane
        ColorPicker.Show
        .Show
    End With
    Application.Visible = True
End Sub

Sub 非アクティブなコードペインを閉じる() 'C
    Dim C As VBIDE.CodePane
    With ThisWorkbook.VBProject.VBE
        For Each C In .CodePanes
            If Not .ActiveCodePane Is C Then C.Window.Close
        Next
    End With
End Sub

Sub 選択されたモジュールのインデントを調整() 'I
    'プロシージャ上のコメントが消えてしまうので要改善
    Dim sl As Long, sc As Long, el As Long, ec As Long
    Application.VBE.ActiveCodePane.GetSelection sl, sc, el, ec
    Dim SB As StringBuilder: Set SB = New StringBuilder
    With Application.VBE.ActiveCodePane.CodeModule
        Dim startProc As String: startProc = .ProcOfLine(sl, vbext_pk_Proc)
        Dim endProc As String: endProc = .ProcOfLine(el, vbext_pk_Proc)
        If startProc <> "" And endProc <> "" Then
            If startProc = endProc Then
                Dim pKind As vbext_ProcKind
                pKind = CurrentProcKind(startProc, sl, el)
                If pKind <> -1 Then
                    pcount = .ProcCountLines(startProc, pKind)
                    pbodystart = .ProcBodyLine(startProc, pKind)
                    pstart = .ProcStartLine(startProc, pKind)
                    prealcount = pcount - (pbodystart - pstart)
                    SB.AppendLine .Lines(pbodystart, prealcount)
                    .DeleteLines pstart, pcount
                    .InsertLines pstart, Indent(SB.Value)
                Else
                    MsgBox "プロシージャを特定できません。", vbExclamation, "エラー"
                End If
            Else
                MsgBox "プロシージャを特定できません。", vbExclamation, "エラー"
            End If
        Else
            MsgBox "Declarations 領域では実行できません。" & vbNewLine & _
                "プロシージャ上で実行してください。", vbExclamation, "エラー"
        End If
    End With
End Sub

Sub メニュー更新() 'U
    Main.Auto_Open
End Sub
