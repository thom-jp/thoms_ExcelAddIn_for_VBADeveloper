Attribute VB_Name = "fnCommon"
Function OpenFolderDialog()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Not .Show Then Exit Function
        OpenFolderDialog = .SelectedItems(1)
    End With
End Function
