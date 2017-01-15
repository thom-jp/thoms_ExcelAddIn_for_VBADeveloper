VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorPicker 
   Caption         =   "UserForm1"
   ClientHeight    =   3480
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   4788
   OleObjectBlob   =   "ColorPicker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const LabelSizeX = 20
Private Const LabelSizeY = 15
Private ColorLabels As Collection
Private Sub UserForm_Initialize()
    Set ColorLabels = New Collection
    hoge
    Me.Height = (LabelSizeY) * 10 - 5
    Me.Width = LabelSizeX * 9 - 4
End Sub

Sub WriteLabel(x, y, color As Long)
    With New EventControl
        Set .AsLabel = Me.Controls.Add("Forms.Label.1")
        With .AsControl
            .Height = LabelSizeY
            .Width = LabelSizeX
            .Left = x * (LabelSizeX - 1)
            .Top = y * (LabelSizeY - 1)
        End With
        With .AsLabel
            .BackColor = color
            .BorderStyle = fmBorderStyleSingle
        End With
        ColorLabels.Add .Self
    End With
End Sub

Sub hoge()
    colorBase = 51
    ColorPallets = Array( _
        Array(5, 0, 0), Array(5, 3, 0), Array(3, 5, 0), _
        Array(0, 5, 0), Array(0, 5, 3), Array(0, 3, 5), _
        Array(0, 0, 5), Array(3, 0, 5), Array(5, 0, 3))

    Dim R, G, B
    For i = 0 To 8
        For j = 0 To 4
            cp = ColorPallets(i)
            R = (cp(0) - j) * colorBase
            G = (cp(1) - j) * colorBase
            B = (cp(2) - j) * colorBase
            If R < 0 Then R = 0
            If G < 0 Then G = 0
            If B < 0 Then B = 0
            WriteLabel 4 - j, i, RGB(R, G, B)
        Next
    Next
    For i = 0 To 8
        For j = 1 To 4
            cp = ColorPallets(i)
            R = (cp(0) + j) * colorBase
            G = (cp(1) + j) * colorBase
            B = (cp(2) + j) * colorBase
            If R > 255 Then R = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            WriteLabel j + 4, i, RGB(R, G, B)
        Next
    Next
End Sub
