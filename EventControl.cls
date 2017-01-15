VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents AsLabel As MSForms.Label
Attribute AsLabel.VB_VarHelpID = -1

Property Get AsControl() As MSForms.Control
    Set AsControl = AsLabel
End Property

Property Get Red()
    Red = &HFF& And AsLabel.BackColor
End Property
Property Get Green()
    Green = (&HFF00& And AsLabel.BackColor) \ 256
End Property
Property Get Blue()
    Blue = (&HFF0000 And AsLabel.BackColor) \ 256 \ 256
End Property

Property Get RGBFuncAsString() As String
    RGBFuncAsString = "RGB(" & Red & ", " & Green & ", " & Blue & ")"
End Property

Private Sub AsLabel_Click()
    ColorPicker.Hide
    Application.VBE.MainWindow.Visible = True
    InsertCode RGBFuncAsString
End Sub

Property Get Self() As Object
    Set Self = Me
End Property
