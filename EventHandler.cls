VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents MenuEvent As VBIDE.CommandBarEvents
Attribute MenuEvent.VB_VarHelpID = -1

Public Property Get Self() As Object
    Set Self = Me
End Property

Private Sub MenuEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Application.Run CommandBarControl.OnAction
    handled = True
    CancelDefault = True
End Sub
