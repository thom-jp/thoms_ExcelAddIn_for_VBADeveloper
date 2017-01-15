VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MenuTag As String
Private RootMenu As CommandBarPopup
Private EventHandlers As Collection
Private MenuMacroComponentFullName As String

Public Sub Init(tag As String, rootCaption As String, vbc As VBComponent)
    MenuTag = tag
    Call Me.RemoveMenu
    Set EventHandlers = New Collection
    Dim VBEMenuBar As CommandBar: Set VBEMenuBar _
        = Application.VBE.CommandBars(1)
    
    With New FileSystemObject
        MenuMacroComponentFullName _
            = "'" & .GetFileName(vbc.Collection.Parent.Filename) & "'!" & vbc.name
    End With
    
    Set RootMenu = VBEMenuBar.Controls.Add(Type:=msoControlPopup)
    RootMenu.Caption = rootCaption
    RootMenu.tag = MenuTag
End Sub

Public Sub AddSubMenu(procName As String, shortcut As String)
    Dim SubMenu As CommandBarControl: Set SubMenu _
        = RootMenu.Controls.Add
    
    With SubMenu
        .Caption = procName & "(&" & shortcut & ")"
        .BeginGroup = False
        .OnAction = MenuMacroComponentFullName & "." & procName
    End With
    
    With New EventHandler
        Set .MenuEvent = Application.VBE.Events.CommandBarEvents(SubMenu)
        EventHandlers.Add .Self
    End With
End Sub

Public Sub RemoveMenu()
    'RootMenu.Deleteとする代わりに、わざわざMenuTagで検索して消すのは、
    '前回の異常終了で残ってしまったメニューも片づけるため。
    Dim MyMenu As CommandBarControl: Set MyMenu _
        = Application.VBE.CommandBars.FindControl(tag:=MenuTag)
    Do Until MyMenu Is Nothing
        MyMenu.Delete
        Set MyMenu = Application.VBE.CommandBars.FindControl(tag:=MenuTag)
    Loop
    Set EventHandlers = Nothing
End Sub
