VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cPopupMenuItem.cls                    7/15/05
'
'           PURPOSE:
'               Represent a popup menu item.
'
'==================================================================================================

Option Explicit

Private mpItem As Long
Private moOwner As ucPopupMenus

Friend Sub fInit(ByVal oOwner As ucPopupMenus, ByVal pItem As Long)
    Set moOwner = oOwner
    mpItem = pItem
    moOwner.fItem_AddRef mpItem
End Sub

Private Sub Class_Initialize()
    PopupMenus_GetDC
End Sub

Private Sub Class_Terminate()
    moOwner.fItem_Release mpItem
    PopupMenus_ReleaseDC
End Sub

Friend Property Get fpItem() As Long
    fpItem = mpItem
End Property


Public Property Get Caption() As String
    Caption = moOwner.fItem_Caption(mpItem)
End Property
Public Property Let Caption(ByRef sNew As String)
    moOwner.fItem_Caption(mpItem) = sNew
End Property

Public Property Get Key() As String
    Key = moOwner.fItem_Key(mpItem)
End Property
Public Property Let Key(ByRef sNew As String)
    moOwner.fItem_Key(mpItem) = sNew
End Property

Public Property Get HelpString() As String
    HelpString = moOwner.fItem_HelpString(mpItem)
End Property
Public Property Let HelpString(ByRef sNew As String)
    moOwner.fItem_HelpString(mpItem) = sNew
End Property

Public Property Get ShortcutDisplay() As String
    ShortcutDisplay = moOwner.fItem_ShortcutDisplay(mpItem)
End Property
Public Property Let ShortcutDisplay(ByRef sNew As String)
    moOwner.fItem_ShortcutDisplay(mpItem) = sNew
End Property

Public Property Get Checked() As Boolean
    Checked = moOwner.fItem_Checked(mpItem)
End Property
Public Property Let Checked(ByVal bNew As Boolean)
    moOwner.fItem_Checked(mpItem) = bNew
End Property

Public Property Get Default() As Boolean
    Default = moOwner.fItem_Default(mpItem)
End Property
Public Property Let Default(ByVal bNew As Boolean)
    moOwner.fItem_Default(mpItem) = bNew
End Property

Public Property Get Enabled() As Boolean
    Enabled = moOwner.fItem_Enabled(mpItem)
End Property
Public Property Let Enabled(ByVal bNew As Boolean)
    moOwner.fItem_Enabled(mpItem) = bNew
End Property

Public Property Get Infrequent() As Boolean
    Infrequent = moOwner.fItem_Infrequent(mpItem)
End Property
Public Property Let Infrequent(ByVal bNew As Boolean)
    moOwner.fItem_Infrequent(mpItem) = bNew
End Property

Public Property Get ItemData() As Long
    ItemData = moOwner.fItem_ItemData(mpItem)
End Property
Public Property Let ItemData(ByVal iNew As Long)
    moOwner.fItem_ItemData(mpItem) = iNew
End Property

Public Property Get BreakLine() As Boolean
    BreakLine = moOwner.fItem_BreakLine(mpItem)
End Property
Public Property Let BreakLine(ByVal bNew As Boolean)
    moOwner.fItem_BreakLine(mpItem) = bNew
End Property

Public Property Get RadioChecked() As Boolean
    RadioChecked = moOwner.fItem_RadioChecked(mpItem)
End Property
Public Property Let RadioChecked(ByVal bNew As Boolean)
    moOwner.fItem_RadioChecked(mpItem) = bNew
End Property

Public Property Get RedisplayOnClick() As Boolean
    RedisplayOnClick = moOwner.fItem_RedisplayOnClick(mpItem)
End Property
Public Property Let RedisplayOnClick(ByVal bNew As Boolean)
    moOwner.fItem_RedisplayOnClick(mpItem) = bNew
End Property

Public Property Get Separator() As Boolean
    Separator = moOwner.fItem_Separator(mpItem)
End Property
Public Property Let Separator(ByVal bNew As Boolean)
    moOwner.fItem_Separator(mpItem) = bNew
End Property

Public Property Get Visible() As Boolean
    Visible = moOwner.fItem_Visible(mpItem)
End Property
Public Property Let Visible(ByVal bNew As Boolean)
    moOwner.fItem_Visible(mpItem) = bNew
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fItem_IconIndex(mpItem)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fItem_IconIndex(mpItem) = iNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fItem_Index(mpItem)
End Property

Public Property Get ShortcutShiftMask() As evbComCtlKeyboardState
    ShortcutShiftMask = moOwner.fItem_ShortcutShiftMask(mpItem)
End Property
Public Property Let ShortcutShiftMask(ByVal iNew As evbComCtlKeyboardState)
    moOwner.fItem_ShortcutShiftMask(mpItem) = iNew
End Property

Public Property Get ShortcutKey() As Long
    ShortcutKey = moOwner.fItem_ShortcutKey(mpItem)
End Property
Public Property Let ShortcutKey(ByVal iNew As Long)
    moOwner.fItem_ShortcutKey(mpItem) = iNew
End Property

Public Property Get Style() As ePopupItemStyle
    Style = moOwner.fItem_Style(mpItem)
End Property
Public Property Let Style(ByVal iNew As ePopupItemStyle)
    moOwner.fItem_Style(mpItem) = iNew
End Property

Public Sub SetStyle(ByVal iStyleOr As ePopupItemStyle, ByVal iStyleAndNot As ePopupItemStyle)
    moOwner.fItem_SetStyle mpItem, iStyleOr, iStyleAndNot
End Sub

Public Property Get SubMenu() As cPopupMenu
    Set SubMenu = moOwner.fItem_SubMenu(mpItem)
End Property

Public Property Get Parent() As cPopupMenu
    Set Parent = moOwner
End Property
