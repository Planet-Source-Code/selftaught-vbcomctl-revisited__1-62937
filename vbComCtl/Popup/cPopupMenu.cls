VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cPopupMenu.cls                    7/15/05
'
'           PURPOSE:
'               Represent a single level of a popupmenu and a collection of its subitems.
'
'==================================================================================================
Option Explicit

Implements iEnumerable

Private mpMenu As Long
Private moOwner As ucPopupMenus

Private Sub Class_Initialize()
    PopupMenus_GetDC
End Sub

Private Sub Class_Terminate()
    moOwner.fMenu_Release mpMenu
    PopupMenus_ReleaseDC
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
'
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fMenu_Enum_GetNextItem mpMenu, tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    moOwner.fMenu_Enum_Reset mpMenu, tEnum
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    moOwner.fMenu_Enum_Skip mpMenu, tEnum, iSkipCount, bSkippedAll
End Sub


Friend Sub fInit(ByVal oOwner As ucPopupMenus, ByVal pMenu As Long)
    Set moOwner = oOwner
    mpMenu = pMenu
    moOwner.fMenu_AddRef mpMenu
End Sub

Public Property Get SidebarExists() As Boolean
    SidebarExists = moOwner.fMenu_SidebarExists(mpMenu)
End Property

Public Sub SetSidebar(ByVal oNew As Object)
    moOwner.fMenu_SetSidebar mpMenu, oNew
End Sub

Public Property Get Parent() As cPopupMenu
    Set Parent = moOwner.fMenu_Parent(mpMenu)
End Property

Public Property Get Root() As cPopupMenu
    Set Root = moOwner.fMenu_Root(mpMenu)
End Property

Public Property Get ShowCheckAndIcon() As Boolean
    ShowCheckAndIcon = moOwner.fMenu_ShowCheckAndIcon(mpMenu)
End Property
Public Property Let ShowCheckAndIcon(ByVal bNew As Boolean)
    moOwner.fMenu_ShowCheckAndIcon(mpMenu) = bNew
End Property

Public Function Add( _
   Optional ByRef sCaption As String, _
   Optional ByRef sHelpString As String, _
   Optional ByRef sKey As String, _
   Optional ByVal iIconIndex As Long = NegOneL, _
   Optional ByVal iStyle As ePopupItemStyle, _
   Optional ByVal iShortcutKey As Integer, _
   Optional ByVal iShortcutMask As evbComCtlKeyboardState, _
   Optional ByVal iItemData As Long, _
   Optional ByVal vItemInsertBefore As Variant) _
                As cPopupMenuItem
    Set Add = moOwner.fMenu_Add(mpMenu, sCaption, sHelpString, sKey, iIconIndex, iStyle, iShortcutKey, iShortcutMask, iItemData, vItemInsertBefore)
End Function

Public Sub Clear()
    moOwner.fMenu_Clear mpMenu
End Sub

Public Sub Remove(ByVal vItem As Variant)
    moOwner.fMenu_Remove mpMenu, vItem
End Sub

Public Property Get Count() As Long
    Count = moOwner.fMenu_Count(mpMenu)
End Property

Public Property Get Exists(ByVal vItem As Variant) As Boolean
    Exists = moOwner.fMenu_Exists(mpMenu, vItem)
End Property

Public Property Get Item(ByVal vItem As Variant) As cPopupMenuItem
    Set Item = moOwner.fMenu_Item(mpMenu, vItem)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Dim loEnum As pcEnumeration
    Set loEnum = New pcEnumeration
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fMenu_Control(mpMenu), ZeroL)
End Property


Public Function Show( _
            ByVal iFlags As ePopupShowFlag, _
            ByVal fLeft As Single, _
            ByVal fTop As Single, _
            ByVal fExcludeLeft As Single, _
            ByVal fExcludeTop As Single, _
            ByVal fExcludeWidth As Single, _
            ByVal fExcludeHeight As Single) _
                As cPopupMenuItem
    Set Show = moOwner.fMenu_Show(mpMenu, iFlags, fLeft, fTop, fExcludeLeft, fExcludeTop, fExcludeWidth, fExcludeHeight)
End Function

Public Function ShowAtCursor( _
            ByVal iFlags As ePopupShowFlag) _
                As cPopupMenuItem
    Set ShowAtCursor = moOwner.fMenu_ShowAtCursor(mpMenu, iFlags)
End Function

Public Function ShowAtControl( _
            ByVal iFlags As ePopupShowFlag, _
            ByVal oControl As Object, _
   Optional ByVal bShowOnRight As Boolean) _
                As cPopupMenuItem
    Set ShowAtControl = moOwner.fMenu_ShowAtControl(mpMenu, iFlags, oControl, bShowOnRight)
End Function
