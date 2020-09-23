VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cButtons.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of toolbar buttons.
'
'==================================================================================================

Option Explicit

Implements iEnumerable

Private moOwner As ucToolbar

Friend Sub fInit(ByVal oOwner As ucToolbar)
    Set moOwner = oOwner
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
'
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fButtons_GetNextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = NegOneL
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
    bSkippedAll = tEnum.iIndex <= moOwner.fButtons_Count()
End Sub


Public Function Add( _
            Optional ByRef sKey As String, _
            Optional ByRef sText As String, _
            Optional ByVal iStyle As eToolbarButtonStyle = tbarButtonNormal, _
            Optional ByVal iIconIndex As Long = NegOneL, _
            Optional ByRef sToolTipText As String, _
            Optional ByVal bAutosize As Boolean, _
            Optional ByVal iItemData As Long, _
            Optional ByVal bEnabled As Boolean = True, _
            Optional ByVal bVisible As Boolean = True, _
            Optional ByVal vButtonBefore As Variant) _
                As cButton
    Set Add = moOwner.fButtons_Add(sKey, sText, iStyle, iIconIndex, sToolTipText, bAutosize, iItemData, bEnabled, bVisible, vButtonBefore)
End Function

Public Sub Remove(ByVal vButton As Variant)
    moOwner.fButtons_Remove vButton
End Sub

Public Property Get Exists(ByVal vButton As Variant) As Boolean
    Exists = moOwner.fButtons_Exists(vButton)
End Property

Public Sub Clear()
    moOwner.fButtons_Clear
End Sub

Public Property Get Count() As Long
    Count = moOwner.fButtons_Count()
End Property

Public Property Get Item(ByVal vButton As Variant) As cButton
Attribute Item.VB_UserMemId = 0
    Set Item = moOwner.fButtons_Item(vButton)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Dim loEnum As pcEnumeration
    Set loEnum = New pcEnumeration
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fButtons_Control)
End Property
