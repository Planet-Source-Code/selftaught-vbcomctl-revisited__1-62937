VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cPanels.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of status bar panels.
'
'==================================================================================================

Option Explicit


Implements iEnumerable

Private moOwner As ucStatusBar

Friend Sub fInit(ByVal oOwner As ucStatusBar)
    Set moOwner = oOwner
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
'
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fPanels_Enum_GetNextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iControl = moOwner.fPanels_Control()
    tEnum.iIndex = NegOneL
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    moOwner.fPanels_Enum_Skip tEnum, iSkipCount, bSkippedAll
End Sub

Public Function Add( _
        Optional ByRef sText As String, _
        Optional ByRef sKey As String, _
        Optional ByRef sToolTipText As String, _
        Optional ByVal iStyle As eStatusBarPanelStyle, _
        Optional ByVal iBorder As eStatusBarPanelBorder, _
        Optional ByVal iIconIndex As Long = NegOneL, _
        Optional ByVal fMinWidth As Single, _
        Optional ByVal fIdealWidth As Single, _
        Optional ByVal bSpring As Boolean, _
        Optional ByVal bFit As Boolean, _
        Optional ByVal vPanelInsertBefore As Variant) _
            As cPanel
    Set Add = moOwner.fPanels_Add(sText, sKey, sToolTipText, iStyle, iBorder, iIconIndex, fMinWidth, fIdealWidth, bSpring, bFit, vPanelInsertBefore)
End Function

Public Sub Remove(ByVal vPanel As Variant)
    moOwner.fPanels_Remove vPanel
End Sub

Public Sub Clear()
    moOwner.fPanels_Clear
End Sub

Public Property Get Item(ByVal vPanel As Variant) As cPanel
Attribute Item.VB_UserMemId = 0
    Set Item = moOwner.fPanels_Item(vPanel)
End Property

Public Property Get Exists(ByVal vPanel As Variant)
    Exists = moOwner.fPanels_Exists(vPanel)
End Property

Public Property Get Count() As Long
    Count = moOwner.fPanels_Count()
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Dim loEnum As pcEnumeration
    Set loEnum = New pcEnumeration
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fPanels_Control())
End Property
