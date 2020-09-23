VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cBands.cls                    12/15/04
'
'           PURPOSE:
'               Represent a collection of rebar bands.
'
'==================================================================================================

Option Explicit

Implements iEnumerable

Private moOwner As ucRebar


Friend Sub fInit(ByVal oOwner As ucRebar)
    Set moOwner = oOwner
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fBands_Enum_GetNextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = NegOneL
    tEnum.iControl = moOwner.fBands_Enum_Control()
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    tEnum.iIndex = tEnum.iIndex + OneL
    bSkippedAll = tEnum.iIndex <= moOwner.fBands_Count()
End Sub

Public Property Get NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Dim loEnum As pcEnumeration
    Set loEnum = New pcEnumeration
    Set NewEnum = loEnum.GetEnum(Me, moOwner.fBands_Enum_Control())
End Property


Public Function Add( _
               ByVal oBandChild As Object, _
      Optional ByRef sKey As String, _
      Optional ByRef sText As String, _
      Optional ByVal bUseChevron As Boolean, _
      Optional ByVal bBreakLine As Boolean, _
      Optional ByVal bGripper As Boolean = True, _
      Optional ByVal bVisible As Boolean = True, _
      Optional ByVal bFixedSize As Boolean, _
      Optional ByVal iItemData As Long, _
      Optional ByVal fWidth As Single, _
      Optional ByVal fHeight As Single, _
      Optional ByVal vBandBefore As Variant) _
            As cBand
    Set Add = moOwner.fBands_Add(oBandChild, sKey, sText, bUseChevron, bBreakLine, bGripper, bVisible, bFixedSize, iItemData, fWidth, fHeight, vBandBefore)
End Function

Public Sub Remove(ByVal vBand As Variant)
    moOwner.fBands_Remove vBand
End Sub

Public Sub Clear()
    moOwner.fBands_Clear
End Sub

Public Property Get Item(ByVal vBand As Variant) As cBand
Attribute Item.VB_UserMemId = 0
    Set Item = moOwner.fBands_Item(vBand)
End Property

Public Property Get Count() As Long
    Count = moOwner.fBands_Count()
End Property

Public Property Get Exists(ByVal vBand As Variant) As Boolean
    Exists = moOwner.fBands_Exists(vBand)
End Property
