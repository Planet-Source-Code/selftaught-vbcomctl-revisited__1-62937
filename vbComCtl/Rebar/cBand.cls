VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cBand.cls                    12/15/04
'
'           PURPOSE:
'               Represent a rebar band.
'
'==================================================================================================

Option Explicit

Private moOwner As ucRebar
Private miId As Long
Private miIndex As Long

Friend Sub fInit(ByVal iId As Long, ByVal iIndex As Long, ByVal oOwner As ucRebar)
    miId = iId
    miIndex = iIndex
    Set moOwner = oOwner
End Sub

Friend Property Get fIsMine(ByVal oOwner As ucRebar) As Boolean
    fIsMine = (oOwner Is moOwner)
End Property

Public Property Get Child() As Object
    Set Child = moOwner.fBand_Child(miId, miIndex)
End Property

Public Property Get IdealWidth() As Single
    IdealWidth = moOwner.fBand_IdealWidth(miId, miIndex)
End Property
Public Property Let IdealWidth(ByVal fNew As Single)
    moOwner.fBand_IdealWidth(miId, miIndex) = fNew
End Property

Public Property Get IdealHeight() As Single
    IdealHeight = moOwner.fBand_IdealHeight(miId, miIndex)
End Property
Public Property Let IdealHeight(ByVal fNew As Single)
    moOwner.fBand_IdealHeight(miId, miIndex) = fNew
End Property

Public Property Get Key() As String
    Key = moOwner.fBand_Key(miId, miIndex)
End Property

Public Property Get Text() As String
    Text = moOwner.fBand_Text(miId, miIndex)
End Property
Public Property Let Text(ByRef sNew As String)
    moOwner.fBand_Text(miId, miIndex) = sNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fBand_Index(miId, miIndex)
End Property
Public Property Let Index(ByVal iNew As Long)
    moOwner.fBand_Index(miId, miIndex) = iNew
End Property

Public Property Get Visible() As Boolean
    Visible = moOwner.fBand_Visible(miId, miIndex)
End Property
Public Property Let Visible(ByVal bNew As Boolean)
    moOwner.fBand_Visible(miId, miIndex) = bNew
End Property

Public Property Get Gripper() As Boolean
    Gripper = moOwner.fBand_Gripper(miId, miIndex)
End Property
Public Property Let Gripper(ByVal bNew As Boolean)
    moOwner.fBand_Gripper(miId, miIndex) = bNew
End Property
