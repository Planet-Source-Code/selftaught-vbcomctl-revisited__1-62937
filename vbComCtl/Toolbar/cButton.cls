VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cButton.cls                    12/15/04
'
'           PURPOSE:
'               Represent a toolbar button.
'
'==================================================================================================

Option Explicit

Private moOwner As ucToolbar
Private miIndex As Long
Private miId As Long

Friend Sub fInit(ByVal oOwner As ucToolbar, ByVal iId As Long, ByVal iIndex As Long)
    Set moOwner = oOwner
    miIndex = iIndex
    miId = iId
End Sub

Friend Property Get fIsMine(ByVal oOwner As ucToolbar)
    fIsMine = oOwner Is moOwner
End Property


Public Property Get Text() As String
    Text = moOwner.fButton_Text(miId, miIndex)
End Property
Public Property Let Text(ByRef sNew As String)
    moOwner.fButton_Text(miId, miIndex) = sNew
End Property

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fButton_IconIndex(miId, miIndex)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fButton_IconIndex(miId, miIndex) = iNew
End Property

Public Property Get ToolTipText() As String
    ToolTipText = moOwner.fButton_ToolTipText(miId, miIndex)
End Property
Public Property Let ToolTipText(ByRef sNew As String)
    moOwner.fButton_ToolTipText(miId, miIndex) = sNew
End Property


Public Property Get Enabled() As Boolean
    Enabled = moOwner.fButton_Enabled(miId, miIndex)
End Property
Public Property Let Enabled(ByVal bNew As Boolean)
    moOwner.fButton_Enabled(miId, miIndex) = bNew
End Property

Public Property Get Checked() As Boolean
    Checked = moOwner.fButton_Checked(miId, miIndex)
End Property
Public Property Let Checked(ByVal bNew As Boolean)
    moOwner.fButton_Checked(miId, miIndex) = bNew
End Property

Public Property Get Pressed() As Boolean
    Pressed = moOwner.fButton_Pressed(miId, miIndex)
End Property
Public Property Let Pressed(ByVal bNew As Boolean)
    moOwner.fButton_Pressed(miId, miIndex) = bNew
End Property

Public Property Get Visible() As Boolean
    Visible = moOwner.fButton_Visible(miId, miIndex)
End Property
Public Property Let Visible(ByVal bNew As Boolean)
    moOwner.fButton_Visible(miId, miIndex) = bNew
End Property

Public Property Get Grayed() As Boolean
    Grayed = moOwner.fButton_Grayed(miId, miIndex)
End Property
Public Property Let Grayed(ByVal bNew As Boolean)
    moOwner.fButton_Grayed(miId, miIndex) = bNew
End Property

Public Property Get Key() As String
    Key = moOwner.fButton_Key(miId, miIndex)
End Property
Public Property Let Key(ByRef sNew As String)
    moOwner.fButton_Key(miId, miIndex) = sNew
End Property

Public Property Get ItemData() As Long
    ItemData = moOwner.fButton_ItemData(miId, miIndex)
End Property
Public Property Let ItemData(ByVal iNew As Long)
    moOwner.fButton_ItemData(miId, miIndex) = iNew
End Property

Public Property Get Left() As Single
    Left = moOwner.fButton_Left(miId, miIndex)
End Property

Public Property Get Width() As Single
    Width = moOwner.fButton_Width(miId, miIndex)
End Property

Public Property Get Top() As Single
    Top = moOwner.fButton_Top(miId, miIndex)
End Property

Public Property Get Height() As Single
    Height = moOwner.fButton_Height(miId, miIndex)
End Property

Public Property Get Index() As Long
    Index = moOwner.fButton_Index(miId, miIndex)
End Property

Public Sub GetIdealPopup(ByRef fLeft As Single, ByRef fTop As Single, Optional ByRef fExcludeLeft As Single, Optional ByRef fExcludeTop As Single, Optional ByRef fExcludeWidth As Single, Optional ByRef fExcludeHeight As Single, Optional ByRef bPreserveVertAlign As Boolean)
    moOwner.fButton_GetIdealPopup miId, miIndex, fLeft, fTop, fExcludeLeft, fExcludeTop, fExcludeWidth, fExcludeHeight, bPreserveVertAlign
End Sub
