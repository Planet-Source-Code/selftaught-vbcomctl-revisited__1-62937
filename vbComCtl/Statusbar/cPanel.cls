VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cPanel.cls                    12/15/04
'
'           PURPOSE:
'               Represent a status bar panel.
'
'==================================================================================================

Option Explicit

Private miId As Long
Private miIndex As Long
Private moOwner As ucStatusBar

Friend Sub fInit(ByVal oOwner As ucStatusBar, ByVal iId As Long, ByVal iIndex As Long)
    Set moOwner = oOwner
    miId = iId
    miIndex = iIndex
End Sub

Public Property Get IconIndex() As Long
    IconIndex = moOwner.fPanel_IconIndex(miId, miIndex)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fPanel_IconIndex(miId, miIndex) = iNew
End Property
    
Public Property Get Key() As String
    Key = moOwner.fPanel_Key(miId, miIndex)
End Property
Public Property Let Key(ByRef sNew As String)
    moOwner.fPanel_Key(miId, miIndex) = sNew
End Property

Public Property Get ToolTipText() As String
    ToolTipText = moOwner.fPanel_ToolTipText(miId, miIndex)
End Property
Public Property Let ToolTipText(ByRef sNew As String)
    moOwner.fPanel_ToolTipText(miId, miIndex) = sNew
End Property

Public Property Get Text() As String
    Text = moOwner.fPanel_Text(miId, miIndex)
End Property
Public Property Let Text(ByRef sNew As String)
    moOwner.fPanel_Text(miId, miIndex) = sNew
End Property

Public Property Get MinWidth() As Single
    MinWidth = moOwner.fPanel_MinWidth(miId, miIndex)
End Property

Public Property Let MinWidth(ByVal fNew As Single)
    moOwner.fPanel_MinWidth(miId, miIndex) = fNew
End Property

Public Property Get IdealWidth() As Single
    IdealWidth = moOwner.fPanel_IdealWidth(miId, miIndex)
End Property

Public Property Let IdealWidth(ByVal iNew As Single)
    moOwner.fPanel_IdealWidth(miId, miIndex) = iNew
End Property

Public Property Get Style() As eStatusBarPanelStyle
    Style = moOwner.fPanel_Style(miId, miIndex)
End Property
Public Property Let Style(ByVal iNew As eStatusBarPanelStyle)
    moOwner.fPanel_Style(miId, miIndex) = iNew
End Property

Public Property Get Border() As eStatusBarPanelBorder
    Border = moOwner.fPanel_Border(miId, miIndex)
End Property
Public Property Let Border(ByVal iNew As eStatusBarPanelBorder)
    moOwner.fPanel_Border(miId, miIndex) = iNew
End Property

Public Property Get Spring() As Boolean
    Spring = moOwner.fPanel_Spring(miId, miIndex)
End Property

Public Property Let Spring(ByVal bNew As Boolean)
    moOwner.fPanel_Spring(miId, miIndex) = bNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fPanel_Index(miId, miIndex)
End Property
