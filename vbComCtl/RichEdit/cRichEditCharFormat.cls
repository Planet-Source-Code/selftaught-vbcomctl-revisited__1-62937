VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichEditCharFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cRichEditCharFormat.cls                    12/15/04
'
'           PURPOSE:
'               Access character format options on a rich edit control.
'
'==================================================================================================

Option Explicit

Private moOwner As ucRichEdit
Private miType As eRichEditFormatType
Private mbConsistent As Boolean

Friend Sub fInit(ByVal oOwner As ucRichEdit, ByVal iType As Long)
    Set moOwner = oOwner
    miType = iType
End Sub

Public Property Get Consistent() As Boolean
    Consistent = mbConsistent
End Property

Public Property Get Bold() As Boolean
    Bold = moOwner.fChar_Bold(miType, mbConsistent)
End Property
Public Property Let Bold(ByVal bNew As Boolean)
    moOwner.fChar_Bold(miType, mbConsistent) = bNew
End Property

Public Property Get Italic() As Boolean
    Italic = moOwner.fChar_Italic(miType, mbConsistent)
End Property
Public Property Let Italic(ByVal bNew As Boolean)
    moOwner.fChar_Italic(miType, mbConsistent) = bNew
End Property

Public Property Get Underline() As Boolean
    Underline = moOwner.fChar_Underline(miType, mbConsistent)
End Property
Public Property Let Underline(ByVal bNew As Boolean)
    moOwner.fChar_Underline(miType, mbConsistent) = bNew
End Property

Public Property Get Strikeout() As Boolean
    Strikeout = moOwner.fChar_Strikeout(miType, mbConsistent)
End Property
Public Property Let Strikeout(ByVal bNew As Boolean)
    moOwner.fChar_Strikeout(miType, mbConsistent) = bNew
End Property

Public Property Get FaceName() As String
    FaceName = moOwner.fChar_FaceName(miType, mbConsistent)
End Property
Public Property Let FaceName(ByRef sNew As String)
    moOwner.fChar_FaceName(miType, mbConsistent) = sNew
End Property

Public Property Get Height() As Long
    Height = moOwner.fChar_Height(miType, mbConsistent)
End Property
Public Property Let Height(ByVal iNew As Long)
    moOwner.fChar_Height(miType, mbConsistent) = iNew
End Property

Public Property Get Offset() As Long
    Offset = moOwner.fChar_Offset(miType, mbConsistent)
End Property
Public Property Let Offset(ByVal iNew As Long)
    moOwner.fChar_Offset(miType, mbConsistent) = iNew
End Property

Public Property Get ColorFore() As OLE_COLOR
    ColorFore = moOwner.fChar_ColorFore(miType, mbConsistent)
End Property
Public Property Let ColorFore(ByVal iNew As OLE_COLOR)
    moOwner.fChar_ColorFore(miType, mbConsistent) = iNew
End Property

Public Property Get Protected() As Boolean
    Protected = moOwner.fChar_Protected(miType, mbConsistent)
End Property
Public Property Let Protected(ByVal bNew As Boolean)
    moOwner.fChar_Protected(miType, mbConsistent) = bNew
End Property

Public Property Get ColorBack() As OLE_COLOR
    ColorBack = moOwner.fChar_ColorBack(miType, mbConsistent)
End Property
Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
    moOwner.fChar_ColorBack(miType, mbConsistent) = iNew
End Property

Public Property Get Link() As Boolean
    Link = moOwner.fChar_Link(miType, mbConsistent)
End Property
Public Property Let Link(ByVal bNew As Boolean)
    moOwner.fChar_Link(miType, mbConsistent) = bNew
End Property
