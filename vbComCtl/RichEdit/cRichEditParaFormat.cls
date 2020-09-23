VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichEditParaFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cRichEditParaFormat.cls                    12/15/04
'
'           PURPOSE:
'               Access paragraph format options on a rich edit control.
'
'==================================================================================================

Option Explicit

Private moOwner As ucRichEdit
Private mbConsistent As Boolean

Friend Sub fInit(ByVal oOwner As ucRichEdit)
    Set moOwner = oOwner
End Sub

Public Property Get Consistent() As Boolean
    Consistent = mbConsistent
End Property

Public Property Get Alignment() As eRichEditParaAlignment
    Alignment = moOwner.fPara_Alignment(mbConsistent)
End Property
Public Property Let Alignment(ByVal iNew As eRichEditParaAlignment)
    moOwner.fPara_Alignment(mbConsistent) = iNew
End Property

Public Property Get Indent() As Long
    Indent = moOwner.fPara_Indent(mbConsistent)
End Property
Public Property Let Indent(ByVal iNew As Long)
    moOwner.fPara_Indent(mbConsistent) = iNew
End Property

Public Property Get HangingIndent() As Long
    HangingIndent = moOwner.fPara_HangingIndent(mbConsistent)
End Property
Public Property Let HangingIndent(ByVal iNew As Long)
    moOwner.fPara_HangingIndent(mbConsistent) = iNew
End Property

Public Property Get RightIndent() As Long
    RightIndent = moOwner.fPara_RightIndent(mbConsistent)
End Property
Public Property Let RightIndent(ByVal iNew As Long)
    moOwner.fPara_RightIndent(mbConsistent) = iNew
End Property
