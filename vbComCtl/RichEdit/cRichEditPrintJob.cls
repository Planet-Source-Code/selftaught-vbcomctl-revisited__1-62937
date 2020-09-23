VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichEditPrintJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'cRichEditPrintJob.cls                    12/15/04
'
'           PURPOSE:
'               Format the display of rich edit control into a print dc.
'
'==================================================================================================

Option Explicit

Private moOwner As ucRichEdit
Private mbCallStartEndDoc As Boolean
Private mbCallStartEndPage As Boolean

Private mtFormatRange As FORMATRANGE

Friend Sub fInit(ByVal oOwner As ucRichEdit, ByVal hdc As Long, ByVal bCallStartEndDoc As Boolean, ByVal bCallStartEndPage As Boolean)
    Set moOwner = oOwner
    mbCallStartEndDoc = bCallStartEndDoc
    mbCallStartEndPage = bCallStartEndPage
    With mtFormatRange
        .hdc = hdc
        .hdcTarget = hdc
        .chrg.cpMax = NegOneL
        With .rcPage
            .Left = ZeroL
            .Top = ZeroL
            .Right = GetDeviceCaps(hdc, HORZRES) / GetDeviceCaps(hdc, LOGPIXELSX) * 1440&
            .bottom = ((GetDeviceCaps(hdc, VERTRES) / GetDeviceCaps(hdc, LOGPIXELSY))) * 1440&
        End With
    End With
End Sub

Public Function DoPrint(ByVal iLeftMargin As Long, ByVal iTopMargin As Long, ByVal iRightMargin As Long, ByVal iBottomMargin As Long) As Boolean
    With mtFormatRange.rc
        .Left = mtFormatRange.rcPage.Left + iLeftMargin
        .Top = mtFormatRange.rcPage.Top + iTopMargin
        .Right = mtFormatRange.rcPage.Right - iRightMargin
        .bottom = mtFormatRange.rcPage.bottom - iBottomMargin
    End With
    DoPrint = moOwner.fPrint_DoPrint(mtFormatRange, mbCallStartEndPage)
End Function

Private Sub Class_Terminate()
    moOwner.fPrint_Terminate mtFormatRange, mbCallStartEndDoc
End Sub
