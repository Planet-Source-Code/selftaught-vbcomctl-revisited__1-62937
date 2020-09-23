Attribute VB_Name = "mRichEdit"
'==================================================================================================
'mRichEdit.bas                      12/15/04
'
'           PURPOSE:
'               General procedures for the richedit control.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================

Option Explicit

Public Property Get RichEdit_Lib() As pcRichEditLib
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Return a shared reference to the richedit dll.
'---------------------------------------------------------------------------------------
    Static o As pcRichEditLib
    If o Is Nothing Then Set o = New pcRichEditLib
    Set RichEdit_Lib = o
End Property

Public Function GetKBState(ByVal wParam As Long) As evbComCtlKeyboardState
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Get a keyboard state value from a wndproc wParam
'---------------------------------------------------------------------------------------
    Const MK_SHIFT As Long = &H4
    Const MK_CONTROL As Long = &H8
    'Const MK_ALT As Long = &H20
    Dim liLoWord As Integer: liLoWord = loword(wParam)
    GetKBState = ((vbccControlMask * Sgn(liLoWord And MK_CONTROL)) Or (vbccShiftMask * Sgn(liLoWord And MK_SHIFT))) _
                        Or (vbccAltMask * Abs(KeyIsDown(VK_MENU)))
End Function

Public Function GetMouseButton(ByVal wParam As Long) As evbComCtlMouseButton
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Get a mouse button value from a wndproc wParam
'---------------------------------------------------------------------------------------
    GetMouseButton = loword(wParam) And (vbccMouseLButton Or vbccMouseMButton Or vbccMouseRButton Or vbccMouseXButton1 Or vbccMouseXButton2)
End Function
