VERSION 5.00
Begin VB.UserControl ucMaskedEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ClipControls    =   0   'False
   HasDC           =   0   'False
   MousePointer    =   3  'I-Beam
   PropertyPages   =   "ucMaskedEdit.ctx":0000
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   ToolboxBitmap   =   "ucMaskedEdit.ctx":000D
End
Attribute VB_Name = "ucMaskedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'ucMaskedEdit.ctl           2/13/05
'
'            PURPOSE:
'               "Look & Feel" just like a textbox, including the automatic popup menu,
'               arrow key operation, hot keys, selection color, etc.
'
'               Restrict textual input based on a mask.  The mask is the same as
'               the pattern used by vb's Like Operator.
'
'---------------------------------------------------------------------------------------
Option Explicit

Implements iSubclass

Public Event Changed()

Private Type tMaskedChar
    sMask       As String
    iChar       As Integer
End Type

Private Const EDT_Insert = 1
Private Const EDT_Delete = 2

Private Const UNDO_Timeout As Long = 8000

Private Const DEF_Mask = vbNullString
Private Const DEF_Themeable = True

Private Const PROP_Font = "Font"
Private Const PROP_Mask = "Mask"
Private Const PROP_Themeable = "Themeable"

Private WithEvents moFont       As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage   As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1
Private mhFont                  As Long

Private miSelStart              As Long
Private miSelLength             As Long
Private miSelMark               As Long
Private miLineHeight            As Long
Private mfAvgCharWidth          As Single

Private mbInFocus               As Boolean
Private mbThemeable             As Boolean

Private mtChars()               As tMaskedChar
Private miCharCount             As Long
Private miTextLen               As Long
Private miTextControl           As Long
Private miXOffset               As Long
Private miFirstVisibleChar      As Long

Private msUndoText              As String
Private miUndoTimeStamp         As Long

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    If uMsg = WM_SETFOCUS Then
        mbInFocus = True
        CreateCaret hWnd, ZeroL, GetCaretWidth(), miLineHeight
        pUpdateCaretPos
        ShowCaret hWnd
        pInvalidate miSelStart, miSelStart + miSelLength
    End If
    
End Sub
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

    If uMsg = WM_KILLFOCUS Then
        HideCaret hWnd
        mbInFocus = False
        pInvalidate miSelStart, miSelStart + miSelLength
    End If

End Sub

Private Sub moFont_Changed()
    moFont.OnAmbientFontChanged Ambient.Font
    pFontChanged
    PropertyChanged PROP_Font
End Sub

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_DblClick()
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Select all of the text in the control.
'---------------------------------------------------------------------------------------
    pUpdateSelection ZeroL, miTextLen
End Sub

Private Sub UserControl_Initialize()
    'we need the messages rather than the usercontrol_*focus events
    'because we need to hide the cursor if the application goes out
    'of focus, which doesn't trigger these events.
    Subclass_Install Me, hWnd, WM_KILLFOCUS, WM_SETFOCUS
    Set moFontPage = New pcSupportFontPropPage
    miTextControl = NegOneL
End Sub

Private Sub UserControl_InitProperties()
    Set moFont = Font_CreateDefault(Ambient.Font)
    pFontChanged
    Mask = DEF_Mask
    mbThemeable = DEF_Themeable
    pSetTheme
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Move the cursor according to arrow keys, delete on the delete/backspace key,
'             show popup menu if menu key is pressed, catch cut/copy/paste/undo hotkeys.
'---------------------------------------------------------------------------------------
    Dim liPos As Long
    
    Select Case CLng(KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd
        Select Case CLng(KeyCode)
        Case vbKeyLeft
            liPos = miSelStart
            If Not CBool(CLng(Shift) And vbShiftMask) Then
                If miSelLength < ZeroL Then liPos = liPos + miSelLength
            End If
            If CBool(CLng(Shift) And vbCtrlMask) _
                Then liPos = pWordBreakProc(NegOneL, liPos) _
                Else liPos = liPos + NegOneL
            If liPos < ZeroL Then liPos = ZeroL
        Case vbKeyRight
            liPos = miSelStart
            If Not CBool(CLng(Shift) And vbShiftMask) Then
                If miSelLength > ZeroL Then liPos = liPos + miSelLength
            End If
            If CBool(CLng(Shift) And vbCtrlMask) Then liPos = pWordBreakProc(OneL, liPos) Else liPos = liPos + OneL
            If liPos > miTextLen Then liPos = miTextLen
        Case vbKeyHome: liPos = ZeroL
        Case vbKeyEnd:  liPos = miTextLen
        End Select

        pUpdateSelection liPos, IIf(CBool(CLng(Shift) And vbShiftMask), miSelStart - liPos + miSelLength, ZeroL), True
        
    Case vbKeyDelete, vbKeyBack
        liPos = miSelStart + miSelLength
        
        If miSelLength = ZeroL Then
            If KeyCode = vbKeyBack Then
                If liPos > ZeroL Then liPos = liPos - OneL
                pGetNextNonLiteral liPos, True
            Else
                pGetNextNonLiteral liPos, False
                If liPos < miTextLen Then liPos = liPos + OneL
            End If
        End If
    
        pDelete miSelStart, liPos - miSelStart
        
    Case 93
        pShowMenu ScaleWidth \ TwoL, ScaleHeight \ TwoL
    Case vbKeyX
        If (Shift And vbCtrlMask) Then Cut
    Case vbKeyC
        If (Shift And vbCtrlMask) Then Copy
    Case vbKeyV
        If (Shift And vbCtrlMask) Then Paste
    Case vbKeyZ
        If (Shift And vbCtrlMask) Then Undo
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Insert the pressed character if it doesn't violate the mask.
'---------------------------------------------------------------------------------------
    pInsert ChrW$(KeyAscii)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Move the cursor to the character position horizontally nearest the mouse.
'---------------------------------------------------------------------------------------
    If Button = vbLeftButton Then
        On Error GoTo catch
        'bugfix 2/20/05 to make cursor show when clicking on ucMaskedEdit control directly from menu tracking mode on ucRebar
        If Not Screen.ActiveControl Is Me Then UserControl.SetFocus
catch:  On Error GoTo 0
        
        miSelMark = pHitTest(x)
        pUpdateSelection miSelMark, ZeroL
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Allow the user to drag the left mouse button to select text.
'---------------------------------------------------------------------------------------
    Dim liHitTest As Long
    liHitTest = pHitTest(x)
    
    If Button = vbLeftButton Then pUpdateSelection liHitTest, miSelMark - liHitTest
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Show the popup menu on right button up.
'---------------------------------------------------------------------------------------
    
    If Button = vbRightButton Then
        On Error GoTo finally
        If Not Screen.ActiveControl Is Me Then UserControl.SetFocus
        
finally:
        On Error GoTo 0
        pShowMenu x, y
    End If
End Sub

Private Sub pSetTheme()
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Change the usercontrol border and x offset to indicate whether we are
'             using the default theme.
'---------------------------------------------------------------------------------------
    If Not IsAppThemed() Or Not mbThemeable Then
        UserControl.Appearance = 1
        UserControl.BorderStyle = 1
        UserControl.BackColor = vbWindowBackground
        miXOffset = 1&
    Else
        UserControl.Appearance = 0
        UserControl.BorderStyle = 0
        UserControl.BackColor = vbWindowBackground
        miXOffset = 3&
    End If
    pUpdateCaretPos
    Refresh
End Sub

Private Sub pDrawBorder(ByVal lhDc As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Draw a one-pixel border the same color as the border drawn by the textbox
'             control under the default theme.
'---------------------------------------------------------------------------------------
    Dim lhPen As Long
    Dim lhPenOld As Long
    Dim ltPoint As POINT
    
    lhPen = GdiMgr_CreatePen(PS_SOLID, OneL, RGB(165, 172, 178))
    If lhPen Then
        lhPenOld = SelectObject(lhDc, lhPen)
        If lhPenOld Then
            MoveToEx lhDc, ZeroL, ZeroL, ltPoint
            LineTo lhDc, ScaleWidth - OneL, ZeroL
            LineTo lhDc, ScaleWidth - OneL, ScaleHeight - OneL
            LineTo lhDc, ZeroL, ScaleHeight - OneL
            LineTo lhDc, ZeroL, ZeroL
            SelectObject lhDc, lhPenOld
        End If
        GdiMgr_DeletePen lhPen
    End If
    
End Sub

Private Sub UserControl_Paint()
    Dim lhDc As Long
    Dim lhFontOld As Long
    Dim liBkModeOld As Long
    Dim ltRect As RECT
    Dim lsText As String
    Dim lpString As Long
    
    Dim liSelLeft As Long
    Dim liSelRight As Long
    Dim liLen As Long
    Dim liTop As Long
    liTop = ScaleHeight \ TwoL - miLineHeight \ TwoL
    
    lhDc = hDc
    
    lsText = StrConv(pGetText(), vbFromUnicode)
    lpString = UnsignedAdd(StrPtr(lsText), miFirstVisibleChar)
    liSelLeft = pLeft(miSelStart, miSelLength) - miFirstVisibleChar
    If liSelLeft < ZeroL Then liSelLeft = ZeroL
    liSelRight = pRight(miSelStart, miSelLength) - miFirstVisibleChar
    liLen = miTextLen - miFirstVisibleChar
    
'---------------------------------------------------------------------------------------
' Step 1:
'     Validate the state.  We must get a valid hDc and a valid font, then sucessfully
'     select the font into the dc.
'---------------------------------------------------------------------------------------
    
    If lhDc Then
        If mhFont Then
            lhFontOld = SelectObject(lhDc, mhFont)
            If lhFontOld Then
            
'---------------------------------------------------------------------------------------
' Step 2:
'     If we have a selection then paint the selection using the system highlight color.
'---------------------------------------------------------------------------------------

                If liSelRight > liSelLeft Then
                    If GetTextExtentPoint32(lhDc, ByVal lpString, liSelLeft, ltRect.Left) Then
                        If GetTextExtentPoint32(lhDc, ByVal lpString, liSelRight, ltRect.Right) Then
                            With ltRect
                                .Left = .Left + miXOffset
                                .Right = .Right + miXOffset
                                .Top = ScaleHeight \ TwoL - miLineHeight \ TwoL
                                .bottom = .Top + miLineHeight
                            End With
                            If mbInFocus Then FillRect lhDc, ltRect, GetSysColorBrush(COLOR_HIGHLIGHT)
                        End If
                    End If
                End If
                
'---------------------------------------------------------------------------------------
' Step 3:
'     Paint the text.
'---------------------------------------------------------------------------------------

                liBkModeOld = SetBkMode(lhDc, OneL)
                
                If liBkModeOld Then
                    If liSelRight > liSelLeft Then
                        Dim liTextColorOld As Long
                        
                        TextOut lhDc, miXOffset, liTop, ByVal lpString, liSelLeft
                        If mbInFocus Then liTextColorOld = SetTextColor(lhDc, TranslateColor(vbHighlightText))
                        TextOut lhDc, ltRect.Left, liTop, ByVal UnsignedAdd(lpString, liSelLeft), liSelRight - liSelLeft
                        If mbInFocus Then SetTextColor lhDc, liTextColorOld
                        TextOut lhDc, ltRect.Right, liTop, ByVal UnsignedAdd(lpString, liSelRight), liLen - liSelRight
                    Else
                        TextOut lhDc, miXOffset, liTop, ByVal lpString, liLen
                    End If
                    
                    SetBkMode lhDc, liBkModeOld
                End If
                
                SelectObject lhDc, lhFontOld
            End If
        End If
        
        If IsAppThemed() And mbThemeable Then pDrawBorder lhDc
        
        'ReleaseDC hWnd, lhDc
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    pFontChanged
    Mask = PropBag.ReadProperty(PROP_Mask, DEF_Mask)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    pSetTheme
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Terminate()
    Subclass_Remove Me, hWnd
    Set moFontPage = Nothing
    If mhFont Then moFont.ReleaseHandle mhFont
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Mask, Mask
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the current insertion point
'---------------------------------------------------------------------------------------
    SelStart = miSelStart
End Property
Public Property Let SelStart(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Set the current insertion point
'---------------------------------------------------------------------------------------
    pUpdateSelection iNew, ZeroL
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the selection length.  This can be a ngative value
'---------------------------------------------------------------------------------------
    SelLength = miSelLength
End Property
Public Property Let SelLength(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Set the selection length.  A negative value selects to the left, a positive
'             value to the right.
'---------------------------------------------------------------------------------------
    pUpdateSelection miSelStart, iNew
End Property

Public Property Get Mask() As String
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the mask used to validate input. This is the same as vb's Like Operator.
'---------------------------------------------------------------------------------------
    Dim liLen As Long
    Dim liLen2 As Long
    Dim i As Long
    
    For i = ZeroL To miCharCount - OneL
        liLen = liLen + Len(mtChars(i).sMask)
    Next
    
    Mask = Space$(liLen)
    
    liLen = OneL
    For i = ZeroL To miCharCount - OneL
        liLen2 = Len(mtChars(i).sMask)
        Mid$(Mask, liLen, liLen2) = mtChars(i).sMask
        liLen = liLen + liLen2
    Next
    
End Property
Public Property Let Mask(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Set the mask used to validate input. This is the same as vb's Like Operator.
'---------------------------------------------------------------------------------------
    Dim liPos As Long
    Dim liLength As Long
    
    liLength = Len(sNew)
    
    ReDim mtChars(ZeroL To liLength)
    
    miCharCount = ZeroL
    
    For liPos = OneL To liLength
        liLength = pLogicalCharLength(sNew, liPos)
        mtChars(miCharCount).sMask = Mid$(sNew, liPos, liLength)
        mtChars(miCharCount).iChar = AscW(mtChars(miCharCount).sMask)
        miCharCount = miCharCount + OneL
        liPos = liPos + liLength - OneL
    Next
    
    If miCharCount Then ReDim Preserve mtChars(ZeroL To miCharCount - OneL)
    
    miTextLen = ZeroL
    miTextControl = NegOneL
    
    miSelStart = ZeroL
    miSelLength = ZeroL
    
    miUndoTimeStamp = ZeroL
    msUndoText = vbNullString
    
    InvalidateRect hWnd, ByVal ZeroL, ZeroL
    
    RaiseEvent Changed
    
End Property

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return a font object used by this control
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Set the font object used by this control
'---------------------------------------------------------------------------------------
    If oNew Is Nothing Then Set oNew = Font_CreateDefault(Ambient.Font)
    Set moFont = oNew
    pFontChanged
    PropertyChanged PROP_Font
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the selected text.
'---------------------------------------------------------------------------------------
    SelText = Mid$(pGetText(), pLeft(miSelStart, miSelLength) + OneL, Abs(miSelLength))
End Property
Public Property Let SelText(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Replace the selected text.
'---------------------------------------------------------------------------------------
    Dim liPos As Long
    liPos = pLeft(miSelStart, miSelLength)
    If Len(sNew) Then
        pInsert sNew
        pUpdateSelection miSelStart, -(miSelStart - liPos)
    End If
End Property

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the text as displayed.
'---------------------------------------------------------------------------------------
    Text = pGetText()
End Property
Public Property Let Text(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Set the text. The literals may be included, but if they are included then
'             they must be in exactly the right places.
'---------------------------------------------------------------------------------------
    miSelStart = ZeroL
    miSelLength = miTextLen
    pInsert sNew
End Property

Public Sub Copy()
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Copy the selected text to the clipboard.
'             If no text is selected, copy all of the text.
'---------------------------------------------------------------------------------------
    If miSelLength Then
        Clipboard.Clear
        Clipboard.SetText Mid$(pGetText(), pLeft(miSelStart, miSelLength) + OneL, Abs(miSelLength)), vbCFText
    ElseIf miTextLen Then
        Clipboard.Clear
        Clipboard.SetText pGetText(), vbCFText
    End If
End Sub

Public Sub Paste()
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Paste the text data from the clipboard, if any.
'---------------------------------------------------------------------------------------
    If Clipboard.GetFormat(vbCFText) Then pInsert Clipboard.GetText()
End Sub

Public Sub Cut()
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Copy the selected text to the clipboard, then delete it.
'---------------------------------------------------------------------------------------
    Copy
    pDelete miSelStart, miSelLength

End Sub

Public Sub Delete()
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Delete the selected text
'---------------------------------------------------------------------------------------
    pDelete miSelStart, miSelLength
End Sub

Public Sub Undo()
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Exchange the current text and the text in the undo buffer.
'---------------------------------------------------------------------------------------

    Dim lsText As String
    
    miSelStart = ZeroL
    miSelLength = miTextLen
    
    lsText = pGetText()
    
    miUndoTimeStamp = ZeroL
    pInsert msUndoText
    miUndoTimeStamp = GetTickCount()
    If miUndoTimeStamp > UNDO_Timeout Then miUndoTimeStamp = miUndoTimeStamp - UNDO_Timeout
    
End Sub

Public Property Get IsValid() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return True if the mask and text are consistent.
'---------------------------------------------------------------------------------------
    IsValid = CBool(miTextLen = miCharCount)
End Property

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property


Private Function pLiteral(ByVal iPos As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return a value indicating whether the character at this position is variable
'             or whether it is a literal value.
'---------------------------------------------------------------------------------------

    If iPos > NegOneL And iPos < miCharCount Then
        
        iPos = AscW(mtChars(iPos).sMask)
        pLiteral = CBool(iPos <> 91 And iPos <> 42 And iPos <> 63 And iPos <> 35)
                      '[, *, ?, #
    End If
End Function

Private Sub pGetNextNonLiteral(ByRef iPos As Long, Optional ByVal bBack As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : If the position points to a literal position in the mask, increment it until
'             either the end of the mask or a non-literal position.
'---------------------------------------------------------------------------------------
    
    Do Until Not pLiteral(iPos)
        If bBack Then iPos = iPos - OneL Else iPos = iPos + OneL
    Loop
    If iPos < ZeroL Then iPos = ZeroL
End Sub

Private Sub pDelete(ByVal iStart As Long, ByVal iLength As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Delete the selected text.
'---------------------------------------------------------------------------------------
    If iLength Then
        
        pGetText
        
        Dim liPos As Long
        Dim liRight As Long
        Dim liOldTextLen As Long
        
        liOldTextLen = miTextLen
        liRight = pRight(iStart, iLength)
        liPos = pShiftText(liRight, -Abs(iLength))
        
        If liRight <> liPos Then
            InvalidateRect hWnd, ByVal ZeroL, OneL
            pCheckUndo EDT_Delete
            pUpdateSelection liPos, (liRight - Abs(iLength)) - liPos, True
            pInvalidate pLeft(iStart, iLength), liOldTextLen
            Incr miTextControl
            RaiseEvent Changed
        End If
        
    End If
End Sub

Private Sub pInsert(ByRef sText As String)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Insert text, replacing the selection.
'---------------------------------------------------------------------------------------
    
    pGetText
    
    Dim liPos As Long
    Dim liSelLeft As Long
    Dim liSelRight As Long
    Dim liLeft As Long
    Dim liTextLengthOld As Long
    Dim lbDirty As Boolean
    
    liTextLengthOld = miTextLen
        
    liSelLeft = pLeft(miSelStart, miSelLength)
    liSelRight = pRight(miSelStart, miSelLength)
    liLeft = liSelLeft
    
    Do While pInsertSub(sText, liPos, liLeft, liSelRight)
        lbDirty = True
    Loop
    
    Debug.Assert Len(sText)
    
    If lbDirty Then
        pCheckUndo EDT_Insert
        Debug.Assert liSelLeft <= liTextLengthOld
        If liSelLeft < liTextLengthOld Then pInvalidate liSelLeft, liTextLengthOld
        Incr miTextControl
        If liLeft < liSelRight Then pShiftText liSelRight, liLeft - liSelRight
        miSelStart = liSelLeft
        miSelLength = miTextLen - miSelStart
        pUpdateSelection liLeft, ZeroL
        RaiseEvent Changed
    End If
    
End Sub

Private Function pInsertSub(ByRef sText As String, ByRef iPos As Long, ByRef iLeft As Long, ByVal iSelRight As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Insert one character at a time.  Return true if insertion is to continue.
'---------------------------------------------------------------------------------------
    
    Dim liChar As Integer
    
    iPos = iPos + OneL
    If iPos <= Len(sText) Then
    
        liChar = AscW(Mid$(sText, iPos, OneL))
        
        Do Until Not pLiteral(iLeft)
            If pMatchMask(liChar, iLeft) Then
                iPos = iPos + OneL
                If iPos > Len(sText) Then Exit Do
                liChar = AscW(Mid$(sText, iPos, OneL))
            End If
            iLeft = iLeft + OneL
        Loop
        
        If iLeft < miCharCount Then
            If pMatchMask(liChar, iLeft) Then
                
                If iLeft >= iSelRight Then
                    If iLeft < miTextLen Then
                        If pShiftText(iLeft, OneL) = iLeft Then Exit Function
                    Else
                        'Debug.Assert miTextLen = iLeft
                        miTextLen = iLeft + OneL
                        pGetNextNonLiteral miTextLen
                    End If
                End If
                mtChars(iLeft).iChar = liChar
                iLeft = iLeft + OneL
                pInsertSub = True
                
            End If
        End If
    End If
End Function

Private Function pMatchMask(ByRef iChar As Integer, ByVal iMaskPos As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Return whether the character matches the mask at the given position.
'---------------------------------------------------------------------------------------
    
    Const UppercaseOffset As Long = 32
    pMatchMask = ChrW$(iChar) Like mtChars(iMaskPos).sMask
    If Not pMatchMask Then
        Select Case iChar
        Case vbKeyA To vbKeyZ
            pMatchMask = ChrW$(iChar + UppercaseOffset) Like mtChars(iMaskPos).sMask
            If pMatchMask Then iChar = iChar + UppercaseOffset
        Case vbKeyA + UppercaseOffset To vbKeyZ + UppercaseOffset
            pMatchMask = ChrW$(iChar - UppercaseOffset) Like mtChars(iMaskPos).sMask
            If pMatchMask Then iChar = iChar - UppercaseOffset
        End Select
    End If
End Function

Private Function pShiftText(ByVal iStart As Long, ByVal iDistance As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Shift the text to the left if deleting or to the right if inserting
'             characters, overstepping literals.  If moving the entire distance
'             produces text that violates the mask, then try to get as close
'             as possible.
'
'
'             for example, assume:
'                   0123456789                        Text of this control
'                     234567                          selected/highlighted portion of text
'
'               delete is pressed.  text must be shifted to the left, starting at index 8
'               and ending at index 2.
'
'                   01      89
'                     <-----
'
'               if the mask allows it, the resulting text will be:
'                   0189
'
'               if the mask at index 2 is "[1-7]" then the value "8" from the eighth index
'               will not be allowed.  If the mask for indices 3-4 are "##", then the
'               result would look like this:
'                  01289
'
'---------------------------------------------------------------------------------------
    Debug.Assert iDistance
    pShiftText = iStart + iDistance
    If pShiftText > miCharCount Then pShiftText = miCharCount: Debug.Assert False
    If pShiftText < ZeroL Then pShiftText = ZeroL: Debug.Assert False
    
    For pShiftText = pShiftText To IIf(iDistance > ZeroL, iStart + OneL, iStart - OneL) Step IIf(iDistance > ZeroL, NegOneL, OneL)
        If pShiftTextSub(iStart, pShiftText) Then Exit For
    Next
    
End Function

Private Function pShiftTextSub(ByVal iSrc As Long, ByVal iDst As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : If there is room and it does not violate the mask, copy
'             the characters from the dst to src.  Returns true if successful.
'---------------------------------------------------------------------------------------
    
    Dim liDstPos As Long
    Dim liSrcPos As Long
    Dim liDstEnd As Long
    Dim liSrcEnd As Long
    Dim liSrcChars As Long
    
    Dim liChars() As Integer
    
    liSrcEnd = miTextLen - OneL
    
    If liSrcEnd >= iSrc Then
    
        ReDim liChars(ZeroL To liSrcEnd - iSrc)
        
        For liSrcPos = iSrc To liSrcEnd
            If Not pLiteral(liSrcPos) Then
                liChars(liSrcChars) = mtChars(liSrcPos).iChar
                liSrcChars = liSrcChars + OneL
            End If
        Next
        
        If liSrcChars Then
            liSrcPos = ZeroL
            For liDstEnd = iDst To miCharCount - OneL
                If Not pLiteral(liDstEnd) Then
                    liSrcPos = liSrcPos + OneL
                    If liSrcPos = liSrcChars Then Exit For
                End If
            Next
            
            If liDstEnd < miCharCount Then
                liDstPos = iDst
                liSrcPos = iSrc
                Do
                    pGetNextNonLiteral liSrcPos: If (liSrcPos > liSrcEnd) Then Exit Do
                    pGetNextNonLiteral liDstPos: If (liDstPos > liDstEnd) Then Exit Do
                    If liSrcPos = liDstPos Then Exit Do
                    If Not pMatchMask(mtChars(liSrcPos).iChar, liDstPos) Then Exit Do
                    liDstPos = liDstPos + OneL
                    liSrcPos = liSrcPos + OneL
                Loop
            End If
        End If
        pShiftTextSub = (liSrcPos > liSrcEnd)
    Else
        'pShiftTextSub = ((liSrcEnd + OneL) = iSrc)
        Debug.Assert ((liSrcEnd + OneL) = iSrc)
        pShiftTextSub = True
    End If
    
    If pShiftTextSub Then
        liDstPos = iDst
        If liSrcChars Then
            For liSrcChars = ZeroL To liSrcChars - OneL
                pGetNextNonLiteral liDstPos
                mtChars(liDstPos).iChar = liChars(liSrcChars)
                liDstPos = liDstPos + OneL
            Next
        End If

        
        If liDstPos > ZeroL Then pGetNextNonLiteral liDstPos
        
        miTextLen = liDstPos
    End If
End Function

Private Sub pFontChanged()
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Update the modular variables mhFont, miLineHeight, and mfAvgCharWidth
'             Repaint the control
'             Update the caret if the control is in focus
'---------------------------------------------------------------------------------------
    
    Const sAllChars = "abcdefghijklmnopqrstuvwxyz"
    
    Dim lhFontOld As Long
    Dim lhDc As Long
    
    lhDc = GetDC(hWnd)
    miLineHeight = ZeroL
    mfAvgCharWidth = ZeroL
    
    If lhDc Then
        If mhFont Then moFont.ReleaseHandle mhFont
        mhFont = moFont.GetHandle()
        If mhFont Then
            lhFontOld = SelectObject(lhDc, mhFont)
            If lhFontOld Then
                Dim ltSize As SIZE
                If GetTextExtentPoint32W(lhDc, sAllChars, Len(sAllChars), ltSize) Then
                    miLineHeight = ltSize.cy
                    mfAvgCharWidth = ltSize.cx / Len(sAllChars)
                End If
                SelectObject lhDc, lhFontOld
            End If
        End If
        ReleaseDC hWnd, lhDc
    End If
    
    If mbInFocus Then
        CreateCaret hWnd, ZeroL, GetCaretWidth(), miLineHeight
        pUpdateCaretPos
        ShowCaret hWnd
    End If

    InvalidateRect hWnd, ByVal ZeroL, OneL
     
End Sub

Private Sub pUpdateCaretPos(Optional ByVal bFromArrowKey As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Translate a position-based offset to a distance x offset and call the api
'             to set the caret position.
'             Update the first visible character. If the call is stemming from an arrow key,
'             move the cursor 1/4 of the width of the control toward the center instead
'             of moving only just far enough to put it into view.
'---------------------------------------------------------------------------------------
    
    Dim lhFontOld As Long
    Dim lhDc As Long
    Dim ltSize As SIZE
    Dim lsText As String
    Dim lbDirty As Boolean
    
    lsText = StrConv(pGetText(), vbFromUnicode)
    
    If miSelStart < miFirstVisibleChar Then
        lbDirty = True
        miFirstVisibleChar = miSelStart
    End If
    
    lhDc = GetDC(hWnd)
    If lhDc Then
        If mhFont Then
            lhFontOld = SelectObject(lhDc, mhFont)
            If lhFontOld Then
                GetTextExtentPoint32 lhDc, ByVal UnsignedAdd(StrPtr(lsText), miFirstVisibleChar), miSelStart - miFirstVisibleChar, ltSize
                
                Do While ltSize.cx > ScaleWidth - miXOffset - miXOffset And miFirstVisibleChar + OneL < miTextLen
                    lbDirty = True
                    miFirstVisibleChar = miFirstVisibleChar + OneL
                    GetTextExtentPoint32 lhDc, ByVal UnsignedAdd(StrPtr(lsText), miFirstVisibleChar), miSelStart - miFirstVisibleChar, ltSize
                Loop
                
                If bFromArrowKey And lbDirty Then
                    If miFirstVisibleChar = miSelStart Then
                        miFirstVisibleChar = miFirstVisibleChar - (pHitTest(ScaleWidth \ 4) - miFirstVisibleChar)
                        If miFirstVisibleChar < ZeroL Then miFirstVisibleChar = ZeroL
                    Else
                        miFirstVisibleChar = pHitTest(ScaleWidth \ 4)
                    End If
                    pValidateFirstVisibleChar lsText, lhDc
                    GetTextExtentPoint32 lhDc, ByVal UnsignedAdd(StrPtr(lsText), miFirstVisibleChar), miSelStart - miFirstVisibleChar, ltSize
                End If
                
                SelectObject lhDc, lhFontOld
            End If
        End If
        
        ReleaseDC hWnd, lhDc
    End If
    
    If lbDirty Then InvalidateRect hWnd, ByVal ZeroL, ZeroL
    If mbInFocus Then SetCaretPos ltSize.cx + miXOffset, (ScaleHeight \ TwoL) - (miLineHeight \ TwoL)
End Sub

Private Sub pValidateFirstVisibleChar(ByRef sString As String, ByVal lhDc As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/19/05
' Purpose   : Make sure that the first visible character is not further to the right than
'             it needs to be.
'---------------------------------------------------------------------------------------
    Dim ltSize As SIZE
    Dim liPos As Long
    
    Dim lpString As Long
    lpString = StrPtr(sString)

    For liPos = miFirstVisibleChar - OneL To ZeroL Step NegOneL
        If GetTextExtentPoint32(lhDc, ByVal UnsignedAdd(lpString, liPos), miTextLen - liPos, ltSize) = ZeroL Then Exit For
        If ltSize.cx >= ScaleWidth - miXOffset - miXOffset Then Exit For
    Next
    miFirstVisibleChar = liPos + OneL
End Sub

Private Function pHitTest(ByVal x As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Determine the character position offset nearest the distance x offset in pixels.
'             Estimate the position using the average character width and move up or
'             down to find the nearest index.
'---------------------------------------------------------------------------------------
    
    Dim liMin As Long
    Dim liMax As Long
    Dim liX As Long
    
    Dim ltSize As SIZE
    Dim lhDc As Long
    Dim lhFontOld As Long
    Dim liHitTestNext As Long
    Dim lsText As String
    
    lsText = StrConv(pGetText(), vbFromUnicode)
    
    liMin = ZeroL
    liMax = miTextLen
    
    If liMax Then
        If mhFont Then
            lhDc = GetDC(hWnd)
            
            If lhDc Then
                lhFontOld = SelectObject(lhDc, mhFont)
                If lhFontOld Then
                    
                    If GetTextExtentPoint32(lhDc, ByVal StrPtr(lsText), miFirstVisibleChar, ltSize) Then x = x + ltSize.cx
                    
                    pHitTest = CLng(x / mfAvgCharWidth)
                    If pHitTest < liMin Then pHitTest = liMin
                    If pHitTest > liMax Then pHitTest = liMax
                    GetTextExtentPoint32 lhDc, ByVal StrPtr(lsText), pHitTest, ltSize
                    liX = ltSize.cx
                    
                    Do
                        If x = liX Then Exit Do
                        liHitTestNext = pHitTest + Sgn(x - liX)
                        If liHitTestNext > miTextLen Or liHitTestNext < ZeroL Then Exit Do
                        GetTextExtentPoint32 lhDc, ByVal StrPtr(lsText), liHitTestNext, ltSize
                        If (ltSize.cx < x And x < liX) Or (ltSize.cx > x And x > liX) Then
                            If Abs(x - ltSize.cx) < Abs(x - liX) Then pHitTest = liHitTestNext
                            Exit Do
                        End If
                        pHitTest = liHitTestNext
                        If x > liX Then liMin = pHitTest Else liMax = pHitTest
                        liX = ltSize.cx
                    Loop
                    
                    SelectObject lhDc, lhFontOld
                End If
                
                ReleaseDC hWnd, lhDc
            End If
        End If
    End If
End Function


Private Function pGetText() As String
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Get the text of the control.  A cache scheme is implemented using miTextControl
'             to indicate changes in the text.
'---------------------------------------------------------------------------------------
            
    Static sText As String
    Static iControl As Long
    If iControl <> miTextControl Then
        pGetTextSub sText
        iControl = miTextControl
    End If
    pGetText = sText
End Function

Private Sub pGetTextSub(ByRef sText As String)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Allocate a string and fill the characters using a hijacked SAFEARRAY.
'---------------------------------------------------------------------------------------
       
    Dim liArray() As Integer
    Dim loPtr As pcArrayPtr
    Dim i As Long
    
    sText = Space$(miTextLen)
    
    Set loPtr = New pcArrayPtr
    loPtr.SetArrayInt liArray
    
    loPtr.PointToString sText
    For i = ZeroL To miTextLen - OneL
        liArray(i) = mtChars(i).iChar
    Next
    Set loPtr = Nothing
    
End Sub

Private Sub pAppendMenu(ByVal hMenu As Long, ByVal wFlags As Long, ByVal dwNewItem As Long, ByVal sCaption As String)
    Dim lsAnsi As String
    lsAnsi = StrConv(sCaption & vbNullChar, vbFromUnicode)
    AppendMenu hMenu, wFlags, dwNewItem, ByVal StrPtr(lsAnsi)
End Sub

Private Sub pShowMenu(ByVal x As Long, ByVal y As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Show the familiar edit popup menu.
'---------------------------------------------------------------------------------------

    Dim lhMenu As Long
    lhMenu = CreatePopupMenu()
    If lhMenu Then
        pAppendMenu lhMenu, MF_STRING Or MF_GRAYED * (CBool(miUndoTimeStamp) + OneL), OneL, "&Undo"
        pAppendMenu lhMenu, MF_SEPARATOR, ZeroL, vbNullString
        pAppendMenu lhMenu, MF_STRING Or MF_GRAYED * (CBool(miSelLength) + OneL), TwoL, "Cu&t"
        pAppendMenu lhMenu, MF_STRING Or MF_GRAYED * (CBool(miSelLength) + OneL), 3&, "&Copy"
        pAppendMenu lhMenu, MF_STRING Or MF_GRAYED * (Clipboard.GetFormat(vbCFText) + OneL), 4&, "&Paste"
        pAppendMenu lhMenu, MF_STRING Or MF_GRAYED * (CBool(miSelLength) + OneL), 5&, "&Delete"
        pAppendMenu lhMenu, MF_SEPARATOR, ZeroL, vbNullString
        pAppendMenu lhMenu, MF_STRING Or MF_GRAYED * -(Not CBool(miTextLen) Or CBool(pLeft(miSelStart, miSelLength) = ZeroL And pRight(miSelStart, miSelLength) = miTextLen)), 6&, "Select &All"
        
        Dim tP As POINT
        Dim liCmd As Long
        
        tP.x = x
        tP.y = y
        ClientToScreen hWnd, tP
        With Screen
            'problems on nt 4.0 when pressing the context menu key while the control is in focus and out of sight
            If tP.x > .Width \ .TwipsPerPixelX Then tP.x = .Width \ .TwipsPerPixelX
            If tP.y > .Height \ .TwipsPerPixelY Then tP.y = .Height \ .TwipsPerPixelY
        End With
        
        liCmd = TrackPopupMenu(lhMenu, TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, tP.x, tP.y, ZeroL, hWnd, ByVal ZeroL)
        
        Select Case liCmd
        Case 1: Undo
        Case 2: Cut
        Case 3: Copy
        Case 4: Paste
        Case 5: pDelete miSelStart, miSelLength
        Case 6: pUpdateSelection ZeroL, miTextLen
        End Select
        
        DestroyMenu lhMenu
    End If
End Sub

Private Function pWordBreakProc(ByVal iDir As Long, ByVal iStart As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Find the nearest literal/non-literal bounding character
'---------------------------------------------------------------------------------------

    If iStart > miTextLen Then iStart = miTextLen
    If iStart < ZeroL Then iStart = ZeroL
    Debug.Assert iDir
    
    Dim lbTermOnLiteral As Boolean
    lbTermOnLiteral = Not pLiteral(iStart)
    
    For pWordBreakProc = iStart To IIf(iDir < ZeroL, ZeroL, miTextLen) Step Sgn(iDir)
        If pLiteral(pWordBreakProc) Then
            If lbTermOnLiteral Then Exit For
        Else
            lbTermOnLiteral = True
        End If
    Next
    
End Function

Private Function pLeft(ByVal iStart As Long, ByVal iLength As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the left point of the selection
'---------------------------------------------------------------------------------------

    If iLength < ZeroL Then pLeft = iStart + iLength Else pLeft = iStart
    'Debug.Assert pLeft <= miTextLen And pLeft > NegOneL
End Function

Private Function pRight(ByVal iStart As Long, ByVal iLength As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the right point of the selection
'---------------------------------------------------------------------------------------
    If iLength > ZeroL Then pRight = iStart + iLength Else pRight = iStart
    'Debug.Assert pRight <= miTextLen And pRight > NegOneL
End Function

Private Sub pInvalidate(ByVal iFrom As Long, ByVal iTo As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Invalidate the text in the specified range.
'---------------------------------------------------------------------------------------

    Dim lhDc As Long
    Dim lhFontOld As Long
    Dim lhWnd As Long
    Dim tR As RECT
    Dim ltSize As SIZE
    Dim liOffset As Long
    
    Dim lsText As String
    lsText = StrConv(pGetText(), vbFromUnicode)
    
    lhWnd = hWnd
    
    If iTo <> iFrom Then
        If mhFont Then
            If iTo < iFrom Then
                iTo = iTo Xor iFrom
                iFrom = iTo Xor iFrom
                iTo = iFrom Xor iTo
            End If
            
            Debug.Assert LenB(lsText) >= iTo
            Debug.Assert iFrom > NegOneL
            
            lhDc = GetDC(lhWnd)
            If lhDc Then
                lhFontOld = SelectObject(lhDc, mhFont)
                If lhFontOld Then
                    If GetTextExtentPoint32(lhDc, ByVal StrPtr(lsText), miFirstVisibleChar, ltSize) Then
                        liOffset = ltSize.cx
                        
                        tR.Top = ScaleHeight \ TwoL - miLineHeight \ TwoL
                        tR.bottom = tR.Top + miLineHeight
                        If GetTextExtentPoint32(lhDc, ByVal StrPtr(lsText), iFrom, ltSize) Then
                            tR.Left = ltSize.cx + miXOffset - liOffset
                            If GetTextExtentPoint32(lhDc, ByVal StrPtr(lsText), iTo, ltSize) Then
                                tR.Right = ltSize.cx + miXOffset - liOffset
                                InvalidateRect lhWnd, tR, ZeroL
                            End If
                        End If
                        SelectObject lhDc, lhFontOld
                    End If
                End If
                ReleaseDC lhWnd, lhDc
            End If
        End If
    End If
End Sub

Private Sub pUpdateSelection(ByVal iStart As Long, ByVal iLength As Long, Optional ByVal bFromArrowKey As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Update the insertion point and selection length, invalidating both old and new.
'---------------------------------------------------------------------------------------
    If iLength Then
    
        If miSelLength Then
            If pLeft(iStart, iLength) <> pLeft(miSelStart, miSelLength) Then pInvalidate pLeft(iStart, iLength), pLeft(miSelStart, miSelLength)
            If pRight(iStart, iLength) <> pRight(miSelStart, miSelLength) Then pInvalidate pRight(iStart, iLength), pRight(miSelStart, miSelLength)
        Else
            pInvalidate iStart, iStart + iLength
        End If
    Else
        pInvalidate miSelStart, miSelStart + miSelLength
    End If
    
    miSelStart = iStart
    miSelLength = iLength
    UpdateWindow hWnd
    pUpdateCaretPos bFromArrowKey
    
End Sub

Private Function pLogicalCharLength(ByRef sMask As String, ByVal iPos As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Return the number of characters in the mask that identify the next physical
'             char.  for example, "*" refers to just one char as does "[A-Z1-3]"
'---------------------------------------------------------------------------------------

    pLogicalCharLength = OneL
    If AscW(Mid$(sMask, iPos, OneL)) = 91 Then ']
        pLogicalCharLength = InStr(iPos, sMask, "]") - iPos + OneL
        If pLogicalCharLength < ZeroL Then pLogicalCharLength = OneL
    End If
End Function

Private Sub pCheckUndo(ByVal iEditType As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/13/05
' Purpose   : Update the undo buffer if a new edit type is used or if the timeout has elapsed.
'---------------------------------------------------------------------------------------
    Static iLastEditType As Long
    
    If TickDiff(GetTickCount(), miUndoTimeStamp) > UNDO_Timeout Or iEditType <> iLastEditType Then
        iLastEditType = iEditType
        msUndoText = pGetText()
        miUndoTimeStamp = GetTickCount()
    End If
End Sub


Public Property Get Themeable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return a value indicating whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        PropertyChanged PROP_Themeable
        mbThemeable = bNew
        pSetTheme
    End If
End Property

Private Function GetCaretWidth() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the default system character width.
'---------------------------------------------------------------------------------------
    SystemParametersInfo SPI_GETCARETWIDTH, ZeroL, GetCaretWidth, ZeroL
    If GetCaretWidth < OneL Then GetCaretWidth = OneL
End Function
