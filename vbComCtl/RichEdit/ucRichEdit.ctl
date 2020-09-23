VERSION 5.00
Begin VB.UserControl ucRichEdit 
   BackColor       =   &H80000005&
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   HasDC           =   0   'False
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ToolboxBitmap   =   "ucRichEdit.ctx":0000
End
Attribute VB_Name = "ucRichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucRicEdit.ctl        12/15/04
'
'           PURPOSE:
'               Implement the Richedit control using either riched20.dll or riched32.dll.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/codelib/richedit/richedit.htm
'               VBRichEdit.ctl
'
'==================================================================================================

Option Explicit

Public Enum eRichEditFormatType
    rtfFormatDefault = SCF_DEFAULT
    rtfFormatSelection = SCF_SELECTION
    rtfFormatAll = SCF_ALL
End Enum

'' /*  UndoName info */
Public Enum eRichEditUndoType
    rtfUnknown = 0
    rtfTyping = 1
    rtfDelete = 2
    rtfDragDrop = 3
    rtfCut = 4
    rtfPaste = 5
End Enum
Public Enum eRichEditSelectionType
    rtfSelEmpty = &H0
    rtfSelText = &H1
    rtfSelObject = &H2
    rtfSelMultiChar = &H4
    rtfSelMultiObject = &H8
End Enum

Public Enum eRichEditParaAlignment
   rtfParaLeft = PFA_LEFT
   rtfParaCenter = PFA_CENTER
   rtfParaRight = PFA_RIGHT
   rtfParaJustify = PFA_FULL_INTERWORD
End Enum

Public Enum eRichEditLinkEvent
    rtfLinkLDblClick = WM_LBUTTONDBLCLK
    rtfLinkLDown = WM_LBUTTONDOWN
    rtfLinkLUp = WM_LBUTTONUP
    rtfLinkMove = WM_MOUSEMOVE
    rtfLinkRDblClick = WM_RBUTTONDBLCLK
    rtfLinkRDown = WM_RBUTTONDOWN
    rtfLinkRUp = WM_RBUTTONUP
    rtfLinkSetCursor = WM_SETCURSOR
End Enum

Public Enum eRichEditScrollBars
   rtfScrollBarsNone = 1
   rtfScrollBarsHorizontal
   rtfScrollBarsVertical
   rtfScrollBarsBoth
End Enum

Public Event SelectionChange(ByVal iMin As Long, ByVal iMax As Long, ByVal iSelType As eRichEditSelectionType)
Public Event LinkEvent(ByVal iType As eRichEditLinkEvent, ByVal iMin As Long, ByVal iMax As Long)
Public Event KeyDown(iKeyCode As Integer, ByVal iState As evbComCtlKeyboardState, ByVal bRepeat As Boolean)
Public Event KeyPress(iKeyAscii As Integer, ByVal iState As evbComCtlKeyboardState, ByVal bRepeat As Boolean)
Public Event KeyUp(iKeyCode As Integer, ByVal iState As evbComCtlKeyboardState)
Public Event MouseDown(ByVal iButton As evbComCtlMouseButton, ByVal iState As evbComCtlKeyboardState, x As Single, y As Single)
Public Event MouseMove(ByVal iButton As evbComCtlMouseButton, ByVal iState As evbComCtlKeyboardState, x As Single, y As Single)
Public Event MouseUp(ByVal iButton As evbComCtlMouseButton, ByVal iState As evbComCtlKeyboardState, x As Single, y As Single)
Public Event MouseDblClick(ByVal iButton As evbComCtlMouseButton, ByVal iState As evbComCtlKeyboardState, x As Single, y As Single)
Public Event StreamInProgress(ByVal fProgress As Single, ByVal fTotal As Single)
Public Event StreamOutProgress(ByVal fProgress As Single)
Public Event ModifyProtected(ByRef bModify As Boolean, ByVal iMin As Long, ByVal iMax As Long)
Public Event ContextMenu(ByVal x As Single, ByVal y As Single)
Public Event VScroll()
Public Event HScroll()
Public Event Change()

Implements iOleInPlaceActiveObjectVB
Implements iOleControlVB
Implements iSubclass
Implements iTimer

Private Const cCharFormat = "cRichEditCharFormat"
'Private Const cParaFormat = "cRichEditParaFormat"
Private Const ucRichEdit = "ucRichEdit"

Private Const NMHDR_hwndFrom = ZeroL
Private Const NMHDR_code = 8&
Private Const SELCHANGE_CHARRANGE_cpMin = 12&
Private Const SELCHANGE_CHARRANGE_cpMax = 16&
Private Const SELCHANGE_seltyp = 20&

Private Const ENLINK_uMsg = 12&
Private Const ENLINK_CHARRANGE_cpMin = 24&
Private Const ENLINK_CHARRANGE_cpMax = 28&

Private Const ENLINK_PROTECTED_cpMin = 24&
Private Const ENLINK_PROTECTED_cpMax = 28&

Private Const MSGFILTER_msg = 12&
Private Const MSGFILTER_wParam = 16&
Private Const MSGFILTER_lParam = 20&

Private Enum eBooleanProps
    bpWantTab = &H10000000
    bpAutoURL = &H20000000
    bpWordWrap = &H40000000
    bpTextOnly = &H80000000
    
    bpMultiLine = ES_MULTILINE
    bpDisableNoScroll = ES_DISABLENOSCROLL
    bpPassword = ES_PASSWORD
    bpVScroll = WS_VSCROLL
    bpHScroll = WS_HSCROLL
    bpNoHideSel = ES_NOHIDESEL
    bpReadOnly = ES_READONLY
    bpAutoVScroll = ES_AUTOVSCROLL
    bpAutoHScroll = ES_AUTOHSCROLL
    bpWantReturn = ES_WANTRETURN
    bpSelectionBar = ES_SELECTIONBAR
    
    bpStyleMask = (bpMultiLine Or bpDisableNoScroll Or bpPassword Or bpVScroll Or bpHScroll Or bpNoHideSel Or bpReadOnly Or bpAutoVScroll Or bpAutoHScroll Or bpWantReturn Or bpSelectionBar)
End Enum

Private Enum eRichEditStreamType
    stFileIn
    stFileOut
    stStringIn
    stStringOut
End Enum

Private mhWnd               As Long
Private miBorderStyle       As evbComCtlBorderStyle

Private mbRedraw            As Boolean

Private miUndoLevels        As Long
Private miPasswordChar      As Integer
Private miBooleanProps      As eBooleanProps
Private miBackColor         As Long
Private miLeftMargin        As Long
Private miRightMargin       As Long

Private miMaxLength         As Long

Private mtCharFormat        As CHARFORMAT2
Private mbLastCharFormatConsistent As Boolean

Private mtParaFormat        As PARAFORMAT2
Private mbLastParaFormatConsistent As Boolean

Private mtCharRange         As CHARRANGE

Private mbGEVersion2        As Boolean
Private mbThemeable         As Boolean

Private Const CharFormatLen  As Long = 60
Private Const CharFormat2Len As Long = 84

Private Const ParaFormatLen  As Long = 156
Private Const ParaFormat2Len As Long = 188

Private Const PROP_BorderStyle      As String = "BorderStyle"
Private Const PROP_UndoLevels       As String = "UndoLevels"
Private Const PROP_PasswordChar     As String = "PassChar"
Private Const PROP_BooleanProps     As String = "BooleanProps"
Private Const PROP_BackColor        As String = "BackColor"
Private Const PROP_Enabled          As String = "Enabled"
Private Const PROP_MaxLen           As String = "MaxLen"
Private Const PROP_RightMargin      As String = "RMargin"
Private Const PROP_LeftMargin       As String = "LMargin"
Private Const PROP_Themeable        As String = "Themeable"

Private Const DEF_BorderStyle       As Long = vbccBorderSunken
Private Const DEF_UndoLevels        As Long = 1
Private Const DEF_PasswordChar      As Integer = 0
Private Const DEF_BooleanProps      As Long = bpAutoURL Or bpMultiLine Or bpWordWrap Or bpAutoVScroll Or bpAutoHScroll Or bpWantReturn 'Or bpSelectionBar
Private Const DEF_Backcolor         As Long = NegOneL
Private Const DEF_Enabled           As Boolean = True
Private Const DEF_MaxLen            As Long = 0
Private Const DEF_RightMargin       As Long = 5
Private Const DEF_LeftMargin        As Long = 5
Private Const DEF_Themeable         As Boolean = True

Private Sub iOleControlVB_OnMnemonic(bHandled As Boolean, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iOleControlVB_GetControlInfo(bHandled As Boolean, iAccelCount As Long, hAccelTable As Long, iFlags As Long)
    If (miBooleanProps And bpWantReturn) Then iFlags = vbccEatsReturn
    bHandled = True
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Forward the keys we want to intercept to the richedit.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Select Case uMsg
        Case WM_KEYDOWN, WM_CHAR, WM_KEYUP
            Select Case wParam And &HFFFF&
            Case vbKeyPageUp To vbKeyDown, vbKeyLeft
                bHandled = True
            Case vbKeyTab
                bHandled = CBool(miBooleanProps And bpWantTab)
                If bHandled And (uMsg <> WM_CHAR) Then lReturn = OneL
            Case vbKeyReturn
                bHandled = CBool(miBooleanProps And bpWantReturn)
         End Select
         If bHandled Then SendMessage mhWnd, uMsg, wParam, lParam
      End Select
   End If
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
    Case WM_SETFOCUS
        vbComCtlTlb.SetFocus mhWnd
    Case WM_KILLFOCUS
        DeActivateIPAO Me
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Handle focus and notifications from the richedit.
'---------------------------------------------------------------------------------------
Dim b As Boolean
Dim iKey As Integer

    Select Case uMsg
    Case WM_NOTIFY
        bHandled = True
        If (MemOffset32(lParam, NMHDR_hwndFrom) = mhWnd) Then
            Select Case MemOffset32(lParam, NMHDR_code)
            Case EN_MSGFILTER
                b = False
                Select Case MemOffset32(lParam, MSGFILTER_msg)
                Case WM_LBUTTONDBLCLK
                    RaiseEvent MouseDblClick(vbccMouseLButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_RBUTTONDBLCLK
                    RaiseEvent MouseDblClick(vbccMouseRButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_LBUTTONDOWN
                    RaiseEvent MouseDown(vbccMouseLButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_RBUTTONDOWN
                    RaiseEvent MouseDown(vbccMouseRButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_MBUTTONDOWN
                    RaiseEvent MouseDown(vbccMouseMButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_LBUTTONUP
                    RaiseEvent MouseUp(vbccMouseLButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_RBUTTONUP
                    RaiseEvent MouseUp(vbccMouseRButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_MBUTTONUP
                    RaiseEvent MouseUp(vbccMouseMButton, GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_MOUSEMOVE
                    RaiseEvent MouseMove(GetMouseButton(MemOffset32(lParam, MSGFILTER_wParam)), GetKBState(MemOffset32(lParam, MSGFILTER_wParam)), pGetX(MemOffset32(lParam, MSGFILTER_lParam)), pGetY(MemOffset32(lParam, MSGFILTER_lParam)))
                Case WM_KEYDOWN
                    iKey = MemOffset16(lParam, MSGFILTER_wParam)
                    RaiseEvent KeyDown(iKey, KBState(), MemOffset32(lParam, MSGFILTER_lParam) And &H40000000)
                    b = Not CBool(iKey)
                    MemOffset16(lParam, MSGFILTER_wParam) = iKey
                Case WM_KEYUP
                    iKey = MemOffset16(lParam, MSGFILTER_wParam)
                    RaiseEvent KeyUp(iKey, KBState())
                    b = Not CBool(iKey)
                    MemOffset16(lParam, MSGFILTER_wParam) = iKey
                Case WM_CHAR
                    iKey = MemOffset16(lParam, MSGFILTER_wParam)
                    RaiseEvent KeyPress(iKey, KBState(), MemOffset32(lParam, MSGFILTER_lParam) And &H40000000)
                    b = Not CBool(iKey)
                    MemOffset16(lParam, MSGFILTER_wParam) = iKey
                Case WM_VSCROLL
                    RaiseEvent VScroll
                Case WM_HSCROLL
                    RaiseEvent HScroll
                Case WM_MOUSEACTIVATE
                    If mhWnd Then
                        If GetFocus() <> mhWnd Then
                            vbComCtlTlb.SetFocus UserControl.hWnd
                            lReturn = MA_NOACTIVATE
                        End If
                    End If
                'Case Else
                    'Debug.Print "Other RichEdit MsgFilter:", MemOffset32(lParam, MSGFILTER_msg)
                End Select
                lReturn = Abs(b)
            Case EN_SELCHANGE
                RaiseEvent SelectionChange(MemOffset32(lParam, SELCHANGE_CHARRANGE_cpMin), MemOffset32(lParam, SELCHANGE_CHARRANGE_cpMax), MemOffset32(lParam, SELCHANGE_seltyp))
            Case EN_LINK
                RaiseEvent LinkEvent(MemOffset32(lParam, ENLINK_uMsg), MemOffset32(lParam, ENLINK_CHARRANGE_cpMin), MemOffset32(lParam, ENLINK_CHARRANGE_cpMax))
            Case EN_PROTECTED
                RaiseEvent ModifyProtected(b, MemOffset32(lParam, ENLINK_PROTECTED_cpMin), MemOffset32(lParam, ENLINK_PROTECTED_cpMax))
                lReturn = CLng(b) + OneL
            Case EN_SETFOCUS
                'Debug.Assert False
                ActivateIPAO Me
            'Case EN_UPDATE
                'Debug.Print "Update"
            'Case Else
                'Debug.Print "Other WM_NOTIFY:", MemOffset32(lParam, NMHDR_code)
            End Select
        End If
        
    Case WM_COMMAND
        Select Case (wParam And &H7FFF0000) \ &H10000
        Case EN_CHANGE
            RaiseEvent Change
        End Select
        
    Case WM_SETFOCUS
        ActivateIPAO Me
    
    Case WM_MOUSEACTIVATE
        If GetFocus() <> mhWnd Then
            vbComCtlTlb.SetFocus UserControl.hWnd
            lReturn = MA_NOACTIVATE
            bHandled = True
        End If
        
    Case WM_CONTEXTMENU
        bHandled = True
        lReturn = ZeroL
        If mhWnd Then lParam = TranslateContextMenuCoords(mhWnd, lParam)
        RaiseEvent ContextMenu(pGetX(lParam), pGetY(lParam))
    
    End Select
    
End Sub

Private Sub iTimer_Proc(ByVal iId As Long, ByVal iElapsed As Long)
    Timer_Remove Me, iId
    On Error Resume Next
    If iId Then
        Dim lsText As String
        If Not pStream(stStringOut, SF_RTF, lsText) Then lsText = vbNullString
    End If
    pCreate
    If LenB(lsText) Then pStream stStringIn, SF_RTF, lsText
    On Error GoTo 0
End Sub

Private Sub UserControl_Initialize()
    LoadShellMod
    mbRedraw = True
    
    mbGEVersion2 = RichEdit_Lib.Init()
    
    If mbGEVersion2 Then
        mtCharFormat.cbSize = CharFormat2Len
        mtParaFormat.cbSize = ParaFormat2Len
    Else
        mtCharFormat.cbSize = CharFormatLen
        mtParaFormat.cbSize = ParaFormatLen
    End If
    
End Sub

Private Sub UserControl_InitProperties()
   
    miBooleanProps = DEF_BooleanProps
    mbThemeable = DEF_Themeable
    
    miBorderStyle = DEF_BorderStyle
    miUndoLevels = DEF_UndoLevels
    miPasswordChar = DEF_PasswordChar
    UserControl.Enabled = DEF_Enabled
    miBackColor = DEF_Backcolor
    miMaxLength = DEF_MaxLen
    miLeftMargin = DEF_LeftMargin
    miRightMargin = DEF_RightMargin
    
    pCreate
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    miBooleanProps = PropBag.ReadProperty(PROP_BooleanProps, DEF_BooleanProps)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    miBackColor = PropBag.ReadProperty(PROP_BackColor, DEF_Backcolor)
    miLeftMargin = PropBag.ReadProperty(PROP_LeftMargin, DEF_LeftMargin)
    miRightMargin = PropBag.ReadProperty(PROP_RightMargin, DEF_RightMargin)
    
    miMaxLength = PropBag.ReadProperty(PROP_MaxLen, DEF_MaxLen)
    UserControl.Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    miPasswordChar = PropBag.ReadProperty(PROP_PasswordChar, DEF_PasswordChar)
    miUndoLevels = PropBag.ReadProperty(PROP_UndoLevels, DEF_UndoLevels)
    miBorderStyle = PropBag.ReadProperty(PROP_BorderStyle, DEF_BorderStyle)
    
    On Error GoTo 0
    
    pCreate
    
End Sub

Private Sub UserControl_Resize()
    If mhWnd Then
        MoveWindow mhWnd, ZeroL, ZeroL, UserControl.Width \ Screen.TwipsPerPixelX, UserControl.Height \ Screen.TwipsPerPixelY, OneL
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_BooleanProps, miBooleanProps, DEF_BooleanProps
    PropBag.WriteProperty PROP_BorderStyle, BorderStyle, DEF_BorderStyle
    PropBag.WriteProperty PROP_UndoLevels, miUndoLevels, DEF_UndoLevels
    PropBag.WriteProperty PROP_PasswordChar, miPasswordChar, DEF_PasswordChar
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_BackColor, miBackColor, DEF_Backcolor
    PropBag.WriteProperty PROP_MaxLen, miMaxLength, DEF_MaxLen
    PropBag.WriteProperty PROP_LeftMargin, miLeftMargin, DEF_LeftMargin
    PropBag.WriteProperty PROP_RightMargin, miRightMargin, DEF_RightMargin
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub UserControl_Terminate()
    Timer_Remove Me, ZeroL
    Timer_Remove Me, OneL
    pDestroy
    ReleaseShellMod
End Sub

Private Sub pPropChanged(ByRef s As String)
    If Ambient.UserMode = False Then PropertyChanged s
End Sub


Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Create the richeditcontrol and the subclass for notification messages.
'---------------------------------------------------------------------------------------
    pDestroy
    
    Dim lsAnsi As String
    lsAnsi = StrConv(IIf(mbGEVersion2, WC_RICHEDIT_20A, WC_RICHEDIT_10A) & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, WS_CHILD Or WS_VISIBLE Or (miBooleanProps And bpStyleMask) Or (WS_DISABLED * (UserControl.Enabled + 1)), ZeroL, ZeroL, UserControl.Width \ Screen.TwipsPerPixelX, UserControl.Height \ Screen.TwipsPerPixelY, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        EnableWindowTheme mhWnd, mbThemeable
    
        If Ambient.UserMode Then
            VTableSubclass_OleControl_Install Me
            VTableSubclass_IPAO_Install Me
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_COMMAND, WM_CONTEXTMENU), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_SETFOCUS, WM_MOUSEACTIVATE), WM_KILLFOCUS
    
            SendMessage mhWnd, EM_SETEVENTMASK, ZeroL, ENM_CHANGE Or ENM_KEYEVENTS Or ENM_MOUSEEVENTS Or ENM_SCROLLEVENTS Or ENM_SCROLL Or ENM_LINK Or ENM_PROTECTED Or ENM_SELCHANGE
    
            If mbGEVersion2 Then
                SendMessage mhWnd, EM_AUTOURLDETECT, Sgn(miBooleanProps And bpAutoURL), ZeroL
            End If
    
            SendMessage mhWnd, EM_SETTARGETDEVICE, ZeroL, Abs(Not (CBool(miBooleanProps And bpWordWrap)))
            SendMessage mhWnd, EM_SETTEXTMODE, IIf(miBooleanProps And bpTextOnly, TM_PLAINTEXT, TM_RICHTEXT), ZeroL
    
            pSetBorder
            pSetUndoLimit miUndoLevels
            SendMessage mhWnd, EM_SETPASSWORDCHAR, miPasswordChar, ZeroL
    
            If TranslateColor(miBackColor) = NegOneL Then
                SendMessage mhWnd, EM_SETBKGNDCOLOR, NegOneL, ZeroL
            Else
                SendMessage mhWnd, EM_SETBKGNDCOLOR, ZeroL, TranslateColor(miBackColor)
            End If
            SendMessage mhWnd, EM_EXLIMITTEXT, ZeroL, miMaxLength
            SendMessage mhWnd, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, miLeftMargin Or (miRightMargin * &H10000)
    
        End If
    End If

End Sub

Public Sub Recreate(Optional ByVal bPreserveText As Boolean)
    Timer_Install Me, -CLng(bPreserveText), 1
End Sub

Private Function pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Destroy the richedit control and subclass.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    If mhWnd Then
        VTableSubclass_OleControl_Remove
        VTableSubclass_IPAO_Remove
        
        SendMessage mhWnd, EM_SETTARGETDEVICE, ZeroL, ZeroL
        
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
    On Error GoTo 0
End Function

Private Sub pSetStyle(ByVal iStyle As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a window style of the richedit control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If bNew Then
            miBooleanProps = miBooleanProps Or iStyle
            SetWindowStyle mhWnd, iStyle, ZeroL
        Else
            miBooleanProps = miBooleanProps And Not iStyle
            SetWindowStyle mhWnd, ZeroL, iStyle
        End If
        SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
        pPropChanged PROP_BooleanProps
    End If
End Sub

Private Property Get pCharFormat(ByVal iType As eRichEditFormatType, ByVal iMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value from the CHARFORMAT/CHARFORMAT2 structure.
'---------------------------------------------------------------------------------------
    With mtCharFormat
        .dwMask = iMask
        
        If mhWnd Then
            SendMessage mhWnd, EM_GETCHARFORMAT, iType, VarPtr(mtCharFormat)
            mbLastCharFormatConsistent = CBool(.dwMask And iMask)
            
            If iMask = CFM_BOLD Then
                pCharFormat = (.dwEffects And CFE_BOLD)
            ElseIf iMask = CFM_ITALIC Then
                pCharFormat = (.dwEffects And CFE_ITALIC)
            ElseIf iMask = CFM_PROTECTED Then
                pCharFormat = (.dwEffects And CFE_PROTECTED)
            ElseIf iMask = CFM_STRIKEOUT Then
                pCharFormat = (.dwEffects And CFE_STRIKEOUT)
            ElseIf iMask = CFM_UNDERLINE Then
                pCharFormat = (.dwEffects And CFE_UNDERLINE)
            ElseIf iMask = CFM_COLOR Then
                pCharFormat = .crTextColor
                If CBool(.dwEffects And CFE_AUTOCOLOR) Then pCharFormat = NegOneL
            ElseIf iMask = CFM_CHARSET Then
                pCharFormat = .bCharSet
            ElseIf iMask = CFM_SIZE Then
                pCharFormat = .yHeight \ Screen.TwipsPerPixelY
            ElseIf iMask = CFM_OFFSET Then
                pCharFormat = .yOffset
            ElseIf iMask = CFM_BACKCOLOR Then
                pCharFormat = .crBackColor
                If CBool(.dwEffects And CFE_AUTOBACKCOLOR) Then pCharFormat = NegOneL
            ElseIf iMask = CFM_LINK Then
                pCharFormat = (.dwEffects And CFE_LINK)
            Else
                Debug.Assert False
                pCharFormat = ZeroL
            End If
            
        End If
        
    End With
End Property
Private Property Let pCharFormat(ByVal iType As eRichEditFormatType, ByVal iMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value in the CHARFORMAT/CHARFORMAT2 structure.
'---------------------------------------------------------------------------------------
    With mtCharFormat
        .dwMask = iMask
        If iMask = CFM_BOLD Then
            If iNew Then .dwEffects = CFE_BOLD Else .dwEffects = ZeroL
        ElseIf iMask = CFM_ITALIC Then
            If iNew Then .dwEffects = CFE_ITALIC Else .dwEffects = ZeroL
        ElseIf iMask = CFM_PROTECTED Then
            If iNew Then .dwEffects = CFE_PROTECTED Else .dwEffects = ZeroL
        ElseIf iMask = CFM_STRIKEOUT Then
            If iNew Then .dwEffects = CFE_STRIKEOUT Else .dwEffects = ZeroL
        ElseIf iMask = CFM_UNDERLINE Then
            If iNew Then .dwEffects = CFE_UNDERLINE Else .dwEffects = ZeroL
        ElseIf iMask = CFM_COLOR Then
            .crTextColor = iNew
            If iNew = NegOneL Then .dwEffects = CFE_AUTOCOLOR Else .dwEffects = ZeroL
        ElseIf iMask = CFM_CHARSET Then
            .bCharSet = iNew
        ElseIf iMask = CFM_SIZE Then
            .yHeight = iNew * Screen.TwipsPerPixelY
        ElseIf iMask = CFM_OFFSET Then
            .yOffset = iNew
        ElseIf iMask = CFM_BACKCOLOR Then
            .crBackColor = iNew
            If iNew = NegOneL Then .dwEffects = CFE_AUTOBACKCOLOR Else .dwEffects = ZeroL
        ElseIf iMask = CFM_LINK Then
            If iNew Then .dwEffects = CFE_LINK Else .dwEffects = ZeroL
        Else
            Debug.Assert False
            .dwMask = ZeroL
        End If
        
        If mhWnd Then
            iMask = SendMessage(mhWnd, EM_SETCHARFORMAT, iType, VarPtr(mtCharFormat))
            Debug.Assert iMask
        End If
        
    End With
End Property

Private Property Get pCharFaceName(ByVal iType As eRichEditFormatType) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the face name from a CHARFORMAT/CHARFORMAT2 structure.
'---------------------------------------------------------------------------------------
    With mtCharFormat
        .dwMask = CFM_FACE
        .szFaceName(0) = 0
        
        If mhWnd Then
            SendMessage mhWnd, EM_GETCHARFORMAT, iType, VarPtr(mtCharFormat)
            mbLastCharFormatConsistent = CBool(.dwMask And CFM_FACE)
            
            lstrToStringA VarPtr(.szFaceName(0)), pCharFaceName
            
        End If
    End With
End Property
Private Property Let pCharFaceName(ByVal iType As eRichEditFormatType, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the face name in a CHARFORMAT/CHARFORMAT2 structure.
'---------------------------------------------------------------------------------------
    With mtCharFormat
        .dwMask = CFM_FACE
        
        Dim ls As String
        Dim liLen As Long
        
        ls = StrConv(sNew, vbFromUnicode) & vbNullChar
        
        liLen = LenB(ls)
        If liLen > LF_FACESIZE Then liLen = LF_FACESIZE
        
        If liLen Then CopyMemory .szFaceName(0), ByVal StrPtr(ls), liLen Else .szFaceName(0) = 0
        
        If mhWnd Then
            SendMessage mhWnd, EM_SETCHARFORMAT, iType, VarPtr(mtCharFormat)
        End If
        
    End With
End Property

Private Property Get pParaFormat(ByVal iMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value from the PARAFORMAT/PARAFORMAT2 structure.
'---------------------------------------------------------------------------------------
    With mtParaFormat
        .dwMask = iMask
        If mhWnd Then
            mbLastParaFormatConsistent = CBool(SendMessage(mhWnd, EM_GETPARAFORMAT, ZeroL, VarPtr(mtParaFormat)) And iMask)
        End If
        
        If iMask = PFM_ALIGNMENT Then
            pParaFormat = .wAlignment
        ElseIf iMask = PFM_NUMBERING Then
            pParaFormat = .wNumbering
        ElseIf iMask = PFM_OFFSET Then
            pParaFormat = .dxOffset
        ElseIf iMask = PFM_RIGHTINDENT Then
            pParaFormat = .dxRightIndent
        ElseIf iMask = PFM_STARTINDENT Then
            pParaFormat = .dxStartIndent
        Else
            Debug.Assert False
        End If
        
    End With
End Property

Private Property Let pParaFormat(ByVal iMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value in the PARAFORMAT/PARAFORMAT2 structure.
'---------------------------------------------------------------------------------------
    With mtParaFormat
        .dwMask = iMask
        
        If iMask = PFM_ALIGNMENT Then
            .wAlignment = iNew
        ElseIf iMask = PFM_NUMBERING Then
            .wNumbering = iNew
        ElseIf iMask = PFM_OFFSET Then
            .dxOffset = iNew
        ElseIf iMask = PFM_RIGHTINDENT Then
            .dxRightIndent = iNew
        ElseIf iMask = PFM_STARTINDENT Then
            .dxStartIndent = iNew
        Else
            Debug.Assert False
            .dwMask = ZeroL
        End If
    
        If mhWnd Then SendMessage mhWnd, EM_SETPARAFORMAT, ZeroL, VarPtr(mtParaFormat)
    End With
End Property

Private Function pGetX(ByVal lParam As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Scale a packed integer to container units.
'---------------------------------------------------------------------------------------
    pGetX = UserControl.ScaleX(loword(lParam), vbPixels, vbContainerPosition)
End Function

Private Function pGetY(ByVal lParam As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Scale a packed integer to container units.
'---------------------------------------------------------------------------------------
    pGetY = UserControl.ScaleX(hiword(lParam), vbPixels, vbContainerPosition)
End Function

Private Function pStream(ByVal iType As eRichEditStreamType, ByVal iFlags As Long, ByRef s As String) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Perform a stream operation to stream to/from a string variable or a file.
'---------------------------------------------------------------------------------------
If mhWnd Then
    
    Dim ltEditStream As EDITSTREAM
    
    Dim loStream As Object
    Dim loFileIn As pcRichEditStream_FileIn
    Dim loFileOut As pcRichEditStream_FileOut
    Dim loStringIn As pcRichEditStream_StringIn
    Dim loStringOut As pcRichEditStream_StringOut
    
    Select Case iType
    Case stFileIn:
        Set loStream = New pcRichEditStream_FileIn: Set loFileIn = loStream
        If Not loFileIn.Init(Me, s) Then gErr vbccOutOfMemory, ucRichEdit
    Case stFileOut:
        Set loStream = New pcRichEditStream_FileOut: Set loFileOut = loStream
        If Not loFileOut.Init(Me, s) Then gErr vbccOutOfMemory, ucRichEdit
    Case stStringIn:
        Set loStream = New pcRichEditStream_StringIn: Set loStringIn = loStream
        If Not loStringIn.Init(Me, s) Then gErr vbccOutOfMemory, ucRichEdit
    Case stStringOut:
        Set loStream = New pcRichEditStream_StringOut: Set loStringOut = loStream
        If Not loStringOut.Init(Me) Then gErr vbccOutOfMemory, ucRichEdit
    Case Else:          Debug.Assert False
    End Select
    
    Debug.Assert Not loStream Is Nothing
    
    If Not loStream Is Nothing Then
        
        With ltEditStream
            .dwCookie = ObjPtr(loStream)
            .pfnCallback = Thunk_Alloc(tnkRichEditProc)
            If .pfnCallback = ZeroL Then gErr vbccOutOfMemory, ucRichEdit
        End With
        
        SendMessage mhWnd, IIf(iType = stFileIn Or iType = stStringIn, EM_STREAMIN, EM_STREAMOUT), iFlags, VarPtr(ltEditStream)
        pStream = (ltEditStream.dwError = ZeroL)
        MemFree ltEditStream.pfnCallback
        If iType = stStringOut Then loStringOut.GetStreamResult s
        If Not pStream Then Err.Raise ltEditStream.dwError, ucRichEdit, "Stream operation failed."
        
    End If
    
End If

Debug.Assert pStream

End Function

Friend Sub fStream_InProgress(ByVal iProgress As Long, ByRef iTotal As Long)
    RaiseEvent StreamInProgress(iProgress, iTotal)
End Sub

Friend Sub fStream_OutProgress(ByVal iProgress As Long)
    RaiseEvent StreamOutProgress(iProgress)
End Sub


Friend Property Get fChar_Bold(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the bold attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Bold = pCharFormat(iType, CFM_BOLD)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Bold(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the bold attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_BOLD) = bNew
End Property

Friend Property Get fChar_Italic(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the Italic attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Italic = pCharFormat(iType, CFM_ITALIC)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Italic(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the Italic attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_ITALIC) = bNew
End Property

Friend Property Get fChar_Underline(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the Underline attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Underline = pCharFormat(iType, CFM_UNDERLINE)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Underline(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the Underline attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_UNDERLINE) = bNew
End Property

Friend Property Get fChar_Strikeout(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the Strikeout attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Strikeout = pCharFormat(iType, CFM_STRIKEOUT)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Strikeout(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the Strikeout attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_STRIKEOUT) = bNew
End Property

Friend Property Get fChar_FaceName(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the facename of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_FaceName = pCharFaceName(iType)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_FaceName(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the facename of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFaceName(iType) = sNew
End Property

Friend Property Get fChar_Height(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the Height of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Height = pCharFormat(iType, CFM_SIZE)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Height(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the Height of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_SIZE) = iNew
End Property

Friend Property Get fChar_Offset(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the Offset of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Offset = pCharFormat(iType, CFM_OFFSET)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Offset(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the Offset of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_OFFSET) = iNew
End Property

Friend Property Get fChar_ColorFore(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the forecolor of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_ColorFore = pCharFormat(iType, CFM_COLOR)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_ColorFore(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the forecolor of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_COLOR) = iNew
End Property

Friend Property Get fChar_Protected(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the protected attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
    
    fChar_Protected = pCharFormat(iType, CFM_PROTECTED)
    bConsistent = mbLastCharFormatConsistent
End Property
Friend Property Let fChar_Protected(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the protected attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    pCharFormat(iType, CFM_PROTECTED) = bNew
End Property

Friend Property Get fChar_ColorBack(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the backcolor of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
        fChar_ColorBack = pCharFormat(iType, CFM_BACKCOLOR)
        bConsistent = mbLastCharFormatConsistent
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property
Friend Property Let fChar_ColorBack(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the backcolor of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        pCharFormat(iType, CFM_BACKCOLOR) = iNew
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property

Friend Property Get fChar_Link(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the link attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If iType = rtfFormatAll Then gErr vbccUnsupported, cCharFormat
        fChar_Link = pCharFormat(iType, CFM_LINK)
        bConsistent = mbLastCharFormatConsistent
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property
Friend Property Let fChar_Link(ByVal iType As eRichEditFormatType, ByRef bConsistent As Boolean, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the link attribute of the given CHARFORMAT.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        pCharFormat(iType, CFM_LINK) = CLng(bNew)
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property

Friend Property Get fPara_Alignment(ByRef bConsistent As Boolean) As eRichEditParaAlignment
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the alignment of the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    fPara_Alignment = pParaFormat(PFM_ALIGNMENT)
    bConsistent = mbLastParaFormatConsistent
End Property
Friend Property Let fPara_Alignment(ByRef bConsistent As Boolean, ByVal iNew As eRichEditParaAlignment)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the alignment of the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    pParaFormat(PFM_ALIGNMENT) = iNew
End Property

Friend Property Get fPara_Indent(ByRef bConsistent As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the Indentation of the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    fPara_Indent = pParaFormat(PFM_STARTINDENT)
    bConsistent = mbLastParaFormatConsistent
End Property
Friend Property Let fPara_Indent(ByRef bConsistent As Boolean, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the indentation of the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    pParaFormat(PFM_STARTINDENT) = iNew
End Property

Friend Property Get fPara_HangingIndent(ByRef bConsistent As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hanging indentation of the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    fPara_HangingIndent = pParaFormat(PFM_OFFSET)
    bConsistent = mbLastParaFormatConsistent
End Property
Friend Property Let fPara_HangingIndent(ByRef bConsistent As Boolean, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the hanging indentation of the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    pParaFormat(PFM_OFFSET) = iNew
End Property

Friend Property Get fPara_RightIndent(ByRef bConsistent As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the indentation from the right of the control for the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    fPara_RightIndent = pParaFormat(PFM_RIGHTINDENT)
    bConsistent = mbLastParaFormatConsistent
End Property
Friend Property Let fPara_RightIndent(ByRef bConsistent As Boolean, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the indentation from the right of the control for the selected PARAFORMAT.
'---------------------------------------------------------------------------------------
    pParaFormat(PFM_RIGHTINDENT) = iNew
End Property

Friend Function fPrint_DoPrint(ByRef tFormatRange As FORMATRANGE, bCallStartEndPage As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Do a EM_FORMATRANGE operation.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liPos As Long
        
        If bCallStartEndPage Then StartPage tFormatRange.hDc
        liPos = SendMessage(mhWnd, EM_FORMATRANGE, OneL, VarPtr(tFormatRange))

        fPrint_DoPrint = liPos >= tFormatRange.chrg.cpMin
        
        If bCallStartEndPage Then EndPage tFormatRange.hDc
        
        If fPrint_DoPrint Then
            tFormatRange.chrg.cpMin = liPos
            tFormatRange.chrg.cpMax = liPos + OneL
            fPrint_DoPrint = (SendMessage(mhWnd, EM_FORMATRANGE, ZeroL, VarPtr(tFormatRange)) >= tFormatRange.chrg.cpMin)
            tFormatRange.chrg.cpMax = NegOneL
        End If
    End If
End Function

Friend Sub fPrint_Terminate(ByRef tFormatRange As FORMATRANGE, ByVal bCallStartEndDoc As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Release memory allocated by the richedit for printing operations.
'---------------------------------------------------------------------------------------
    If bCallStartEndDoc Then
        EndDoc tFormatRange.hDc
    End If
    If mhWnd Then SendMessage mhWnd, EM_FORMATRANGE, ZeroL, ZeroL
End Sub


Public Property Get ScrollBars() As eRichEditScrollBars
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an value indicating which scrollbars are visible.
'---------------------------------------------------------------------------------------
    If CBool(miBooleanProps And bpVScroll) And CBool(miBooleanProps And bpHScroll) Then
        ScrollBars = rtfScrollBarsBoth
    ElseIf CBool(miBooleanProps And bpVScroll) Then
        ScrollBars = rtfScrollBarsVertical
    ElseIf CBool(miBooleanProps And bpHScroll) Then
        ScrollBars = rtfScrollBarsHorizontal
    Else
        ScrollBars = rtfScrollBarsNone
    End If
End Property
Public Property Let ScrollBars(ByVal iNew As eRichEditScrollBars)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set which scrollbars are visible.
'---------------------------------------------------------------------------------------
    pSetStyle bpVScroll Or bpHScroll, False
    If iNew = rtfScrollBarsBoth Then
        pSetStyle bpVScroll Or bpHScroll, True
    ElseIf iNew = rtfScrollBarsVertical Then
        pSetStyle bpVScroll, True
    ElseIf iNew = rtfScrollBarsHorizontal Then
        pSetStyle bpHScroll, True
    End If
End Property

Public Property Get DisableNoScroll() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an value indicating whether scrollbars are disabled or hidden when not available.
'---------------------------------------------------------------------------------------
    DisableNoScroll = CBool(miBooleanProps And bpDisableNoScroll)
End Property
Public Property Let DisableNoScroll(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether scrollbars are disabled or hidden when not available.
'---------------------------------------------------------------------------------------
    pSetStyle bpDisableNoScroll, bNew
End Property

Public Property Get HideSelection() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the selection is hidden when not in focus.
'---------------------------------------------------------------------------------------
    HideSelection = Not CBool(miBooleanProps And bpNoHideSel)
End Property
Public Property Let HideSelection(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the selection is hidden when not in focus.
'---------------------------------------------------------------------------------------
    pSetStyle bpNoHideSel, Not bNew
    If mhWnd Then SendMessage mhWnd, EM_HIDESELECTION, Abs(bNew), OneL
End Property

Public Property Get MultiLine() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether multiple lines are allowed in the control.
'---------------------------------------------------------------------------------------
    MultiLine = CBool(miBooleanProps And bpMultiLine)
End Property
Public Property Let MultiLine(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether multiple lines are allowed in the control.
'---------------------------------------------------------------------------------------
    pSetStyle bpMultiLine, bNew
End Property

Public Property Get ReadOnly() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the user can edit the text in the control.
'---------------------------------------------------------------------------------------
    ReadOnly = CBool(miBooleanProps And bpReadOnly)
End Property
Public Property Let ReadOnly(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the user can edit the text in the control.
'---------------------------------------------------------------------------------------
    pSetStyle bpReadOnly, bNew
    If mhWnd Then
        SendMessage mhWnd, EM_SETREADONLY, Abs(bNew), ZeroL
    End If
End Property

Public Property Get AutoVScroll() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the control is scolled in response to selection changes.
'---------------------------------------------------------------------------------------
    AutoVScroll = CBool(miBooleanProps And bpAutoVScroll)
End Property
Public Property Let AutoVScroll(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the control is scolled in response to selection changes.
'---------------------------------------------------------------------------------------
    pSetStyle bpAutoVScroll, bNew
End Property

Public Property Get AutoHScroll() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the control is scolled in response to selection changes.
'---------------------------------------------------------------------------------------
    AutoHScroll = CBool(miBooleanProps And bpAutoHScroll)
End Property
Public Property Let AutoHScroll(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the control is scolled in response to selection changes.
'---------------------------------------------------------------------------------------
    pSetStyle bpAutoHScroll, bNew
End Property

Public Property Get AutoURLDetect() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether urls are identified during a stream in operation.
'---------------------------------------------------------------------------------------
    AutoURLDetect = CBool(miBooleanProps And bpAutoURL)
End Property
Public Property Let AutoURLDetect(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether urls are identified during a stream in operation.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If bNew Then
            miBooleanProps = miBooleanProps Or bpAutoURL
        Else
            miBooleanProps = miBooleanProps And Not bpAutoURL
        End If
        If mhWnd Then
            SendMessage mhWnd, EM_AUTOURLDETECT, Abs(bNew), ZeroL
        End If
        pPropChanged PROP_BooleanProps
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property

Public Property Get WantReturn() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the ENTER key is captured.
'---------------------------------------------------------------------------------------
    WantReturn = CBool(miBooleanProps And bpWantReturn)
End Property
Public Property Let WantReturn(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the ENTER key is captured.
'---------------------------------------------------------------------------------------
    pSetStyle bpWantReturn, bNew
    OnControlInfoChanged Me
    ActivateIPAO Me
End Property

Public Property Get WantTab() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the TAB key is captured.
'---------------------------------------------------------------------------------------
    WantTab = CBool(miBooleanProps And bpWantTab)
End Property
Public Property Let WantTab(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the TAB key is captured.
'---------------------------------------------------------------------------------------
    If bNew Then
        miBooleanProps = miBooleanProps Or bpWantTab
    Else
        miBooleanProps = miBooleanProps And Not bpWantTab
    End If
    pPropChanged PROP_BooleanProps
End Property
    
Public Property Get SelectionBar() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether there is a small white space on the left
'             edge of the control that allows you to select text a line at a time.
'---------------------------------------------------------------------------------------
    SelectionBar = CBool(miBooleanProps And bpSelectionBar)
End Property
Public Property Let SelectionBar(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether there is a small white space on the left edge of the
'             control that allows you to select text a line at a time.
'---------------------------------------------------------------------------------------
    pSetStyle bpSelectionBar, bNew
End Property

Public Property Get PasswordChar() As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the password character displayed instead of all other characters.
'---------------------------------------------------------------------------------------
   If miPasswordChar <> 0 Then PasswordChar = ChrW$(miPasswordChar)
End Property
Public Property Let PasswordChar(ByVal sChar As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the password character displayed instead of all other characters.
'---------------------------------------------------------------------------------------
    If LenB(sChar) Then miPasswordChar = AscW(sChar) Else miPasswordChar = 0
    If mhWnd Then
        pSetStyle bpPassword, CBool(miPasswordChar)
        SendMessage mhWnd, EM_SETPASSWORDCHAR, miPasswordChar, ZeroL
        pPropChanged PROP_PasswordChar
    End If
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the control responds to user input.
'---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the control responds to user input.
'---------------------------------------------------------------------------------------
    UserControl.Enabled = bNew
    If mhWnd Then EnableWindow mhWnd, -CLng(bNew)
    SetWindowStyle mhWnd, WS_DISABLED * (bNew + 1), WS_DISABLED
    pPropChanged PROP_Enabled
End Property

Public Property Get MaxLength() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the maximum length of text that the user is allowed to input.
'---------------------------------------------------------------------------------------
   If mhWnd Then
      MaxLength = SendMessage(mhWnd, EM_GETLIMITTEXT, ZeroL, ZeroL)
   End If
End Property
Public Property Let MaxLength(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the maximum length of text that the user is allowed to input.
'---------------------------------------------------------------------------------------
   If mhWnd Then
      If iNew < ZeroL Then iNew = ZeroL
      SendMessage mhWnd, EM_EXLIMITTEXT, ZeroL, iNew
      miMaxLength = iNew
      pPropChanged PROP_MaxLen
   End If
End Property

Public Property Get BorderStyle() As evbComCtlBorderStyle
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the control's border style.
'---------------------------------------------------------------------------------------
   BorderStyle = miBorderStyle
End Property
Public Property Let BorderStyle(ByVal iNew As evbComCtlBorderStyle)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Change the border and refresh the control.
'---------------------------------------------------------------------------------------
    miBorderStyle = iNew
    pSetBorder
    pPropChanged PROP_BorderStyle
End Property

Private Sub pSetBorder()
    If mhWnd Then
        SetWindowStyle mhWnd, ZeroL, WS_BORDER
        SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
        Select Case miBorderStyle
        Case vbccBorderSunken
            SetWindowStyleEx mhWnd, WS_EX_CLIENTEDGE, ZeroL
        Case vbccBorderThin
            SetWindowStyleEx mhWnd, WS_EX_STATICEDGE, ZeroL
        Case vbccBorderSingle
            SetWindowStyle mhWnd, WS_BORDER, ZeroL
        End Select
        SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
    End If
End Sub

Public Property Get ColorBack() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the backcolor of the control.
'---------------------------------------------------------------------------------------
    ColorBack = miBackColor
End Property
Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Change the backcolor of the control.
'---------------------------------------------------------------------------------------
    miBackColor = iNew
    If mhWnd Then
        iNew = TranslateColor(iNew)
        If iNew = NegOneL Then
            SendMessage mhWnd, EM_SETBKGNDCOLOR, iNew, ZeroL
        Else
            SendMessage mhWnd, EM_SETBKGNDCOLOR, ZeroL, iNew
        End If
    End If
    pPropChanged PROP_BackColor
End Property

Public Property Let RightMargin(ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the margin on the right of the control.
'---------------------------------------------------------------------------------------
    miRightMargin = ScaleX(fNew, vbContainerSize, vbPixels)
    miRightMargin = miRightMargin And &H7FFF&
    If mhWnd Then
        SendMessage mhWnd, EM_SETMARGINS, EC_RIGHTMARGIN Or EC_LEFTMARGIN, (miRightMargin * &H10000) Or miLeftMargin
        InvalidateRect mhWnd, ByVal ZeroL, OneL
    End If
    pPropChanged PROP_RightMargin
End Property
Public Property Get RightMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the margin on the right of the control.
'---------------------------------------------------------------------------------------
    RightMargin = ScaleX(miRightMargin, vbPixels, vbContainerSize)
End Property
Public Property Let LeftMargin(ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the margin on the left of the control.
'---------------------------------------------------------------------------------------
    miLeftMargin = ScaleX(fNew, vbContainerSize, vbPixels)
    miLeftMargin = miLeftMargin And &H7FFF&
    If mhWnd Then
        SendMessage mhWnd, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, miLeftMargin Or (miRightMargin * &H10000)
        InvalidateRect mhWnd, ByVal ZeroL, OneL
    End If
    pPropChanged PROP_LeftMargin
End Property
Public Property Get LeftMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the margin on the left of the control.
'---------------------------------------------------------------------------------------
    LeftMargin = ScaleX(miLeftMargin, vbPixels, vbContainerSize)
End Property

Public Property Get Modified() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the text has been modified since the
'             modified property was last set to false.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Modified = CBool(SendMessage(mhWnd, EM_GETMODIFY, ZeroL, ZeroL))
    End If
End Property
Public Property Let Modified(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the modified flag to a given value.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, EM_SETMODIFY, Abs(bNew), ZeroL
    End If
End Property

Public Property Get CharFormat(ByVal iType As eRichEditFormatType) As cRichEditCharFormat
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cRichEditCharFormat object representing the given range of characters.
'---------------------------------------------------------------------------------------
    Set CharFormat = New cRichEditCharFormat
    CharFormat.fInit Me, iType And (rtfFormatSelection Or rtfFormatDefault Or rtfFormatAll)
End Property

Public Property Get ParaFormat() As cRichEditParaFormat
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cRichEditParaFormat object representing the selection.
'---------------------------------------------------------------------------------------
    Set ParaFormat = New cRichEditParaFormat
    ParaFormat.fInit Me
End Property

Public Property Get NewPrintJob(ByVal hDc As Long, Optional ByRef sDocTitle As String, Optional ByVal bCallStartEndDoc As Boolean = True, Optional ByVal bCallStartEndPage As Boolean = True) As cRichEditPrintJob
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cRichEditPrintJob object that can be used to print the contents
'             of the control.
'---------------------------------------------------------------------------------------
    Set NewPrintJob = New cRichEditPrintJob
    NewPrintJob.fInit Me, hDc, bCallStartEndDoc, bCallStartEndPage
    If bCallStartEndDoc Then
        Dim ltDI As DOCINFO
        Dim lsTitle As String
        
        lsTitle = StrConv(sDocTitle & vbNullChar, vbFromUnicode)
        
        ltDI.cbSize = LenB(ltDI)
        ltDI.lpszDocName = StrPtr(lsTitle)
        StartDoc hDc, ltDI
    End If
End Property

Public Property Get FindText(ByRef sText As String, Optional ByVal bWholeWord As Boolean, Optional ByVal bMatchCase As Boolean, Optional ByVal iMin As Long = ZeroL, Optional ByVal iMax As Long = NegOneL) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Search for the given text and return its character position if found.
'---------------------------------------------------------------------------------------
If mhWnd Then
    Dim liOptions As Long
    liOptions = FR_DOWN Or _
        (-bWholeWord * FR_WHOLEWORD) Or _
        (-bMatchCase * FR_MATCHCASE)
    
    Dim tFTE As FINDTEXTEX
    With tFTE
        .chrg.cpMin = iMin
        .chrg.cpMax = iMax
        .chrgText.cpMin = NegOneL
        .chrgText.cpMax = NegOneL
        Dim ls As String: ls = StrConv(sText, vbFromUnicode) & vbNullChar
        .lpstrText = StrPtr(ls)
    End With
    
    FindText = SendMessage(mhWnd, EM_FINDTEXTEX, liOptions, VarPtr(tFTE))
    Debug.Assert FindText = NegOneL Or ((tFTE.chrgText.cpMax - tFTE.chrgText.cpMin) = Len(sText))
    
End If
End Property

Public Property Get Text(Optional ByVal bRTF As Boolean, Optional ByVal bSelectionOnly As Boolean) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the text of the control.
'---------------------------------------------------------------------------------------
    pStream stStringOut, (Abs(bRTF) + OneL) Or (-bSelectionOnly * SFF_SELECTION), Text
End Property

Public Property Let Text(Optional ByVal bRTF As Boolean, Optional ByVal bSelectionOnly As Boolean, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of the control.
'---------------------------------------------------------------------------------------
    pStream stStringIn, (Abs(bRTF) + OneL) Or (-bSelectionOnly * SFF_SELECTION), sNew
End Property

Public Function SaveFile(ByRef sFile As String, Optional ByVal bRTF As Boolean, Optional ByVal bSelectionOnly As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Save the text of the control to a file.
'---------------------------------------------------------------------------------------
    SaveFile = pStream(stFileOut, (Abs(bRTF) + OneL) Or (-bSelectionOnly * SFF_SELECTION), sFile)
End Function

Public Function LoadFile(ByRef sFile As String, Optional ByVal bRTF As Boolean, Optional ByVal bSelectionOnly As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Load the text of a file into the control.
'---------------------------------------------------------------------------------------
    LoadFile = pStream(stFileIn, (Abs(bRTF) + OneL) Or (-bSelectionOnly * SFF_SELECTION), sFile)
End Function

Public Property Get LineCount() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of lines of text in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        LineCount = SendMessage(mhWnd, EM_GETLINECOUNT, ZeroL, ZeroL)
    End If
End Property

Public Property Get FirstVisibleLine() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the index of the first visible line.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        FirstVisibleLine = SendMessage(mhWnd, EM_GETFIRSTVISIBLELINE, ZeroL, ZeroL)
    End If
End Property

Public Property Get LineForCharacter(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the line number of a character index.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        LineForCharacter = SendMessage(mhWnd, EM_EXLINEFROMCHAR, ZeroL, iIndex)
    End If
End Property

Public Property Get CharFromPos(ByVal x As Single, ByVal y As Single) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the character index from the given point.
'---------------------------------------------------------------------------------------
    Dim tP As POINT
    tP.x = ScaleX(x, vbContainerPosition, vbPixels)
    tP.y = ScaleY(y, vbContainerPosition, vbPixels)
    If mhWnd Then CharFromPos = SendMessage(mhWnd, EM_CHARFROMPOS, ZeroL, VarPtr(tP))
End Property

Public Sub GetPosFromChar(ByVal iIndex As Long, ByRef x As Single, ByRef y As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the position of the given character index.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        iIndex = SendMessage(mhWnd, EM_POSFROMCHAR, iIndex, ZeroL)
        x = ScaleX((iIndex And &HFFFF&), vbPixels, vbContainerPosition)
        y = ScaleY((iIndex \ &H10000) And &HFFFF&, vbPixels, vbContainerPosition)
    End If
End Sub

Public Property Get CanPaste() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether valid data exists to paste into the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        CanPaste = SendMessage(mhWnd, EM_CANPASTE, ZeroL, ZeroL)
    End If
End Property
Public Property Get CanCopy() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the selection state allows a copy operation.
'---------------------------------------------------------------------------------------
    CanCopy = CBool(SelEnd > SelStart)
End Property
Public Property Get CanUndo() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether there is anything in the undo buffer.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        CanUndo = SendMessage(mhWnd, EM_CANUNDO, ZeroL, ZeroL)
    End If
End Property
Public Property Get CanRedo() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether there is anything in the redo buffer.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If mhWnd Then
            CanRedo = SendMessage(mhWnd, EM_CANREDO, ZeroL, ZeroL)
        End If
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property
Public Property Get UndoType() As eRichEditUndoType
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating what type of operation is next in the undo buffer.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If mhWnd Then
            UndoType = SendMessage(mhWnd, EM_GETUNDONAME, ZeroL, ZeroL)
        End If
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property
Public Property Get RedoType() As eRichEditUndoType
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating what type of operation is next in the redo buffer.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If mhWnd Then
            RedoType = SendMessage(mhWnd, EM_GETREDONAME, ZeroL, ZeroL)
        End If
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Property
Public Sub Cut()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Copy the selected text to the clipboard and delete it from the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, WM_CUT, ZeroL, ZeroL
    End If
End Sub
Public Sub Copy()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Copy the selected text to the clipboard.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, WM_COPY, ZeroL, ZeroL
    End If
End Sub
Public Sub Paste()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Paste the data from the clipboard into the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, WM_PASTE, ZeroL, ZeroL
    End If
End Sub
Public Sub PasteSpecial()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Ask the richedit to do a paste special operation.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, EM_PASTESPECIAL, ZeroL, ZeroL
    End If
End Sub
Public Sub Undo()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Undo the last operation.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, EM_UNDO, ZeroL, ZeroL
    End If
End Sub
Public Sub Redo()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Redo the last operation that was undone.
'---------------------------------------------------------------------------------------
    If mbGEVersion2 Then
        If mhWnd Then
            SendMessage mhWnd, EM_REDO, ZeroL, ZeroL
        End If
    Else
        gErr vbccUnsupported, ucRichEdit
    End If
End Sub

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the lowest character index that is selected.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, EM_EXGETSEL, ZeroL, VarPtr(mtCharRange)
        SelStart = mtCharRange.cpMin
    End If
End Property
Public Property Let SelStart(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the lowest character index that is selected.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        mtCharRange.cpMax = SelEnd
        mtCharRange.cpMin = iNew
        If mtCharRange.cpMax < mtCharRange.cpMin Then mtCharRange.cpMax = mtCharRange.cpMin
        SendMessage mhWnd, EM_EXSETSEL, ZeroL, VarPtr(mtCharRange)
    End If
End Property

Public Property Get SelEnd() As Long
Attribute SelEnd.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the highest character index that is selected.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, EM_EXGETSEL, ZeroL, VarPtr(mtCharRange)
        SelEnd = mtCharRange.cpMax
    End If
End Property
Public Property Let SelEnd(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the highest character index that is selected.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        mtCharRange.cpMin = SelStart
        mtCharRange.cpMax = iNew
        If mtCharRange.cpMin > mtCharRange.cpMax Then mtCharRange.cpMin = mtCharRange.cpMax
        SendMessage mhWnd, EM_EXSETSEL, ZeroL, VarPtr(mtCharRange)
    End If
End Property

Public Sub SetSelection(ByVal iStart As Long, Optional ByVal iLength As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the current selection.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        mtCharRange.cpMin = iStart
        mtCharRange.cpMax = iStart + iLength
        SendMessage mhWnd, EM_EXSETSEL, ZeroL, VarPtr(mtCharRange)
    End If
End Sub

Public Property Get TextRange(Optional ByVal iStart As Long, Optional ByVal iEnd As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the text from the given range.
'---------------------------------------------------------------------------------------
    If CBool(mhWnd) And iStart <> iEnd Then
        If iStart > iEnd Then gErr vbccInvalidProcedureCall, ucRichEdit
        
        Dim tR As TextRange
        Dim lR As Long
        
        tR.chrg.cpMin = iStart
        tR.chrg.cpMax = iEnd
        
        lR = iEnd - iStart + OneL
        
        If lR > ZeroL Then
            TextRange = Space$(lR)
            Mid$(TextRange, lR - OneL, OneL) = vbNullChar
        End If
        
        tR.lpstrText = StrPtr(TextRange)
        
        lR = SendMessage(mhWnd, EM_GETTEXTRANGE, ZeroL, VarPtr(tR))
        If (lR > ZeroL) Then TextRange = StrConv(LeftB$(TextRange, lR), vbUnicode) Else TextRange = vbNullString
    End If
End Property

Public Property Get WordWrap() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether text is wrapped at the right edge of the control.
'---------------------------------------------------------------------------------------
    WordWrap = CBool(miBooleanProps And bpWordWrap)
End Property
Public Property Let WordWrap(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether text is wrapped at the right edge of the control.
'---------------------------------------------------------------------------------------
    If bNew Then
        miBooleanProps = miBooleanProps Or bpWordWrap
    Else
        miBooleanProps = miBooleanProps And Not bpWordWrap
    End If
    pPropChanged PROP_BooleanProps
    If mhWnd Then SendMessage mhWnd, EM_SETTARGETDEVICE, ZeroL, Abs(Not (CBool(miBooleanProps And bpWordWrap)))
End Property

Public Property Get CharacterCount() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of characters in the control.
'---------------------------------------------------------------------------------------
If mhWnd Then
    CharacterCount = SendMessage(mhWnd, WM_GETTEXTLENGTH, ZeroL, ZeroL)
End If
End Property

Public Property Get PlainText() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether the control allows only one format.
'---------------------------------------------------------------------------------------
    PlainText = CBool(miBooleanProps And bpTextOnly)
End Property
Public Property Let PlainText(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the control allows only one format.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        'If bNew Then
        '    miBooleanProps = miBooleanProps Or bpTextOnly
        'Else
        '    miBooleanProps = miBooleanProps And Not bpTextOnly
        'End If
        
        pSetStyle bpTextOnly, bNew
        
        If bNew Then
            SendMessage mhWnd, EM_SETTEXTMODE, TM_PLAINTEXT, ZeroL
        Else
            SendMessage mhWnd, EM_SETTEXTMODE, TM_RICHTEXT, ZeroL
        End If
        
        pPropChanged PROP_BooleanProps
        
    End If
End Property

Public Property Get UndoLevels() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of undo levels that are stored.
'---------------------------------------------------------------------------------------
    UndoLevels = miUndoLevels
End Property
Public Property Let UndoLevels(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the number of undo levels that are stored.
'---------------------------------------------------------------------------------------
    pSetUndoLimit iNew
End Property

Private Sub pSetUndoLimit(ByVal iNew As Long)
    If mhWnd Then
        If iNew < OneL Then
            SendMessage mhWnd, EM_SETTEXTMODE, TM_SINGLELEVELUNDO, ZeroL
            miUndoLevels = SendMessage(mhWnd, EM_SETUNDOLIMIT, ZeroL, ZeroL)
        ElseIf iNew = OneL Then
            SendMessage mhWnd, EM_SETTEXTMODE, TM_SINGLELEVELUNDO, ZeroL
            miUndoLevels = SendMessage(mhWnd, EM_SETUNDOLIMIT, iNew, ZeroL)
        Else
            SendMessage mhWnd, EM_SETTEXTMODE, TM_MULTILEVELUNDO, ZeroL
            miUndoLevels = SendMessage(mhWnd, EM_SETUNDOLIMIT, iNew, ZeroL)
        End If
        pPropChanged PROP_UndoLevels
    End If
End Sub

Public Property Get Redraw() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether the control is redrawn as it is changed.
'---------------------------------------------------------------------------------------
   Redraw = mbRedraw
End Property
Public Property Let Redraw(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the control is redrawn as it is changed.
'---------------------------------------------------------------------------------------
    If (mbRedraw Xor bNew) Then
        If mhWnd Then
            SendMessage mhWnd, WM_SETREDRAW, Abs(bNew), ZeroL
            If bNew Then
                InvalidateRect mhWnd, ByVal ZeroL, OneL
                UpdateWindow mhWnd
            End If
        End If
        mbRedraw = bNew
    End If
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndRichEdit() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hwnd of the richedit
'---------------------------------------------------------------------------------------
    If mhWnd Then
        hWndRichEdit = mhWnd
    End If
End Property

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
        mbThemeable = bNew
        If mhWnd Then
            EnableWindowTheme mhWnd, mbThemeable
            SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
            RedrawWindow mhWnd, ByVal ZeroL, ZeroL, RDW_INVALIDATE
        End If
        pPropChanged PROP_Themeable
    End If
End Property
