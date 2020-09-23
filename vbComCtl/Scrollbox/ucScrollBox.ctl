VERSION 5.00
Begin VB.UserControl ucScrollBox 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucScrollBox.ctx":0000
End
Attribute VB_Name = "ucScrollBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'ucScrollBox.ctl        4/15/05
'
'             PURPOSE:
'               Provide a client area where controls may
'               be placed that can be scrolled automatically.
'
'---------------------------------------------------------------------------------------
Option Explicit

Event ScrollBarChange()
Event Resize()

Implements iSubclass

Private Enum eScrollBar
    scrHorizontal = SB_HORZ
    scrVertical = SB_VERT
End Enum

Private Type tScrollState
    bVisible(SB_HORZ To SB_VERT) As Boolean
    iTotalOffset(SB_HORZ To SB_VERT) As Long
    iSmallChange(SB_HORZ To SB_VERT) As Long
    iMax(SB_HORZ To SB_VERT) As Long
    iLargeChange(SB_HORZ To SB_VERT) As Long
End Type

Const PROP_Themeable = "Themeable"
Const PROP_BackColor = "Backcolor"

Const DEF_Themeable = False
Const DEF_Backcolor = vbButtonFace

Private mtState As tScrollState
Private mbThemeable As Boolean

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    bHandled = True
    
    Dim liBar As eScrollBar
    
    Select Case uMsg
    Case WM_MOUSEWHEEL
        Dim liLines As Long, liDelta As Long
        SystemParametersInfo SPI_GETWHEELSCROLLLINES, ZeroL, liLines, ZeroL
        If liLines < 1& Then liLines = 1&
        
        If (wParam And &H8000000) _
            Then liDelta = &H8000& - (wParam And &H7FFF0000) \ &H10000 _
            Else liDelta = -((wParam And &H7FFF0000) \ &H10000)
        
        If KeyIsDown(VK_CONTROL, False) _
            Then liBar = scrHorizontal _
            Else liBar = scrVertical
        
        If Not mtState.bVisible(liBar) Then liBar = (liBar + OneL) And OneL
        
        If liDelta < ZeroL Then
            If pPos(liBar) <= ZeroL Then
                liBar = (liBar + OneL) And OneL
                If pPos(liBar) <= ZeroL Then liDelta = ZeroL
            End If
        ElseIf liDelta > ZeroL Then
            If pPos(liBar) >= mtState.iMax(liBar) - mtState.iLargeChange(liBar) Then
                liBar = (liBar + OneL) And OneL
                If pPos(liBar) >= mtState.iMax(liBar) - mtState.iLargeChange(liBar) Then liDelta = ZeroL
            End If
        End If
        
        If Not mtState.bVisible(liBar) Then liDelta = ZeroL
        
        If liDelta Then
            pPos(liBar) = pPos(liBar) + ((liDelta \ WHEEL_DELTA) * mtState.iSmallChange(liBar) * liLines)
            lReturn = OneL
        End If
        
    Case WM_VSCROLL, WM_HSCROLL
        If uMsg = WM_HSCROLL Then liBar = scrHorizontal Else liBar = scrVertical
        Select Case (wParam And &HFFFF&)
        Case SB_LEFT, SB_TOP:           pPos(liBar) = ZeroL
        Case SB_RIGHT, SB_BOTTOM:       pPos(liBar) = mtState.iMax(liBar)
        Case SB_LINELEFT, SB_LINEUP:    pPos(liBar) = pPos(liBar) - mtState.iSmallChange(liBar)
        Case SB_LINERIGHT, SB_LINEDOWN: pPos(liBar) = pPos(liBar) + mtState.iSmallChange(liBar)
        Case SB_PAGELEFT, SB_PAGEUP:    pPos(liBar) = pPos(liBar) - mtState.iLargeChange(liBar)
        Case SB_PAGERIGHT, SB_PAGEDOWN: pPos(liBar) = pPos(liBar) + mtState.iLargeChange(liBar)
        Case SB_THUMBTRACK:             pPos(liBar) = pTrackPos(liBar)
        End Select
        
    End Select

End Sub

Private Property Get pTrackPos(ByVal liBar As eScrollBar) As Long
    Dim ltSI As SCROLLINFO
    ltSI.cbSize = LenB(ltSI)
    ltSI.fMask = SIF_TRACKPOS
    GetScrollInfo hWnd, liBar, ltSI
    pTrackPos = ltSI.nTrackPos
End Property

Private Sub UserControl_Initialize()
    LoadShellMod
    Subclass_Install Me, hWnd, Array(WM_HSCROLL, WM_VSCROLL, WM_MOUSEWHEEL)
End Sub

Private Sub UserControl_InitProperties()
    mbThemeable = DEF_Themeable
    UserControl.BackColor = DEF_Backcolor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    UserControl.BackColor = PropBag.ReadProperty(PROP_BackColor, DEF_Backcolor)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    pSize
End Sub

Private Sub UserControl_Show()
    Static bInit As Boolean
    If Not bInit Then Me.AutoSize
    bInit = True
End Sub

Private Sub UserControl_Terminate()
    Subclass_Remove Me, hWnd
    ReleaseShellMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    PropBag.WriteProperty PROP_BackColor, UserControl.BackColor, DEF_Backcolor
End Sub

Private Sub pSize(Optional ByVal cx As Long = NegOneL, Optional ByVal cy As Long = NegOneL)
    If Ambient.UserMode Then
        
        If cx = NegOneL Then cx = mtState.iMax(scrHorizontal) Else mtState.iMax(scrHorizontal) = cx
        If cy = NegOneL Then cy = mtState.iMax(scrVertical) Else mtState.iMax(scrVertical) = cy
        
        Dim liViewable(scrHorizontal To scrVertical) As Long
        Dim lbOldVisible(scrHorizontal To scrVertical) As Boolean
        
        liViewable(scrHorizontal) = ScaleX(Width, vbTwips, vbPixels)
        liViewable(scrVertical) = ScaleY(Height, vbTwips, vbPixels)
        
        lbOldVisible(scrHorizontal) = mtState.bVisible(scrHorizontal)
        lbOldVisible(scrVertical) = mtState.bVisible(scrVertical)
        
        mtState.bVisible(scrHorizontal) = False
        mtState.bVisible(scrVertical) = False
        
        If liViewable(scrHorizontal) < cx Then
            mtState.bVisible(scrHorizontal) = True
            liViewable(scrVertical) = liViewable(scrVertical) - GetSystemMetrics(SM_CYHSCROLL)
            If liViewable(scrVertical) < cy Then
                liViewable(scrHorizontal) = liViewable(scrHorizontal) - GetSystemMetrics(SM_CXVSCROLL)
                mtState.bVisible(scrVertical) = True
            End If
        ElseIf liViewable(scrVertical) < cy Then
            mtState.bVisible(scrVertical) = True
            liViewable(scrHorizontal) = liViewable(scrHorizontal) - GetSystemMetrics(SM_CXVSCROLL)
            If liViewable(scrHorizontal) < cx Then
                liViewable(scrVertical) = liViewable(scrVertical) - GetSystemMetrics(SM_CYHSCROLL)
                mtState.bVisible(scrHorizontal) = True
            End If
        End If
        
        If (mtState.bVisible(scrHorizontal) Xor lbOldVisible(scrHorizontal)) Or _
           (mtState.bVisible(scrVertical) Xor lbOldVisible(scrVertical)) Then
            RaiseEvent ScrollBarChange
        End If
        
        pShowScrollBar liViewable(), scrHorizontal, cx
        pShowScrollBar liViewable(), scrVertical, cy
        
    End If
End Sub

Private Sub pShowScrollBar(ByRef iViewable() As Long, ByVal iBar As eScrollBar, ByVal c As Long)
    mtState.iMax(iBar) = c
    
    If mtState.bVisible(iBar) Then
        With pSI(SIF_PAGE Or SIF_RANGE)
            .nMin = ZeroL
            .nMax = mtState.iMax(iBar)
            .nPage = iViewable(iBar)
            SetScrollInfo hWnd, iBar, .cbSize, True
        End With
        mtState.iLargeChange(iBar) = iViewable(iBar)
        mtState.iSmallChange(iBar) = iViewable(iBar) \ 8&
        If iViewable(iBar) - mtState.iTotalOffset(iBar) > c _
            Then pOffsetControls iBar, _
                    iViewable(iBar) - mtState.iTotalOffset(iBar) - c
    Else
        pPos(iBar) = ZeroL
    End If
    
    ShowScrollBar hWnd, iBar, -CLng(mtState.bVisible(iBar))
End Sub

Private Sub pOffsetControls(ByVal iBar As eScrollBar, ByVal iOffset As Long)
    
    If iOffset Then
        
        mtState.iTotalOffset(iBar) = mtState.iTotalOffset(iBar) + iOffset
        
'        'This would be nice, but it only works for controls with a hwnd property (not label, shape, etc.)
'        Dim lhDc As Long
'        lhDc = GetWindowDc(hwnd)
'        If iBar = scrHorizontal Then
'            ScrollWindowEx hwnd, iOffset, ZeroL, ByVal ZeroL, ByVal ZeroL, 0, ByVal ZeroL, SW_INVALIDATE Or SW_SCROLLCHILDREN
'            OffsetWindowOrgEx lhDc, iOffset, ZeroL, ByVal ZeroL
'            OffsetViewportOrgEx lhDc, iOffset, ZeroL, ByVal ZeroL
'        Else
'            ScrollWindowEx hwnd, ZeroL, iOffset, ByVal ZeroL, ByVal ZeroL, 0, ByVal ZeroL, SW_INVALIDATE Or SW_SCROLLCHILDREN
'            OffsetWindowOrgEx lhDc, ZeroL, iOffset, ByVal ZeroL
'            OffsetViewportOrgEx lhDc, ZeroL, iOffset, ByVal ZeroL
'        End If
'
'        ReleaseDC hwnd, lhDc
'        UpdateWindow hwnd
'
'        Exit Sub
        
        Dim lfOffset As Single
        
        If iBar = scrHorizontal _
            Then lfOffset = ScaleX(iOffset, vbPixels, vbTwips) _
            Else lfOffset = ScaleY(iOffset, vbPixels, vbTwips)
        
        Dim loControl As Object
        
        On Error GoTo handler
        
        For Each loControl In ContainedControls
            If iBar = scrHorizontal _
                Then loControl.Left = loControl.Left + lfOffset _
                Else loControl.Top = loControl.Top + lfOffset
            
            If False Then
handler:
                Resume hereandnow
hereandnow:
            End If
        Next
                
        On Error GoTo 0
        
    End If
End Sub

Private Function pSI(ByVal iFlags As Long) As SCROLLINFO
    pSI.cbSize = LenB(pSI)
    pSI.fMask = iFlags
End Function

Private Property Get pPos(ByVal iBar As eScrollBar) As Long
    With pSI(SIF_pPos)
        GetScrollInfo hWnd, iBar And scrVertical, .cbSize
        pPos = .nPos
    End With
End Property
Private Property Let pPos(ByVal iBar As eScrollBar, ByVal iNew As Long)
    Dim liOldPos As Long: liOldPos = pPos(iBar)
    With pSI(SIF_pPos)
        .nPos = iNew
        SetScrollInfo hWnd, iBar And scrVertical, .cbSize, True
        pOffsetControls iBar, liOldPos - pPos(iBar)
    End With
End Property

Public Property Get ColorBack() As OLE_COLOR
    ColorBack = UserControl.BackColor
End Property
Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
    UserControl.BackColor = iNew
    PropertyChanged PROP_BackColor
End Property

Public Sub AutoSize()
    Dim lfMaxRight As Single
    Dim lfMaxBottom As Single
    
    Dim lfTemp As Single
    
    Dim loControl As Object
    
    On Error GoTo handler
    
    For Each loControl In ContainedControls
        lfTemp = loControl.Left + loControl.Width
        If lfTemp > lfMaxRight Then lfMaxRight = lfTemp
        
        lfTemp = loControl.Top + loControl.Height
        If lfTemp > lfMaxBottom Then lfMaxBottom = lfTemp
        
        If False Then
handler:
            Resume hereandnow
hereandnow:
        End If
    Next
    
    On Error GoTo 0
    
    pSize ScaleX(lfMaxRight, vbTwips, vbPixels) - mtState.iTotalOffset(scrHorizontal), ScaleY(lfMaxBottom, vbTwips, vbPixels) - mtState.iTotalOffset(scrVertical)
    
End Sub

Public Property Get Themeable() As Boolean
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
    mbThemeable = bNew
    EnableWindowTheme hWnd, bNew
    PropertyChanged PROP_Themeable
End Property

Public Property Get ViewportLeft() As Single
Attribute ViewportLeft.VB_MemberFlags = "400"
    ViewportLeft = ScaleX(pPos(scrHorizontal), vbPixels, vbContainerPosition)
End Property

Public Property Let ViewportLeft(ByVal fNew As Single)
    pPos(scrHorizontal) = ScaleX(fNew, vbContainerPosition, vbPixels)
End Property

Public Property Get ViewPortTop() As Single
Attribute ViewPortTop.VB_MemberFlags = "400"
    ViewPortTop = ScaleY(pPos(scrVertical), vbPixels, vbContainerPosition)
End Property

Public Property Let ViewPortTop(ByRef fNew As Single)
    pPos(scrVertical) = ScaleY(fNew, vbPixels, vbContainerPosition)
End Property

Public Property Get ViewportWidth() As Single
Attribute ViewportWidth.VB_MemberFlags = "400"
    ViewportWidth = ScaleX(ScaleWidth, vbPixels, vbContainerSize)
End Property

Public Property Get ViewportHeight() As Single
Attribute ViewportHeight.VB_MemberFlags = "400"
    ViewportHeight = ScaleY(ScaleHeight, vbPixels, vbContainerSize)
End Property

Public Property Get ScrollHeight() As Single
Attribute ScrollHeight.VB_MemberFlags = "400"
    ScrollHeight = ScaleY(mtState.iMax(scrVertical), vbPixels, vbContainerSize)
End Property

Public Property Let ScrollHeight(ByVal fNew As Single)
    pSize , ScaleY(fNew, vbContainerSize, vbPixels)
End Property

Public Property Get ScrollWidth() As Single
Attribute ScrollWidth.VB_MemberFlags = "400"
    ScrollWidth = ScaleX(mtState.iMax(scrHorizontal), vbPixels, vbContainerSize)
End Property

Public Property Let ScrollWidth(ByVal fNew As Single)
    pSize ScaleX(fNew, vbContainerSize, vbPixels)
End Property

Public Property Get ScrollBarWidth() As Single
    If mtState.bVisible(scrVertical) Then
        ScrollBarWidth = ScaleX(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbContainerSize)
    End If
End Property

Public Property Get ScrollBarHeight() As Single
    If mtState.bVisible(scrHorizontal) Then
        ScrollBarHeight = ScaleY(GetSystemMetrics(SM_CYHSCROLL), vbPixels, vbContainerSize)
    End If
End Property
