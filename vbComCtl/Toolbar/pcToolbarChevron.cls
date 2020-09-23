VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcToolbarChevron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcToolbarChevron.cls                  8/25/04
'
'           PURPOSE:
'               Create a dialog window and site a new toolbar onto it to show as a popup
'               window for a rebar chevron.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Toolbar/vbAccelerator_ToolBar_and_CoolMenu_Control/VB6_Toolbar_Complete_Source.asp
'               cChevronWindow.cls
'
'==================================================================================================

Option Explicit

Private Const UM_CANCEL As Long = WM_USER + &H66BD&
Private Const UM_SUSPEND As Long = WM_USER + &H66BE&

Implements iSubclass
Implements iHook

Private mbSuspend As Boolean
Private mbIsShown As Boolean

Private mhWnd As Long
Private mhWndToolbar As Long
Private mhWndParent As Long
Private miOwner As Long

Private mhWndOldParent As Long

Private miLastNCActivate As Long

Private mbThemeable As Boolean

Friend Sub fShow(ByVal oOwner As ucToolbar, ByVal hWndParent As Long, ByVal hWndToolbar As Long, ByVal x As Long, ByVal y As Long, ByVal iWidth As Long, ByVal iHeight As Long, ByRef ptRectExclude As RECT, ByVal iPopupPosition As eToolbarPopupPosition, ByVal bThemeable As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Show the chevron and hold execution until it is hidden.
'---------------------------------------------------------------------------------------
    pDestroy
    
    miOwner = ObjPtr(oOwner)
    
    mhWndParent = hWndParent
    mhWndToolbar = hWndToolbar
    mbThemeable = bThemeable
    
    If pCreate(iWidth + 4&, iHeight + 4&) Then
        mhWndOldParent = GetParent(mhWndToolbar)
        SetParent mhWndToolbar, mhWnd
        SetWindowPos mhWndToolbar, ZeroL, TwoL, TwoL, iWidth, iHeight, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
        pCheckExclude x, iWidth + 4&, y, iHeight + 4&, ptRectExclude, iPopupPosition
        SetWindowPos mhWnd, HWND_TOPMOST, x, y, ZeroL, ZeroL, SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOACTIVATE
        mbIsShown = True
        
        ShowWindow mhWndToolbar, SW_SHOWNORMAL
        
        Hook_Install Me, WH_KEYBOARD
        Hook_Install Me, WH_MOUSE
        
        Do While mbIsShown
            If MsgWaitForMultipleObjects(ZeroL, ZeroL, ZeroL, 25&, QS_ALLINPUT) <> WAIT_TIMEOUT Then DoEvents
        Loop
    End If
    
End Sub

Private Sub pCheckExclude(x As Long, ByVal iWidth As Long, y As Long, ByVal iHeight As Long, ptRectExclude As RECT, ByVal iPopupPosition As eToolbarPopupPosition)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Make sure that the chevron is inside the work area.
'---------------------------------------------------------------------------------------

    Dim ltWorkarea As RECT
    SystemParametersInfo SPI_GETWORKAREA, ZeroL, ltWorkarea, ZeroL
    
    With ltWorkarea
        Select Case iPopupPosition
        Case tbarPopRightDown
            x = ptRectExclude.Right
            y = ptRectExclude.Top
            
            If x + iWidth > .Right Then x = ptRectExclude.Left - iWidth
            
        Case tbarPopRightUp
            x = ptRectExclude.Right
            y = ptRectExclude.bottom - iHeight
            
            If x + iWidth > .Right Then x = ptRectExclude.Left - iWidth
            
        Case tbarPopBottomRight
            x = ptRectExclude.Left
            y = ptRectExclude.bottom
            
            If y + iHeight > .bottom Then y = ptRectExclude.Top - iHeight
        
        Case tbarPopBottomLeft
            x = ptRectExclude.Right - iWidth
            y = ptRectExclude.bottom
            
            If y + iHeight > .bottom Then y = ptRectExclude.Top - iHeight
            
        End Select
        
        If y + iHeight > .bottom Then y = .bottom - iHeight
        If x + iWidth > .Right Then x = .Right - iWidth
        If y < .Top Then y = .Top
        If x < .Left Then x = .Left
        
    End With
End Sub

Private Property Get moOwner() As ucToolbar
    Debug.Assert miOwner
    If miOwner Then
        Dim loToolbar As ucToolbar
        CopyMemory loToolbar, miOwner, 4&
        Set moOwner = loToolbar
        CopyMemory loToolbar, ZeroL, 4&
    End If
End Property

Private Function pCreate(ByVal Width As Long, ByVal Height As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Create the dialog window used as the chevron and the subclasses.
'---------------------------------------------------------------------------------------
    Dim lStyle As Long
   
    lStyle = WS_POPUP Or _
             WS_CLIPSIBLINGS Or _
             WS_CLIPCHILDREN Or _
             WS_OVERLAPPED Or _
             WS_CHILD

    Dim lsAnsi As String
    lsAnsi = StrConv(WC_DIALOG & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(WS_EX_TOOLWINDOW, StrPtr(lsAnsi), ZeroL, lStyle, ZeroL, ZeroL, Width, Height, mhWndParent, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        If mhWndParent Then
            Subclass_Install Me, mhWndParent, Array(WM_CANCELMODE, WM_ACTIVATEAPP, WM_NCACTIVATE)
        End If
        
        SetParent mhWnd, ZeroL
        
        Subclass_Install Me, mhWnd, Array(UM_CANCEL, UM_SUSPEND, WM_DESTROY, WM_ACTIVATE), WM_ERASEBKGND
        
        pCreate = True
    End If
   
End Function

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Destroy the dialog window and the subclasses and hooks.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    
    If mhWnd Then
    
        Hook_Remove Me, WH_KEYBOARD
        Hook_Remove Me, WH_MOUSE
        
        Subclass_Remove Me, mhWndParent
        Subclass_Remove Me, mhWnd
        
        SetParent mhWndToolbar, mhWndOldParent
        DestroyWindow mhWnd
        
        If GetActiveWindow() = mhWndParent Then miLastNCActivate = OneL
        SendMessage mhWndParent, WM_NCACTIVATE, miLastNCActivate, ZeroL
        
        miOwner = ZeroL
    
        mbIsShown = False
        mbSuspend = False
        mhWndOldParent = ZeroL
        mhWnd = ZeroL
        mhWndToolbar = ZeroL
        mhWndParent = ZeroL
        
    End If
    
End Sub

Private Sub pPaint(ByVal hDc As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : paint a border on the chevron window.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    
    If mhWnd Then
        GetWindowRect mhWnd, tR
        OffsetRect tR, -tR.Left, -tR.Top
        DrawEdge hDc, tR, BDR_RAISEDINNER, BF_RECT
    End If
End Sub
Private Sub Class_Terminate()
   pDestroy
End Sub

Private Sub iHook_After(lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long)

End Sub

Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Handle keyboard and mouse hook notifications.
'---------------------------------------------------------------------------------------
    If iType = WH_KEYBOARD Then
        If Not mbSuspend Then
            If nCode = HC_ACTION Then
                If (lParam And &H80000000) = ZeroL Then
                    If ((wParam And &HFFFF&) = VK_ESCAPE) Or ((wParam And &HFFFF&) = VK_MENU) Then
                        pDestroy
                    End If
                End If
            End If
        End If
    Else
        If nCode = HC_ACTION Then
            Dim lhWnd As Long: lhWnd = WindowFromPoint(MemOffset32(lParam, ZeroL), MemOffset32(lParam, 4&))
            If (wParam = WM_LBUTTONDOWN) Or (wParam = WM_LBUTTONDBLCLK) Then
                Select Case lhWnd
                Case moOwner.hWndToolbar
                    If moOwner.MenuStyle Then
                        Me.fSuspend True
                        PostMessage mhWnd, UM_SUSPEND, ZeroL, ZeroL
                        Exit Sub
                    End If
                Case mhWndToolbar
                    If Not moOwner.DroppedMenuButton Is Nothing Then
                        If (wParam = WM_LBUTTONDOWN) Then moOwner.fMenu_CancelDropDown
                        bHandled = True
                        lReturn = OneL
                    End If
                    Exit Sub
                End Select
            End If
            If wParam = WM_LBUTTONDOWN Or wParam = WM_RBUTTONDOWN Or wParam = WM_MBUTTONDOWN _
            Or wParam = WM_NCLBUTTONDOWN Or wParam = WM_NCRBUTTONDOWN Or wParam = WM_NCMBUTTONDOWN Then
                If lhWnd <> mhWndToolbar Then
                    Dim lsClass As String
                    lsClass = Space$(MAX_PATH \ 2)
                    
                    If GetClassName(lhWnd, StrPtr(lsClass), LenB(lsClass)) Then
                        If StrConv(MidB$(lsClass, 1, 7), vbUnicode) = "#32768" Then
                            bHandled = True
                            lReturn = OneL
                            pDestroy
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
    Case WM_ERASEBKGND
        If wParam <> ZeroL Then
            If CBool(IsAppThemed) And mbThemeable Then pDrawGradient wParam
            pPaint wParam
            lReturn = OneL
        End If
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Destroy the chevron on WM_CANCELMODE, WM_ACTIVATEAPP, UM_CANCEL, WM_DESTROY
'             Paint background.
'             prevent parent's titlebar from going out of focus by intercepting WM_NCACTIVATE
'---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_NCACTIVATE
        miLastNCActivate = wParam
        lReturn = 1
        bHandled = True
    Case UM_SUSPEND
        fSuspend False
    Case WM_CANCELMODE
        If Not mbSuspend Then pDestroy
    Case WM_ACTIVATEAPP
        If wParam = ZeroL Then pDestroy
    Case WM_ACTIVATE
        If Not mbSuspend Then PostMessage mhWnd, UM_CANCEL, ZeroL, ZeroL
    Case UM_CANCEL, WM_DESTROY
        pDestroy
    End Select
End Sub

Private Sub pDrawGradient(ByVal hDc As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Draw a gradient that looks like a rebar band in comctl 6.
'---------------------------------------------------------------------------------------
    Dim ltR As RECT
    GetWindowRect mhWnd, ltR
    DrawGradient hDc, ZeroL, ZeroL, ltR.Right - ltR.Left, ltR.bottom - ltR.Top, RGB(244, 243, 238), RGB(224, 226, 235)
    
    'Dim ltRect as RECT
    'ltRect.Right = tR.Width
    'ltRect.Bottom = tR.Height
    'DrawRebarThemeBackground moWnd.hwnd, hdc, ltRect
End Sub

Friend Sub fSuspend(ByVal bVal As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Suspend the chevron input state, allowing a popup menu to be shown
'             without interfering.
'---------------------------------------------------------------------------------------
    If mbSuspend And Not (bVal Or mhWnd = ZeroL) Then
        If GetActiveWindow <> mhWnd Then
            'moWnd.SetPos ZeroL, ZeroL, 0, 0, HWND_TOPMOST, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOACTIVATE
            ShowWindow mhWnd, SW_HIDE
            ShowWindow mhWnd, SW_SHOW
        End If
    End If
    mbSuspend = bVal
End Sub

Friend Sub fHide()
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Destroy the chevron.
'---------------------------------------------------------------------------------------
    pDestroy
End Sub

Friend Property Get fVisible() As Boolean
    fVisible = mbIsShown
End Property

Friend Property Get fhWnd() As Long
    fhWnd = mhWnd
End Property
