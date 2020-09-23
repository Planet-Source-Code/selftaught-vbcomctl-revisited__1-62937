VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcTrackToolMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcTrackToolMenu.cls                  12/15/04
'
'           PURPOSE:
'               Manage menu tracking for a Toolbar control.
'
'               Two modes:
'                   1. Button Mode
'                       left/right navigate
'                       down moves to popup mode
'                       space activates system menu
'                       letters activate the mnemonics
'                       esc cancels mode
'                   2.Popup Mode
'                       left/right navigate
'                       esc moves to buttonmode
'
'               In both modes, Alt-Space still activates the system menu
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Toolbar/vbAccelerator_ToolBar_and_CoolMenu_Control/VB6_Toolbar_Complete_Source.asp
'               cToolbarMenu.cls
'
'==================================================================================================

Option Explicit

Public Enum eToolMenuTrackState
    tbarTrackNone
    tbarTrackButtons
    tbarTrackDropped
End Enum

Implements iHook

Private miTrackIndex As Long
Private miTrackState As eToolMenuTrackState
Private miNewTrackIndex As Long

Private mbProcessLeft As Boolean
Private mbProcessRight As Boolean
Private mbSwitchMode As Boolean
Private mhMenuFirst As Long

Private mpOwner As Long

Private Sub Class_Terminate()
    pDestroy
End Sub

Private Property Get moOwner() As ucToolbar
    If mpOwner Then
        Dim loTemp As ucToolbar
        ObjectPtr(loTemp) = mpOwner
        Set moOwner = loTemp
        ObjectPtr(loTemp) = ZeroL
    End If
End Property

Private Sub iHook_After(lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Respond to keyboard input to allow arrowkey navigation and accelerators.
'             Respond to mouse input by tracking a menu on the hot button.
'---------------------------------------------------------------------------------------
    If iType <> WH_GETMESSAGE Or wParam = PM_REMOVE Then
        Dim ltMsg As msg
        CopyMemory ltMsg, ByVal lParam, LenB(ltMsg)
        With ltMsg
            Select Case .Message
            Case WM_MENUSELECT:     pMenuSelect .lParam
            Case WM_MOUSEMOVE:      pMouseMove .pt, bHandled
            
            Case WM_KEYDOWN, WM_SYSKEYDOWN
                pKeyPress .wParam And &HFF&, lParam, bHandled
            
            Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, _
                 WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
                pMouseDown .Message, .pt, bHandled
                
            End Select
            
            If bHandled Then
                lReturn = OneL
                If miTrackState = tbarTrackDropped And miNewTrackIndex <> miTrackIndex _
                    Then moOwner.fMenu_CancelDropDown
                
                If iType = WH_GETMESSAGE Then ZeroMemory ByVal lParam, 16&
            End If
            
        End With
    End If

End Sub

Private Sub pMouseDown(ByVal iMsg As Long, ByRef tP As POINT, ByRef bHandled As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Cancel the drop down state if the user clicks the mouse.
'---------------------------------------------------------------------------------------
    If iMsg = WM_LBUTTONDOWN Then
        If miTrackState = tbarTrackDropped Then
            If moOwner.fMenu_HitTest(tP) = miTrackIndex Then
                miNewTrackIndex = NegOneL
                bHandled = True
            End If
            
        Else
            pDestroy
            
        End If
        
    Else
        
        If miTrackState = tbarTrackButtons Then
            Hook_Remove Me, WH_GETMESSAGE
            Hook_Remove Me, WH_MSGFILTER
            miNewTrackIndex = NegOneL
            moOwner.fMenu_Track miTrackState, miNewTrackIndex
        End If
        
    End If
End Sub

Private Sub pKeyPress(ByVal iVirtKey As Long, ByVal lParam As Long, ByRef bHandled As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Perform the action according to the key that was pressed.
'---------------------------------------------------------------------------------------
    bHandled = (miTrackState = tbarTrackButtons)
    Select Case iVirtKey
    Case VK_0 To VK_9, VK_A To VK_Z
        If miTrackState = tbarTrackButtons Then
            Dim lbDupAccel As Boolean
            iVirtKey = moOwner.fMenu_MapAccelerator(iVirtKey, miTrackIndex, lbDupAccel)
            If iVirtKey > NegOneL Then
                bHandled = True
                If miTrackState = tbarTrackDropped Then
                    If iVirtKey <> miTrackIndex Then miNewTrackIndex = iVirtKey
                Else
                    If Not lbDupAccel Then
                        miTrackIndex = iVirtKey
                        miNewTrackIndex = NegOneL
                        mbSwitchMode = True
                    Else
                        If iVirtKey <> miTrackIndex Then miNewTrackIndex = iVirtKey
                    End If
                End If
            End If
        End If
        
    Case VK_MENU, VK_F10
        Const fRepeat As Long = &H40000000
        If (lParam And fRepeat) = ZeroL Then
            If miTrackState = tbarTrackDropped Then
                moOwner.fMenu_CancelDropDown
            Else
                miNewTrackIndex = NegOneL
            End If
            moOwner.fMenu_IgnoreMenuKeyPress = True
            TrackKey.StartTrack Me, iVirtKey
        End If
        bHandled = True
        
    Case VK_DOWN
        If miTrackState = tbarTrackButtons Then
            miNewTrackIndex = NegOneL
            mbSwitchMode = True
            bHandled = True
        End If
        
    Case VK_ESCAPE
        If miTrackState = tbarTrackDropped Then mbSwitchMode = True
        miNewTrackIndex = NegOneL
                            
    Case VK_SPACE
        If miTrackState = tbarTrackButtons Then
            miNewTrackIndex = NegOneL
            SendMessage RootParent(moOwner.hWnd), WM_SYSCOMMAND, &HF100&, &H20&
            bHandled = True
        End If
        
    Case VK_LEFT
        If mbProcessLeft Then
            miNewTrackIndex = moOwner.fMenu_PrevVisible(miTrackIndex)
            bHandled = True
        End If
        
    Case VK_RIGHT
        If mbProcessRight Then
            miNewTrackIndex = moOwner.fMenu_NextVisible(miTrackIndex)
            bHandled = True
        End If
        
    End Select
End Sub

Private Sub pMouseMove(ByRef tP As POINT, ByRef bHandled As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Store the index of the button that the mouse is over.
'---------------------------------------------------------------------------------------
    If Not moOwner.fMenu_ChevronHitTest(tP) Then
        'bHandled = True
    'Else
        Dim liTrackIndex As Long
        liTrackIndex = moOwner.fMenu_HitTest(tP)
        If liTrackIndex <> NegOneL And liTrackIndex <> miTrackIndex Then
            miNewTrackIndex = liTrackIndex
            If miTrackState = tbarTrackDropped Then moOwner.fMenu_CancelDropDown
        End If
    End If
End Sub

Private Sub pMenuSelect(ByVal hMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Store whether left and right arrow keys should be allowed.
'---------------------------------------------------------------------------------------
    mbProcessRight = True
    If mhMenuFirst = ZeroL Then mhMenuFirst = hMenu
    mbProcessLeft = (hMenu = mhMenuFirst)
    
    Dim tMI As MENUITEMINFO
    Dim i As Long
    
    tMI.cbSize = LenB(tMI)
    tMI.fMask = MIIM_STATE Or MIIM_SUBMENU
    For i = ZeroL To GetMenuItemCount(hMenu) - OneL
        GetMenuItemInfo hMenu, i, True, tMI
        If (tMI.fState And MF_HILITE) Then
            mbProcessRight = CBool(tMI.fState And MF_DISABLED) Or Not CBool(tMI.hSubMenu)
            Exit For
        End If
    Next
    
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Remove the subclass and hook and cancel the tracking mode.
'---------------------------------------------------------------------------------------
    
    Hook_Remove Me, WH_GETMESSAGE
    Hook_Remove Me, WH_MSGFILTER
    
    moOwner.fMenu_KillTimer
    
    Dim lbOldSwitchMode As Boolean: lbOldSwitchMode = mbSwitchMode
    Dim liOldTrackState As Long:    liOldTrackState = miTrackState
    Dim liOldTrackIndex As Long:    liOldTrackIndex = miTrackIndex
    
    miTrackState = tbarTrackNone
    miTrackIndex = NegOneL
    miNewTrackIndex = NegOneL
    
    mhMenuFirst = ZeroL
    
    mbProcessLeft = True
    mbProcessRight = True
    mbSwitchMode = False
    
    If liOldTrackState <> tbarTrackNone Then
        moOwner.fMenu_Track miTrackState, miTrackIndex
        If lbOldSwitchMode Then
            If liOldTrackIndex < ZeroL Then liOldTrackIndex = ZeroL
            moOwner.fMenu_PostTrack (liOldTrackState And Not TwoL) + OneL, liOldTrackIndex
        Else
            moOwner.fMenu_ExitTrack
        End If
    End If
    
End Sub

Friend Sub fTrack(ByVal iState As eToolMenuTrackState, ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Track the given toolbar button.
'---------------------------------------------------------------------------------------
    pDestroy
    
    miTrackState = iState
    
    If miTrackState = tbarTrackDropped Then
        
        miNewTrackIndex = iIndex
        
        Hook_Install Me, WH_MSGFILTER
        
        Do
            miTrackIndex = miNewTrackIndex
            miNewTrackIndex = NegOneL
            moOwner.fMenu_Track miTrackState, miTrackIndex
            mhMenuFirst = ZeroL
            mbProcessLeft = True
            mbProcessRight = True
        Loop While miNewTrackIndex > NegOneL
        
        pDestroy
        
    ElseIf miTrackState = tbarTrackButtons Then
        
        miTrackIndex = NegOneL
        miNewTrackIndex = iIndex
        Hook_Install Me, WH_GETMESSAGE
        
        moOwner.fMenu_SetTimer
        
    End If
    
End Sub

Friend Sub fTimer()
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Inform the toolbar of the new hot item and stop tracking if there is no hot item.
'---------------------------------------------------------------------------------------
    Select Case GetForegroundWindow()
    Case RootParent(moOwner.hWnd), moOwner.fMenu_hWndChevron
    Case Else: miNewTrackIndex = NegOneL
    End Select
    
    If miTrackIndex <> miNewTrackIndex And miNewTrackIndex <> NegOneL Then
        miTrackIndex = miNewTrackIndex
        moOwner.fMenu_Track miTrackState, miTrackIndex
    End If
    
    If miNewTrackIndex = NegOneL Then pDestroy
    
End Sub

Friend Property Set fOwner(ByVal oOwner As ucToolbar)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Set the toolbar that this instance will service.
'---------------------------------------------------------------------------------------
    mpOwner = ObjPtr(oOwner)
End Property

Friend Property Get fTrackState() As eToolMenuTrackState
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Get the current tracking state.
'---------------------------------------------------------------------------------------
    fTrackState = miTrackState
End Property

Friend Property Get fTrackIndex() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Get the current tracking index.
'---------------------------------------------------------------------------------------
    fTrackIndex = miTrackIndex
End Property

Friend Sub fTrackKey_KeyUp(ByVal bNonTrackedKeyPressed As Boolean)
    moOwner.fMenu_IgnoreMenuKeyPress = False
End Sub

Friend Sub fHotItemChange(ByVal iNew As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Change the tracking index if the hot item on the toolbar changes and stop
'             the mouseover effect from disappearing if the mouse moves off the toolbar.
'---------------------------------------------------------------------------------------
    If iNew = NegOneL Then
        If miNewTrackIndex <> NegOneL Then lReturn = OneL
    Else
        miNewTrackIndex = iNew
    End If
End Sub
