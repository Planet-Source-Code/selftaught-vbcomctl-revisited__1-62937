VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcTrackKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcTrackKeyboard.cls      4/24/05
'
'           PURPOSE:
'               Install a keyboard hook and wait for a key to be released.  This is used to monitor
'               the Alt and F10 keys to enter menu tracking mode when either one is pressed and
'               released without any other keys pressed while it was down.
'
'==================================================================================================

Option Explicit

Implements iHook

Private miTrackKey As Long
Private mbNonTrackedKeyPressed As Boolean
Private moOwner As Object

Public Sub StartTrack(ByVal oOwner As Object, ByVal iVirtKey As Long)
    'Debug.Assert moOwner Is Nothing Or (oOwner Is moOwner And iVirtKey = miTrackKey)
    If moOwner Is Nothing Then
        Set moOwner = oOwner
        miTrackKey = iVirtKey
        mbNonTrackedKeyPressed = False
        Hook_Install Me, WH_KEYBOARD_LL
    End If
End Sub

Public Sub StopTrack(ByVal oOwner As Object)
    If moOwner Is oOwner Then
        Set moOwner = Nothing
        Hook_Remove Me, WH_KEYBOARD_LL
    End If
End Sub

Private Sub Class_Terminate()
    Hook_Remove Me, WH_KEYBOARD_LL
End Sub

Private Function pKeyMatch(ByVal iKey As Long) As Boolean
    Select Case iKey And &HFF&
    Case VK_LMENU, VK_RMENU, VK_MENU
        pKeyMatch = miTrackKey = VK_MENU
    Case VK_F10
        pKeyMatch = miTrackKey = VK_F10
    End Select
End Function

Private Sub iHook_After(lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, ByVal iType As eHookType, ByVal nCode As eHookCode, ByVal wParam As Long, ByVal lParam As Long)
    Const KBDLLHOOKSTRUCT_VKCode As Long = 0
    If nCode = HC_ACTION Then
        If Not pKeyMatch(MemOffset32(lParam, KBDLLHOOKSTRUCT_VKCode)) Then
            mbNonTrackedKeyPressed = True
        ElseIf wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            Debug.Assert Not moOwner Is Nothing
            If Not mbNonTrackedKeyPressed Then
                bHandled = True
                lReturn = ZeroL
            End If
            Hook_Remove Me, WH_KEYBOARD_LL
            If TypeOf moOwner Is ucToolbar Then
                Dim loToolbar As ucToolbar
                Set loToolbar = moOwner
                loToolbar.fTrackKey_KeyUp mbNonTrackedKeyPressed
            ElseIf TypeOf moOwner Is pcTrackToolMenu Then
                Dim loTrackMenu As pcTrackToolMenu
                Set loTrackMenu = moOwner
                loTrackMenu.fTrackKey_KeyUp mbNonTrackedKeyPressed
            End If
            Set moOwner = Nothing
        ElseIf wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
            bHandled = True
            lReturn = OneL
        End If
    End If
End Sub
