VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcPopupSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cPopupSubclass.cls                    7/15/05
'           PURPOSE:
'               Manage a subclass of a top-level window for popup menu messages.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================
Option Explicit

Implements iSubclass

Private moOwner As ucPopupMenus
Private mpMenu As Long
Private mhWndSubclass As Long

Friend Sub fSubclass(ByVal oOwner As ucPopupMenus, ByVal pMenu As Long, ByVal hWndSubclass As Long)
    fUnSubclass
    
    mhWndSubclass = hWndSubclass
    
    If mhWndSubclass Then
        Set moOwner = oOwner
        mpMenu = pMenu
        Subclass_Install Me, mhWndSubclass, Array(WM_MENUSELECT, WM_MEASUREITEM, WM_DRAWITEM, _
                                                  WM_MENUCHAR, WM_INITMENUPOPUP, WM_UNINITMENUPOPUP, WM_TIMER)
    End If
End Sub

Friend Sub fUnSubclass()
    If mhWndSubclass Then
        Subclass_Remove Me, mhWndSubclass
        mhWndSubclass = ZeroL
    End If
End Sub

Private Sub Class_Terminate()
    fUnSubclass
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    moOwner.fSubclass_Proc mpMenu, bHandled, lReturn, hWnd, uMsg, wParam, lParam
End Sub
