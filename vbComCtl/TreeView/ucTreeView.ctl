VERSION 5.00
Begin VB.UserControl ucTreeView 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   PropertyPages   =   "ucTreeView.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucTreeView.ctx":000D
End
Attribute VB_Name = "ucTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucTreeView.ctl        12/15/04
'
'           PURPOSE:
'               Implement the comctl32 treeview control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/TreeView/TreeView_Control/VB6_TreeView_Full_Source.asp
'               vbalTreeView.ctl
'
'==================================================================================================
Option Explicit

Public Enum eTreeViewNodeRelation
    tvwChild = &H0
    tvwLast = &H1
    tvwFirst = &H2
    tvwSorted = &H4
    tvwSibling = &H8
End Enum

Public Enum eTreeViewGetNode
    tvwGetNodeFirstChild = TVGN_CHILD
    tvwGetNodeNextVisible = TVGN_NEXTVISIBLE
    tvwGetNodePreviousVisible = TVGN_PREVIOUSVISIBLE
    tvwGetNodeNextSibling = TVGN_NEXT
    tvwGetNodePreviousSibling = TVGN_PREVIOUS
    tvwGetNodeRoot = TVGN_ROOT
    tvwGetNodeParent = TVGN_PARENT
End Enum

Public Enum eTreeViewHitTest
   tvwHitTestAbove = TVHT_ABOVE
   tvwHitTestBelow = TVHT_BELOW
   tvwHitTestBelowLast = TVHT_NOWHERE
   tvwHitTestItemPlusMinus = TVHT_ONITEMBUTTON
   tvwHitTestItemIcon = TVHT_ONITEMICON
   tvwHitTestItemIndent = TVHT_ONITEMINDENT
   tvwHitTestItemText = TVHT_ONITEMLABEL
   tvwHitTestItemRight = TVHT_ONITEMRIGHT
   tvwHitTestItemState = TVHT_ONITEMSTATEICON
   tvwHitTestLeft = TVHT_TOLEFT
   tvwHitTestRight = TVHT_TORIGHT
End Enum

Public Event AfterLabelEdit(ByVal oNode As cNode, ByRef sNew As String, ByRef bCancel As OLE_CANCELBOOL)
Public Event BeforeCollapse(ByVal oNode As cNode, ByRef bCancel As OLE_CANCELBOOL)
Public Event BeforeExpand(ByVal oNode As cNode, ByRef bCancel As OLE_CANCELBOOL)
Public Event BeforeLabelEdit(ByVal oNode As cNode, ByRef bCancel As OLE_CANCELBOOL)
Public Event Collapse(ByVal oNode As cNode)
Public Event Expand(ByVal oNode As cNode)
Public Event KeyDown(ByVal iKeyCode As Integer, ByVal iState As evbComCtlKeyboardState)
Public Event Click()
Public Event RightClick()
Public Event NodeDrag(ByVal oNode As cNode)
Public Event NodeRightDrag(ByVal oNode As cNode)
Public Event NodeClick(ByVal oNode As cNode, ByVal iHitTestCode As eTreeViewHitTest)
Public Event NodeDblClick(ByVal oNode As cNode, ByVal iHitTestCode As eTreeViewHitTest)
Public Event NodeRightClick(ByVal oNode As cNode, ByVal iHitTestCode As eTreeViewHitTest)
Public Event NodeCheck(ByVal oNode As cNode)
Public Event NodeSelect(ByVal oNode As cNode)
Public Event OLECompleteDrag(Effect As evbComCtlOleDropEffect)
Public Event OLEDragDrop(Data As DataObject, Effect As evbComCtlOleDropEffect, Button As evbComCtlMouseButton, Shift As evbComCtlKeyboardState, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As evbComCtlOleDropEffect, Button As evbComCtlMouseButton, Shift As evbComCtlKeyboardState, x As Single, y As Single, State As evbComCtlOleDragOverState)
Public Event OLEGiveFeedback(Effect As evbComCtlOleDropEffect, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As evbComCtlOleDropEffect)

Implements iSubclass
Implements iOleInPlaceActiveObjectVB
Implements iOleControlVB

' See KB Q261289
Private Const UM_CHECKSTATECHANGED = WM_USER + &H112
'
Private Const UM_STARTDRAG = WM_USER + &H113

Private Enum eStyle
    bpHasButtons = TVS_HASBUTTONS
    bpHasLines = TVS_HASLINES
    bpCheckBoxes = TVS_CHECKBOXES
    bpFullRowSelect = TVS_FULLROWSELECT
    bpNoScroll = TVS_NOSCROLL
    bpTrackSelect = TVS_TRACKSELECT
    bpLinesAtRoot = TVS_LINESATROOT
    bpShowSelAlways = TVS_SHOWSELALWAYS
    bpSingleExpand = TVS_SINGLEEXPAND
    bpEditLabels = TVS_EDITLABELS
End Enum

Private Const PROP_Font = "Font"
Private Const PROP_Style = "Style"
Private Const PROP_BorderStyle = "Border"
Private Const PROP_ShowNumbers = "ShowNums"
Private Const PROP_ItemHeight = "ItemHeight"
Private Const PROP_Indent = "Indent"
Private Const PROP_BackColor = "Backcolor"
Private Const PROP_ForeColor = "Forecolor"
Private Const PROP_LineColor = "Linecolor"
Private Const PROP_PathSeparator = "Path"
Private Const PROP_Themeable = "Themeable"
Private Const PROP_OleDrop = "OleDrop"

Private Const DEF_Style = bpHasButtons Or bpHasLines
Private Const DEF_BorderStyle = vbccBorderThin
Private Const DEF_ShowNumbers = True
Private Const DEF_ItemHeight = ZeroL
Private Const DEF_Indent = ZeroL
Private Const DEF_Backcolor = vbWindowBackground
Private Const DEF_ForeColor = vbWindowText
Private Const DEF_LineColor = vbWindowText
Private Const DEF_PathSeparator = "\"
Private Const DEF_Themeable = True
Private Const DEF_OleDrop = False

Private WithEvents moFont As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private moImageList     As cImageList
Private WithEvents moImageListEvent As cImageList
Attribute moImageListEvent.VB_VarHelpID = -1

Private mhWnd           As Long
Private moKeyMap        As pcStringMap
Private moItemDataMap   As pcIntegerMap

Private miStyle         As eStyle
Private miBorderStyle   As evbComCtlBorderStyle

Private mbShowNumbers   As Boolean
Private mbThemeable     As Boolean

Private miItemHeight    As Long
Private miIndent        As Long

Private miBackColor     As OLE_COLOR
Private miForeColor     As OLE_COLOR
Private miLineColor     As OLE_COLOR

Private msPathSeparator As String

Private mtInsert        As TVINSERTSTRUCT

Private lR              As Long

Private msTextBuffer    As String * 130
Private mhImlCheckBoxes As Long
Private mhFont          As Long

Private Const ucTreeView = "ucTreeView"
Private Const cNode = "cNode"

Const NMHDR_hwndFrom    As Long = 0
Const NMHDR_code        As Long = 8
Const NMTVKEYDOWN_wVKey As Long = 12
Const TVDISPINFO_TVITEM_hItem  As Long = 16
'Const NMTREEVIEW_itemOld_hItem As Long = 20
Const TVDISPINFO_TVITEM_pszText As Long = 28
Const NMTREEVIEW_itemOld_hItem  As Long = 12
Const NMTREEVIEW_itemOld_lParam As Long = 52
Const NMTREEVIEW_action         As Long = 12
Const NMTREEVIEW_itemNew_hItem  As Long = 60
Const NMCUSTOMDRAW_dwDrawStage  As Long = 12
Const NMCUSTOMDRAW_hdc          As Long = 16
'Const NMCUSTOMDRAW_rc           As Long = 20
Const NMCUSTOMDRAW_dwItemSpec   As Long = 36
'Const NMCUSTOMDRAW_uItemState   As Long = 40
Const NMCUSTOMDRAW_lItemlParam  As Long = 44

Private Const NODE_lpKey        As Long = ZeroL
Private Const NODE_lpKeyNext    As Long = 4&
Private Const NODE_iItemData    As Long = 8&
Private Const NODE_iItemDataNext As Long = 12&
Private Const NODE_iItemNumber  As Long = 16&
Private Const NODE_hItem        As Long = 20&
Private Const NODE_Len          As Long = 24&

Private mhItemInsertMark        As Long
Private mbInsertMarkAfter       As Boolean

Private miLastXMouseDown        As Long
Private miLastYMouseDown        As Long
Private mbInLabelEdit           As Boolean
Private mbIgnoreNotification    As Boolean
Private mbRedraw                As Boolean

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
' Purpose   : Handle notifications from the treeview.
'---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_NOTIFY
        If MemOffset32(lParam, NMHDR_hwndFrom) = mhWnd Then
            lReturn = ZeroL
            bHandled = True
            Select Case MemOffset32(lParam, NMHDR_code)
            Case TVN_DELETEITEM
                If mhItemInsertMark Then
                    If MemOffset32(lParam, NMTREEVIEW_itemOld_hItem) = mhItemInsertMark _
                        Then mhItemInsertMark = ZeroL
                End If
                If Not mbIgnoreNotification Then pcNode_Terminate MemOffset32(lParam, NMTREEVIEW_itemOld_lParam)
            Case NM_CUSTOMDRAW
                pCustomDraw lParam, lReturn
            Case NM_CLICK, NM_RCLICK, NM_DBLCLK
                pOnClick MemOffset32(lParam, NMHDR_code)
            Case TVN_KEYDOWN
                pOnKeyDown MemOffset16(lParam, NMTVKEYDOWN_wVKey)
            Case TVN_BEGINDRAG, TVN_BEGINRDRAG
                PostMessage mhWnd, UM_STARTDRAG, MemOffset32(lParam, NMHDR_code), MemOffset32(lParam, NMTREEVIEW_itemNew_hItem)
            Case NM_SETFOCUS
                ActivateIPAO Me
            Case TVN_BEGINLABELEDIT, TVN_ENDLABELEDIT
                pOnLabelEdit MemOffset32(lParam, NMHDR_code), lParam, lReturn
            Case TVN_ITEMEXPANDED, TVN_ITEMEXPANDING
                If Not mbIgnoreNotification Then pOnItemExpand MemOffset32(lParam, NMHDR_code), lParam, lReturn
            Case TVN_SELCHANGED
                RaiseEvent NodeSelect(pItem(MemOffset32(lParam, NMTREEVIEW_itemNew_hItem)))
                
            End Select
        End If
    Case WM_PARENTNOTIFY
        bHandled = True
        If Not CBool(miStyle And bpSingleExpand) Then
            If (wParam And &HFFFF&) = WM_RBUTTONDOWN Or (wParam And &HFFFF&) = WM_LBUTTONDOWN Then
                miLastXMouseDown = loword(lParam)
                miLastYMouseDown = hiword(lParam)
                SendMessage mhWnd, TVM_SELECTITEM, TVGN_CARET, pHitTest(miLastXMouseDown, miLastYMouseDown)
            End If
        End If
        
    Case UM_CHECKSTATECHANGED
        SendMessage mhWnd, TVM_SELECTITEM, TVGN_CARET, lParam
        RaiseEvent NodeCheck(pItem(lParam))
        
    Case UM_STARTDRAG
        pOnDrag wParam, lParam
        
    Case WM_SETFOCUS
        ActivateIPAO Me
        
    Case WM_MOUSEACTIVATE
        If GetFocus() <> mhWnd Then
            If Not mbInLabelEdit Then
                bHandled = True
                vbComCtlTlb.SetFocus UserControl.hWnd
                lReturn = MA_NOACTIVATE
            End If
        End If
        
    End Select

End Sub

Private Sub moImageListEvent_Changed()
    Set Me.ImageList = moImageList
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    LoadShellMod
    InitCC ICC_TREEVIEW_CLASSES
    Set moFontPage = New pcSupportFontPropPage
    mbRedraw = True
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    
    Set moFont = Font_CreateDefault(Ambient.Font)
    miStyle = DEF_Style
    miBorderStyle = DEF_BorderStyle
    mbShowNumbers = DEF_ShowNumbers
    miItemHeight = DEF_ItemHeight
    miIndent = DEF_Indent
    miBackColor = DEF_Backcolor
    miForeColor = DEF_ForeColor
    miLineColor = DEF_LineColor
    msPathSeparator = DEF_PathSeparator
    mbThemeable = DEF_Themeable
    UserControl.OLEDropMode = -DEF_OleDrop
    pCreate
    On Error GoTo 0
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    ImageDrag_Stop
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ImageDrag_Stop
    RaiseEvent OLEDragDrop(Data, Effect, CLng(Button), CLng(Shift), ScaleX(x, ScaleMode, vbContainerPosition), ScaleY(y, ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    ImageDrag_Move mhWnd, x, y, State
    
    Static shItemExpand As Long
    Static siTickCount As Long
    
    Dim lhItemExpand As Long
    
    If State <> vbccOleDragLeave Then
        lhItemExpand = pHitTest(x, y, TVHT_ONITEMBUTTON)
        If lhItemExpand Then
            If CBool(pItem_State(lhItemExpand, TVIS_EXPANDED)) Then lhItemExpand = ZeroL
            If lhItemExpand Then
                If lhItemExpand <> shItemExpand Then
                    siTickCount = ZeroL
                Else
                    If siTickCount = ZeroL Then
                        siTickCount = GetTickCount()
                    Else
                        If TickDiff(GetTickCount(), siTickCount) > 1000& Then
                            ImageDrag_Show False
                            If mhItemInsertMark Then SendMessage mhWnd, TVM_SETINSERTMARK, ZeroL, ZeroL
                            SendMessage mhWnd, TVM_EXPAND, TVE_EXPAND, lhItemExpand
                            If mhItemInsertMark Then SendMessage mhWnd, TVM_SETINSERTMARK, CLng(-mbInsertMarkAfter), mhItemInsertMark
                            lhItemExpand = ZeroL
                            UpdateWindow mhWnd
                            ImageDrag_Show True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    shItemExpand = lhItemExpand
    
    RaiseEvent OLEDragOver(Data, Effect, CLng(Button), CLng(Shift), ScaleX(x, ScaleMode, vbContainerPosition), ScaleY(y, ScaleMode, vbContainerPosition), CLng(State))
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
    'ImageDrag_Show CBool(Effect And Not vbDropEffectScroll)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    miStyle = PropBag.ReadProperty(PROP_Style, DEF_Style)
    miBorderStyle = PropBag.ReadProperty(PROP_BorderStyle, DEF_BorderStyle)
    mbShowNumbers = PropBag.ReadProperty(PROP_ShowNumbers, DEF_ShowNumbers)
    miItemHeight = PropBag.ReadProperty(PROP_ItemHeight, DEF_ItemHeight)
    miIndent = PropBag.ReadProperty(PROP_Indent, DEF_Indent)
    miBackColor = PropBag.ReadProperty(PROP_BackColor, DEF_Backcolor)
    miForeColor = PropBag.ReadProperty(PROP_ForeColor, DEF_ForeColor)
    miLineColor = PropBag.ReadProperty(PROP_LineColor, DEF_LineColor)
    msPathSeparator = PropBag.ReadProperty(PROP_PathSeparator, DEF_PathSeparator)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    UserControl.OLEDropMode = -PropBag.ReadProperty(PROP_OleDrop, DEF_OleDrop)
    pCreate
    On Error GoTo 0
End Sub

Private Sub UserControl_Resize()
    If mhWnd Then MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
End Sub

Private Sub UserControl_Show()
    'see KB241102
    If mhWnd Then
        SetWindowStyle mhWnd, TVS_NOTOOLTIPS, ZeroL
        SetWindowStyle mhWnd, ZeroL, TVS_NOTOOLTIPS
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Style, miStyle, DEF_Style
    PropBag.WriteProperty PROP_BorderStyle, miBorderStyle, DEF_BorderStyle
    PropBag.WriteProperty PROP_ShowNumbers, mbShowNumbers, DEF_ShowNumbers
    PropBag.WriteProperty PROP_ItemHeight, miItemHeight, DEF_ItemHeight
    PropBag.WriteProperty PROP_Indent, miIndent, DEF_Indent
    PropBag.WriteProperty PROP_BackColor, miBackColor, DEF_Backcolor
    PropBag.WriteProperty PROP_ForeColor, miForeColor, DEF_ForeColor
    PropBag.WriteProperty PROP_LineColor, miLineColor, DEF_LineColor
    PropBag.WriteProperty PROP_PathSeparator, msPathSeparator, DEF_PathSeparator
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    PropBag.WriteProperty PROP_OleDrop, CBool(-UserControl.OLEDropMode), DEF_OleDrop
    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
    pDestroy
    ReleaseShellMod
    If mhFont Then moFont.ReleaseHandle mhFont
    Set moFontPage = Nothing
End Sub

Private Sub iOleControlVB_OnMnemonic(bHandled As Boolean, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iOleControlVB_GetControlInfo(bHandled As Boolean, iAccelCount As Long, hAccelTable As Long, iFlags As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Notify the container that we capture the return key to prevent it from displaying
'             default buttons.
'---------------------------------------------------------------------------------------
    bHandled = True
    iFlags = vbccEatsReturn
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Capture the keys we want to forward to the treeview.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Then
            Select Case wParam And &HFFFF&
            Case vbKeyPageUp To vbKeyDown, vbKeyReturn
                Dim lhWnd As Long
                lhWnd = SendMessage(mhWnd, TVM_GETEDITCONTROL, ZeroL, ZeroL)
                If lhWnd = ZeroL Then
                    SendMessage mhWnd, uMsg, wParam, lParam
                Else
                    SendMessage lhWnd, uMsg, wParam, lParam
                End If
                bHandled = True
            End Select
        End If
    End If
End Sub

Private Function pcNode_SetKey(ByVal lpNode As Long, ByRef sIn As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Allocate a string with the new key and store its pointer, maintain
'             the keys collection.
'---------------------------------------------------------------------------------------
    Debug.Assert lpNode
    
    If lpNode Then
        
        Dim lpKeyOld As Long:   lpKeyOld = MemOffset32(lpNode, NODE_lpKey)
        
        If LenB(sIn) Then
            
            If moKeyMap Is Nothing Then Set moKeyMap = New pcStringMap
            
            Dim lsAnsi As String:   lsAnsi = StrConv(sIn & vbNullChar, vbFromUnicode)
            Dim lpKeyNew As Long:   lpKeyNew = StrPtr(lsAnsi)
            Dim liHashNew As Long:  liHashNew = Hash(lpKeyNew, lstrlen(lpKeyNew))
            
            If moKeyMap.Find(lpKeyNew, liHashNew) = ZeroL Then
                
                lpKeyNew = MemAllocFromString(lpKeyNew, LenB(lsAnsi))
                
                If CBool(lpKeyNew) Then
                    pcNode_SetKey = OneL
                    MemOffset32(lpNode, NODE_lpKey) = lpKeyNew
                    moKeyMap.Add lpNode, liHashNew
                Else
                    pcNode_SetKey = ZeroL
                End If
            Else
                pcNode_SetKey = NegOneL
            End If
            
        Else
            
            pcNode_SetKey = OneL
            MemOffset32(lpNode, NODE_lpKey) = ZeroL
            
        End If
        
        If (pcNode_SetKey = OneL) And CBool(lpKeyOld) Then
            moKeyMap.Remove lpNode, Hash(lpKeyOld, lstrlen(lpKeyOld))
            MemFree lpKeyOld
        End If
        
    End If
    
End Function

Private Function pcNode_SetItemData(ByVal lpItem As Long, ByVal iItemData As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Store the itemdata as a unique identifier.
'---------------------------------------------------------------------------------------
    Debug.Assert lpItem
    
    If lpItem Then
        
        Dim liItemDataOld As Long:  liItemDataOld = MemOffset32(lpItem, NODE_iItemData)
        Dim liHash As Long
        
        If iItemData <> ZeroL Then
            If moItemDataMap Is Nothing Then Set moItemDataMap = New pcIntegerMap
            
            liHash = HashLong(iItemData)
            
            If moItemDataMap.Find(iItemData, liHash) = ZeroL _
                Then pcNode_SetItemData = OneL _
                Else pcNode_SetItemData = NegOneL
            
        Else
            
            pcNode_SetItemData = OneL
            
        End If
        
        If (pcNode_SetItemData = OneL) Then
            If liItemDataOld Then moItemDataMap.Remove lpItem, HashLong(liItemDataOld)
            MemOffset32(lpItem, NODE_iItemData) = iItemData
            If iItemData Then moItemDataMap.Add lpItem, liHash
        End If
        
    End If
    
End Function


Private Function pcNode_Initialize(ByRef sKey As String, iItemData As Long, ByVal iItemNumber As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Allocate memory for extra node data an initialize the values.
'---------------------------------------------------------------------------------------
    pcNode_Initialize = MemAlloc(NODE_Len)
    
    If pcNode_Initialize Then
        MemOffset32(pcNode_Initialize, NODE_lpKey) = ZeroL
        MemOffset32(pcNode_Initialize, NODE_iItemData) = ZeroL
        MemOffset32(pcNode_Initialize, NODE_iItemNumber) = iItemNumber
        
        Dim liResult As Long
        liResult = pcNode_SetKey(pcNode_Initialize, sKey)
        If liResult = OneL Then liResult = pcNode_SetItemData(pcNode_Initialize, iItemData)
        
        If liResult <> OneL Then
            MemFree pcNode_Initialize
            If liResult = ZeroL Then gErr vbccOutOfMemory, ucTreeView
            If liResult = NegOneL Then gErr vbccKeyAlreadyExists, ucTreeView
            Debug.Assert False
        End If
    End If
End Function

Private Sub pcNode_Terminate(ByVal lpNode As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Release the memory allocated for extra node data.
'---------------------------------------------------------------------------------------
    Debug.Assert lpNode
    If lpNode Then
        Dim lpKey As Long
        
        lpKey = pcNode_lpKey(lpNode)
        If lpKey Then
            moKeyMap.Remove lpNode, Hash(lpKey, lstrlen(lpKey))
            MemFree lpKey
        End If
        
        lpKey = pcNode_ItemData(lpNode)
        If lpKey Then moItemDataMap.Remove lpNode, HashLong(lpKey)
        
        MemFree lpNode
        
    End If
End Sub

Private Property Get pcNode_lpKey(ByVal lpNode As Long) As Long
    Debug.Assert lpNode
    If lpNode Then
        pcNode_lpKey = MemOffset32(lpNode, NODE_lpKey)
    End If
End Property
Private Property Let pcNode_lpKey(ByVal lpNode As Long, ByVal iNew As Long)
    Debug.Assert lpNode
    If lpNode Then
        MemOffset32(lpNode, NODE_lpKey) = iNew
    End If
End Property

Private Property Get pcNode_ItemData(ByVal lpNode As Long) As Long
    Debug.Assert lpNode
    If lpNode Then
        pcNode_ItemData = MemOffset32(lpNode, NODE_iItemData)
    End If
End Property
'Private Property Let pcNode_ItemData(ByVal lpNode As Long, ByVal iNew As Long)
'    Debug.Assert lpNode
'    If lpNode Then
'        MemOffset32(lpNode, NODE_iItemData) = iNew
'    End If
'End Property

Private Property Get pcNode_ItemNumber(ByVal lpNode As Long) As Long
    Debug.Assert lpNode
    If lpNode Then
        pcNode_ItemNumber = MemOffset32(lpNode, NODE_iItemNumber)
    End If
End Property
Private Property Let pcNode_ItemNumber(ByVal lpNode As Long, ByVal iNew As Long)
    Debug.Assert lpNode
    If lpNode Then
        MemOffset32(lpNode, NODE_iItemNumber) = iNew
    End If
End Property

Private Property Get pcNode_hItem(ByVal lpNode As Long) As Long
    Debug.Assert lpNode
    If lpNode Then
        pcNode_hItem = MemOffset32(lpNode, NODE_hItem)
    End If
End Property
Private Property Let pcNode_hItem(ByVal lpNode As Long, ByVal iNew As Long)
    Debug.Assert lpNode
    If lpNode Then
        MemOffset32(lpNode, NODE_hItem) = iNew
    End If
End Property

Private Function pHitTest(ByVal x As Long, ByVal y As Long, Optional ByVal iMask As Long = TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON Or TVHT_ONITEMBUTTON) As Long
'---------------------------------------------------------------------------------------
' Date      : 3/5/05
' Purpose   : Return the hItem if the x,y is on an item label or icon.
'---------------------------------------------------------------------------------------
    Dim tHT As TVHITTESTINFO
    With tHT.pt
        .x = x
        .y = y
    End With
    SendMessage mhWnd, TVM_HITTEST, ZeroL, VarPtr(tHT)
    If tHT.Flags And iMask Then pHitTest = tHT.hItem
End Function

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Update the font handle used by the treeview.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    Dim hFont As Long
    hFont = moFont.GetHandle
    SendMessage mhWnd, WM_SETFONT, hFont, OneL
    If mhFont Then moFont.ReleaseHandle mhFont
    mhFont = hFont
    On Error GoTo 0
    
    Exit Sub
handler:
    Resume Next
End Sub

Private Sub moFont_Changed()
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    pPropChanged PROP_Font
End Sub



Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Create the treeview and the subclasses.
'---------------------------------------------------------------------------------------

    pDestroy
    
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_TREEVIEW & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, (miStyle Or WS_CHILD Or WS_VISIBLE) And Not bpCheckBoxes, ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
    
        EnableWindowTheme mhWnd, mbThemeable
        If CBool(miStyle And bpCheckBoxes) Then
            miStyle = miStyle And Not bpCheckBoxes
            CheckBoxes = True
        End If
        UserControl_Resize
        
        Dim liMajor As Long
        GetCCVersion liMajor
        SendMessage mhWnd, CCM_SETVERSION, liMajor, ZeroL
        SendMessage mhWnd, TVM_SETBKCOLOR, ZeroL, TranslateColor(miBackColor)
        SendMessage mhWnd, TVM_SETTEXTCOLOR, ZeroL, TranslateColor(miForeColor)
        SendMessage mhWnd, TVM_SETLINECOLOR, ZeroL, TranslateColor(miLineColor)
        SendMessage mhWnd, TVM_SETINDENT, miIndent, ZeroL
        
        pSetFont
        
        pSetBorder
        
        If UserControl.Ambient.UserMode Then
            VTableSubclass_OleControl_Install Me
            VTableSubclass_IPAO_Install Me
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_PARENTNOTIFY), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_SETFOCUS, WM_MOUSEACTIVATE, UM_CHECKSTATECHANGED, UM_STARTDRAG), WM_KILLFOCUS
            
        End If
    End If

End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Destroy the treeview and the subclasses.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        VTableSubclass_OleControl_Remove
        VTableSubclass_IPAO_Remove
        
        SendMessage mhWnd, TVM_DELETEITEM, ZeroL, TVI_ROOT
        
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        
        DestroyWindow mhWnd
        mhWnd = ZeroL
        
        If mhImlCheckBoxes Then ImageList_Destroy mhImlCheckBoxes
        mhImlCheckBoxes = ZeroL
    End If
    
End Sub

Private Sub pSetBorder()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the border of the control and redraw it.
'---------------------------------------------------------------------------------------
If mhWnd Then
    Select Case miBorderStyle
    Case vbccBorderSunken
        If Not Ambient.UserMode Then UserControl.BackColor = Me.ColorBack
        UserControl.BorderStyle = vbFixedSingle
        UserControl.Appearance = OneL
        SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    Case vbccBorderSingle
        UserControl.BorderStyle = vbFixedSingle
        UserControl.Appearance = ZeroL
        SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    Case vbccBorderNone
        UserControl.BorderStyle = vbBSNone
        UserControl.Appearance = ZeroL
        SetWindowStyleEx mhWnd, ZeroL, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    Case vbccBorderThin
        UserControl.BorderStyle = vbBSNone
        UserControl.Appearance = ZeroL
        SetWindowStyleEx mhWnd, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE
    End Select
    
    SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER

End If
End Sub

Private Sub pSetStyle(ByVal iStyle As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the treeview window style.
'---------------------------------------------------------------------------------------
    If bNew Then
        miStyle = miStyle Or iStyle
    Else
        miStyle = miStyle And Not iStyle
    End If
    If mhWnd Then
        SetWindowStyle mhWnd, iStyle * -bNew, iStyle * -(Not bNew)
        SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
    End If
    pPropChanged PROP_Style
End Sub

Private Sub pPropChanged(ByRef s As String)
    If Ambient.UserMode = False Then PropertyChanged s
End Sub

Private Function pItem(ByVal hItem As Long) As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cNode item for the given hNode.
'---------------------------------------------------------------------------------------
    If hItem Then
        Set pItem = New cNode
        pItem.fInit Me, hItem
    End If
End Function

Private Sub pOnDrag(ByVal iDrag As Long, ByVal hItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Raise a drag event.
'---------------------------------------------------------------------------------------
    If hItem Then
        If iDrag = TVN_BEGINRDRAG Then
            RaiseEvent NodeRightDrag(pItem(hItem))
        Else
            RaiseEvent NodeDrag(pItem(hItem))
        End If
    End If
End Sub

Private Sub pOnKeyDown(ByVal iKey As Integer)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Raise a keydown event.
'---------------------------------------------------------------------------------------
    Dim hItem As Long
    If CBool(miStyle And bpCheckBoxes) Then
        If iKey = vbKeySpace Then
            hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CARET, TVI_ROOT)
            If hItem Then PostMessage mhWnd, UM_CHECKSTATECHANGED, ZeroL, hItem
        End If
    End If
    RaiseEvent KeyDown(iKey, KBState())
End Sub

Private Sub pOnItemExpand(ByVal iMsg As Long, ByVal lParam As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Raise expand/collapse events.
'---------------------------------------------------------------------------------------
    Dim loNode As cNode
    Dim bCancel As OLE_CANCELBOOL
    
    If iMsg = TVN_ITEMEXPANDED Then
        Set loNode = pItem(MemOffset32(lParam, NMTREEVIEW_itemNew_hItem))
        Debug.Assert Not loNode Is Nothing
        Select Case MemOffset32(lParam, NMTREEVIEW_action)
        Case TVE_EXPAND, TVE_EXPANDPARTIAL
            RaiseEvent Expand(loNode)
        Case TVE_COLLAPSE, TVE_COLLAPSERESET
            RaiseEvent Collapse(loNode)
        Case Else
            Debug.Assert False
        End Select
                
    ElseIf iMsg = TVN_ITEMEXPANDING Then
        Set loNode = pItem(MemOffset32(lParam, NMTREEVIEW_itemNew_hItem))
        Debug.Assert Not loNode Is Nothing
        Select Case MemOffset32(lParam, NMTREEVIEW_action)
        Case TVE_EXPAND, TVE_EXPANDPARTIAL
            RaiseEvent BeforeExpand(loNode, bCancel)
        Case TVE_COLLAPSE, TVE_COLLAPSERESET
            RaiseEvent BeforeCollapse(loNode, bCancel)
        Case Else
            Debug.Assert False
        End Select
        
        lReturn = CLng(-bCancel)
        
    End If
End Sub

Private Sub pOnLabelEdit(ByVal iMsg As Long, ByVal lParam As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Raise labeledit events.
'---------------------------------------------------------------------------------------
    Dim lpKey As Long
    Dim loNode As cNode
    Dim bCancel As OLE_CANCELBOOL
    Dim ls As String
    
    If iMsg = TVN_BEGINLABELEDIT Then
        mbInLabelEdit = True
        Set loNode = pItem(MemOffset32(lParam, TVDISPINFO_TVITEM_hItem))
        Debug.Assert Not loNode Is Nothing
        If Not loNode Is Nothing Then
            RaiseEvent BeforeLabelEdit(loNode, bCancel)
        End If
        lReturn = -bCancel
    ElseIf iMsg = TVN_ENDLABELEDIT Then
        mbInLabelEdit = False
        lpKey = MemOffset32(lParam, TVDISPINFO_TVITEM_pszText)
        
        If lpKey Then
            
            lstrToStringA lpKey, ls
            Set loNode = pItem(MemOffset32(lParam, TVDISPINFO_TVITEM_hItem))
            Debug.Assert Not loNode Is Nothing
            
            If Not loNode Is Nothing Then
                RaiseEvent AfterLabelEdit(loNode, ls, bCancel)
                ls = StrConv(ls & vbNullChar, vbFromUnicode)
                MidB$(msTextBuffer, 1, LenB(ls)) = ls
                MidB$(msTextBuffer, LenB(msTextBuffer), 1) = vbNullChar
                MemOffset32(lParam, TVDISPINFO_TVITEM_pszText) = StrPtr(msTextBuffer)
                
            End If
            lReturn = -(Not bCancel)
        End If
    End If
End Sub


Private Sub pOnClick(ByVal iNotification As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Raise click events.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim tHT As TVHITTESTINFO
        Dim loNode As cNode
        
        GetCursorPos tHT.pt
        ScreenToClient mhWnd, tHT.pt
        SendMessage mhWnd, TVM_HITTEST, ZeroL, VarPtr(tHT)
        If tHT.Flags And (TVHT_ONITEMINDENT Or TVHT_ONITEMRIGHT) Then tHT.hItem = 0
        
        Set loNode = pItem(tHT.hItem)
        
        If iNotification = NM_CLICK Then
            If loNode Is Nothing Then
                RaiseEvent Click
            Else
                RaiseEvent NodeClick(loNode, tHT.Flags)
            End If
        ElseIf iNotification = NM_DBLCLK Then
            If loNode Is Nothing Then
                'RaiseEvent DblClick
            Else
                RaiseEvent NodeDblClick(loNode, tHT.Flags)
            End If
        ElseIf iNotification = NM_RCLICK Then
            If loNode Is Nothing Then
                RaiseEvent RightClick
            Else
                RaiseEvent NodeRightClick(loNode, tHT.Flags)
            End If
        End If
        
        If tHT.hItem Then
            If CBool(miStyle And bpCheckBoxes) Then
                If CBool(tHT.Flags And TVHT_ONITEMSTATEICON) Then
                    PostMessage mhWnd, UM_CHECKSTATECHANGED, ZeroL, tHT.hItem
                End If
            End If
        End If
        
    End If
End Sub

Private Sub pCustomDraw(ByVal lParam As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Handle the custom draw notification.
'---------------------------------------------------------------------------------------
    Dim hDc As Long
    Dim liNum As Long
    Dim liOrigColor As Long
    Dim tR As RECT
    Dim tSize As SIZE
    Dim ls As String
    Dim lpItem As Long
    
    If Not mbShowNumbers Then
        lReturn = CDRF_DODEFAULT
        
    Else

        Select Case MemOffset32(lParam, NMCUSTOMDRAW_dwDrawStage)
            Case CDDS_ITEMPOSTPAINT
                lReturn = CDRF_SKIPDEFAULT
                
                tR.Left = MemOffset32(lParam, NMCUSTOMDRAW_dwItemSpec)
                
                lpItem = MemOffset32(lParam, NMCUSTOMDRAW_lItemlParam)
                
                If lpItem Then
                
                    liNum = pcNode_ItemNumber(lpItem)
                    
                    If liNum >= ZeroL Then
                        hDc = MemOffset32(lParam, NMCUSTOMDRAW_hdc)
                        liOrigColor = SetTextColor(hDc, vbBlue)
                        
                        ls = StrConv("(" & CStr(liNum) & ")" & vbNullChar, vbFromUnicode)
                         
                        If GetTextExtentPoint32(hDc, ByVal StrPtr(ls), LenB(ls) - OneL, tSize) = ZeroL Then tSize.cx = 50
                        
                        If SendMessage(mhWnd, TVM_GETITEMRECT, OneL, VarPtr(tR)) Then
                            tR.Left = tR.Right + 10&
                            tR.Right = tR.Left + tSize.cx
                            
                            DrawText hDc, ByVal StrPtr(ls), NegOneL, tR, DT_SINGLELINE Or DT_VCENTER
                            
                        End If
                        
                        SetTextColor hDc, liOrigColor
                        
                    End If
                End If
            Case CDDS_ITEMPREPAINT
                lReturn = CDRF_NOTIFYPOSTPAINT
                
            Case Else
                lReturn = CDRF_NOTIFYITEMDRAW
                
        End Select
    End If

End Sub


Friend Sub fItems_Enum_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the next cNode in an enumeration.
'---------------------------------------------------------------------------------------
    tEnum.iControl = pItems_Enum_NextItem(tEnum.iControl)
    bNoMoreItems = Not CBool(tEnum.iControl)
    If Not bNoMoreItems Then
        Dim loItem As cNode
        Set loItem = New cNode
        loItem.fInit Me, tEnum.iControl
        Set vNextItem = loItem
    End If
End Sub

Friend Sub fItems_Enum_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Skip a number of items in an enumeration.
'---------------------------------------------------------------------------------------
    bSkippedAll = True
    Do Until iSkipCount < OneL
        tEnum.iControl = pItems_Enum_NextItem(tEnum.iControl)
        If tEnum.iControl = ZeroL Then
            bSkippedAll = False
            Exit Sub
        End If
        iSkipCount = iSkipCount - OneL
    Loop
End Sub

Private Function pItems_Enum_NextItem(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the next item for enumeration.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        pItems_Enum_NextItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
        If pItems_Enum_NextItem = ZeroL And hItem <> TVI_ROOT Then
            pItems_Enum_NextItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
            If pItems_Enum_NextItem = ZeroL Then
                
                Dim hItemTest As Long
                hItemTest = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_PARENT, hItem)
                
                Do While hItemTest
                    pItems_Enum_NextItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItemTest)
                    If pItems_Enum_NextItem <> ZeroL Then Exit Do
                    hItemTest = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_PARENT, hItemTest)
                Loop
                
            End If
        End If
    End If
End Function


Friend Function fItems_Add( _
            ByRef vNodeRelative As Variant, _
            ByVal iRelation As eTreeViewNodeRelation, _
            ByRef sKey As String, _
            ByRef sText As String, _
            ByVal iIconIndex As Long, _
            ByVal iIconIndexSelected As Long, _
            ByVal iIconIndexState As Long, _
            ByVal iItemData As Long, _
            ByVal iItemNumber As Long, _
            ByVal bShowPlusMinus As Boolean) _
                As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add an node to the treeview.
'---------------------------------------------------------------------------------------
If mhWnd Then

    Set fItems_Add = pItem_Add(pItem_hRelative(vNodeRelative, iRelation), iRelation, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
    
End If
End Function

Friend Sub fItems_Remove(ByRef vNode As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove a node from the treeview.
'---------------------------------------------------------------------------------------
    Dim hItem As Long
    hItem = pItem_hItem(vNode)
    If hItem = ZeroL Then gErr vbccKeyOrIndexNotFound, ucTreeView
    If mhWnd Then SendMessage mhWnd, TVM_DELETEITEM, ZeroL, hItem
End Sub

Friend Function fItems_Item(ByRef vNode As Variant) As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an object representing a given node.
'---------------------------------------------------------------------------------------
    Set fItems_Item = pItem(pItem_hItem(vNode))
    If fItems_Item Is Nothing Then gErr vbccKeyOrIndexNotFound, ucTreeView
End Function

Friend Sub fItems_Clear()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove all nodes from the treeview.
'---------------------------------------------------------------------------------------
    If mhWnd Then SendMessage mhWnd, TVM_DELETEITEM, ZeroL, TVI_ROOT
End Sub

Friend Property Get fItems_Exists(ByRef vNode As Variant) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether a node exists in the treeview.
'---------------------------------------------------------------------------------------
    fItems_Exists = CBool(pItem_hItem(vNode))
End Property

Friend Property Get fItems_Count() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the count of nodes in the treeview.
'---------------------------------------------------------------------------------------
   fItems_Count = SendMessage(mhWnd, TVM_GETCOUNT, ZeroL, ZeroL)
End Property

Friend Property Get fItems_NewEnum(ByVal oItems As cNodes) As IUnknown
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a new enumeration for the given collection.
'---------------------------------------------------------------------------------------
    Dim loEnum As New pcEnumeration
    Set fItems_NewEnum = loEnum.GetEnum(oItems, TVI_ROOT)
End Property

Friend Sub fItem_Enum_GetNextItem(ByVal hItem As Long, tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the next node in an enumeration.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        
        If tEnum.iData = ZeroL Then
            tEnum.iData = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
        Else
            tEnum.iData = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, tEnum.iData)
        End If
        
        bNoMoreItems = Not CBool(tEnum.iData)
        
        If Not bNoMoreItems Then
            Dim loItem As cNode
            Set loItem = New cNode
            loItem.fInit Me, tEnum.iData
            Set vNextItem = loItem
        End If
        
    Else
        bNoMoreItems = True
        
    End If
    
End Sub

Friend Sub fItem_Enum_Skip(ByVal hItem As Long, tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Skip a number of nodes in an enumeration.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        bSkippedAll = True
        If iSkipCount > ZeroL Then
        
            If tEnum.iData = ZeroL Then
                tEnum.iData = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
            Else
                tEnum.iData = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, tEnum.iData)
            End If
            
            iSkipCount = iSkipCount - OneL
            
            Do Until iSkipCount < OneL
                tEnum.iData = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, tEnum.iData)
                If tEnum.iData = ZeroL Then
                    bSkippedAll = False
                    Exit Do
                End If
                iSkipCount = iSkipCount - OneL
            Loop
            
        End If
        
    Else
        
        bSkippedAll = False
        
    End If
End Sub


Friend Function fItem_AddChildNode( _
        ByVal hItem As Long, _
        ByRef vNodeAfter As Variant, _
        ByRef sKey As String, _
        ByRef sText As String, _
        ByVal iIconIndex As Long, _
        ByVal iIconIndexSelected As Long, _
        ByVal iIconIndexState As Long, _
        ByVal iItemData As Long, _
        ByVal iItemNumber As Long, _
        ByVal bShowPlusMinus As Boolean) _
                As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add a child node to an existing node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        Dim hInsertAfter As Long
        If Not IsMissing(vNodeAfter) Then
            hInsertAfter = pItem_hItem(vNodeAfter)
            If hInsertAfter = ZeroL Then gErr vbccKeyOrIndexNotFound, cNode
        Else
            hInsertAfter = TVI_LAST
        End If
        
        Set fItem_AddChildNode = pItem_Add(hItem, hInsertAfter, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
    End If
End Function

Friend Function fItem_AddChildNodeFirst( _
        ByVal hItem As Long, _
        ByRef sKey As String, _
        ByRef sText As String, _
        ByVal iIconIndex As Long, _
        ByVal iIconIndexSelected As Long, _
        ByVal iIconIndexState As Long, _
        ByVal iItemData As Long, _
        ByVal iItemNumber As Long, _
        ByVal bShowPlusMinus As Boolean) _
                As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Insert a child node immediately below the specified node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        Set fItem_AddChildNodeFirst = pItem_Add(hItem, TVI_FIRST, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
    End If
End Function


Friend Function fItem_AddChildNodeSorted( _
        ByVal hItem As Long, _
        ByRef sKey As String, _
        ByRef sText As String, _
        ByVal iIconIndex As Long, _
        ByVal iIconIndexSelected As Long, _
        ByVal iIconIndexState As Long, _
        ByVal iItemData As Long, _
        ByVal iItemNumber As Long, _
        ByVal bShowPlusMinus As Boolean) _
                As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add a child node in sorted order.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        Set fItem_AddChildNodeSorted = pItem_Add(hItem, TVI_SORT, sKey, sText, iIconIndex, iIconIndexSelected, iIconIndexState, iItemData, iItemNumber, bShowPlusMinus)
    End If
End Function


Friend Property Get fItem_GetNode(ByVal hItem As Long, ByVal iNode As eTreeViewGetNode) As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the node that is requested.
'---------------------------------------------------------------------------------------
    'Debug.Assert iNode <> tvwGetNodeParent
    If pItem_Verify(hItem, True) Then
        If iNode = tvwGetNodeRoot Then
            Set fItem_GetNode = New cNode
            fItem_GetNode.fInit Me, TVI_ROOT
        ElseIf hItem <> TVI_ROOT Or (iNode = tvwGetNodeFirstChild Or iNode = tvwGetNodeNextVisible Or iNode = tvwGetNodeRoot Or iNode = tvwGetNodeRoot) Then
            hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, iNode, hItem)
            If hItem Then
                Set fItem_GetNode = New cNode
                fItem_GetNode.fInit Me, hItem
            Else
                If hItem = ZeroL And iNode = tvwGetNodeParent Then
                    Set fItem_GetNode = New cNode
                    fItem_GetNode.fInit Me, TVI_ROOT
                End If
            End If
        End If
    End If
End Property

Friend Sub fItem_Delete(ByRef hItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Delete an item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        SendMessage mhWnd, TVM_DELETEITEM, ZeroL, hItem
        If hItem <> TVI_ROOT Then hItem = ZeroL
    End If
End Sub

Friend Sub fItem_DeleteChildren(ByVal hItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Delete all child nodes of an item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
        
        Dim lhItemNext As Long
        Do While CBool(hItem)
            lhItemNext = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
            SendMessage mhWnd, TVM_DELETEITEM, ZeroL, hItem
            hItem = lhItemNext
        Loop
    End If
End Sub

Friend Property Get fItem_ChildCount(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of child nodes from a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
        Do While CBool(hItem)
            fItem_ChildCount = fItem_ChildCount + OneL
            hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
        Loop
    End If
End Property

Friend Property Get fItem_hItem(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hitem.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        fItem_hItem = hItem
    End If
End Property

Friend Property Get fItem_Bold(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the bold state of the given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_Bold = CBool(pItem_State(hItem, TVIS_BOLD))
    End If
End Property
Friend Property Let fItem_Bold(ByVal hItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the bold state of the given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_State(hItem, TVIS_BOLD) = (CLng(-bNew) * TVIS_BOLD)
    End If
End Property

Friend Property Get fItem_IconIndexState(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the state icon index.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_IconIndexState = ((pItem_State(hItem, TVIS_STATEIMAGEMASK) And TVIS_STATEIMAGEMASK) \ &H1000&)
    End If
End Property
Friend Property Let fItem_IconIndexState(ByVal hItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the state icon index.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_State(hItem, TVIS_STATEIMAGEMASK) = ((iNew And &HF&) * &H1000&)
    End If
End Property

Friend Property Get fItem_IconIndex(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the icon index.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_IconIndex = pItem_Info(hItem, TVIF_IMAGE)
    End If
End Property
Friend Property Let fItem_IconIndex(ByVal hItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the icon index.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_Info(hItem, TVIF_IMAGE) = iNew
    End If
End Property

Friend Property Get fItem_IconIndexSelected(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the selected icon index.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_IconIndexSelected = pItem_Info(hItem, TVIF_SELECTEDIMAGE)
    End If
End Property
Friend Property Let fItem_IconIndexSelected(ByVal hItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the selected icon index.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_Info(hItem, TVIF_SELECTEDIMAGE) = iNew
    End If
End Property

Friend Property Get fItem_DropHighlighted(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the node is displayed as a drop highlight.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_DropHighlighted = CBool(pItem_State(hItem, TVIS_DROPHILITED))
    End If
End Property
Friend Property Let fItem_DropHighlighted(ByVal hItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether a node is displayed as a drop highlight.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_State(hItem, TVIS_DROPHILITED) = (-CLng(bNew) * TVIS_DROPHILITED)
    End If
End Property

Friend Sub fItem_EnsureVisible(ByVal hItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Ensure that a node is visible in the list.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        If mhWnd Then SendMessage mhWnd, TVM_ENSUREVISIBLE, ZeroL, hItem
    End If
End Sub

Friend Property Get fItem_Expanded(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether a node is expanded.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_Expanded = CBool(pItem_State(hItem, TVIS_EXPANDED))
    End If
End Property
Friend Property Let fItem_Expanded(ByVal hItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether a node is expanded.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        If mhWnd Then
            If bNew Then
                SendMessage mhWnd, TVM_EXPAND, TVE_EXPAND, hItem
            Else
                SendMessage mhWnd, TVM_EXPAND, TVE_COLLAPSE, hItem
            End If
        End If
    End If
End Property

Friend Property Get fItem_FullPath(ByVal hItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the full path of item text to the root node.
'---------------------------------------------------------------------------------------
    If mhWnd <> ZeroL And hItem <> TVI_ROOT Then
        If pItem_Verify(hItem) Then
            Dim ls As String
            Do While CBool(hItem)
                fItem_FullPath = msPathSeparator & pItem_Text(hItem) & fItem_FullPath
                hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_PARENT, hItem)
            Loop
        End If
    End If
End Property

Friend Property Get fItem_ItemData(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the item data of the given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_ItemData = pcNode_ItemData(pItem_Info(hItem, TVIF_PARAM))
    End If
End Property
Friend Property Let fItem_ItemData(ByVal hItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the item data of the given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        If pcNode_SetItemData(pItem_Info(hItem, TVIF_PARAM), iNew) = NegOneL Then gErr vbccKeyAlreadyExists, cNode
    End If
End Property

Friend Property Get fItem_ItemNumber(ByVal hItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the item number that is displayed for a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_ItemNumber = pcNode_ItemNumber(pItem_Info(hItem, TVIF_PARAM))
    End If
End Property
Friend Property Let fItem_ItemNumber(ByVal hItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the item number that is displayed for a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        Dim liOld As Long
        Dim lpItem As Long
        
        lpItem = pItem_Info(hItem, TVIF_PARAM)
        liOld = pcNode_ItemNumber(lpItem)
        pcNode_ItemNumber(lpItem) = iNew
        
        Dim tR As RECT
        If mhWnd Then
            Dim tSize As SIZE
            Dim liWidthOld As Long
            Dim lDc As Long
            Dim ls As String
            
            lDc = GetDC(mhWnd)
            
            If lDc Then
                            
                ls = "(" & liOld & ")"
                GetTextExtentPoint32W lDc, ls, Len(ls), tSize
                liWidthOld = tSize.cx
                
                ls = "(" & iNew & ")"
                GetTextExtentPoint32W lDc, ls, Len(ls), tSize
                If liWidthOld > tSize.cx Then tSize.cx = liWidthOld
                     
                ReleaseDC mhWnd, lDc
            End If
           
            tR.Left = hItem
            SendMessage mhWnd, TVM_GETITEMRECT, ZeroL, VarPtr(tR)
            'must invalidate the item also to cause the item customdraw notification to be sent
            tR.Right = tR.Right + 10& + tSize.cx
            InvalidateRect mhWnd, tR, OneL
        End If
    End If
End Property

Friend Property Get fItem_Key(ByVal hItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the key for a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        Dim lpKey As Long
        lpKey = pcNode_lpKey(pItem_Info(hItem, TVIF_PARAM))
        If lpKey Then
            lstrToStringA lpKey, fItem_Key
        End If
    End If
End Property
Friend Property Let fItem_Key(ByVal hItem As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the key for a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        
        Dim liResult As Long
        liResult = pcNode_SetKey(pItem_Info(hItem, TVIF_PARAM), sNew)
        
        If liResult <> OneL Then
            If liResult = ZeroL Then gErr vbccOutOfMemory, ucTreeView
            If liResult = NegOneL Then gErr vbccKeyAlreadyExists, ucTreeView
            Debug.Assert False
        End If
        
    End If
End Property

Friend Property Get fItem_Selected(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the selected state for a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_Selected = pItem_State(hItem, TVIS_SELECTED)
    End If
End Property
Friend Property Let fItem_Selected(ByVal hItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the selected state for a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_State(hItem, TVIS_SELECTED) = (CLng(-bNew) * TVIS_SELECTED)
    End If
End Property

Friend Property Get fItem_ShowPlusMinus(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get whether the node displays the expand/collapse button.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_ShowPlusMinus = CBool(pItem_Info(hItem, TVIF_CHILDREN))
    End If
End Property
Friend Property Let fItem_ShowPlusMinus(ByVal hItem As Long, ByRef bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the node displays the expand/collapse button.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_Info(hItem, TVIF_CHILDREN) = CLng(-bNew)
    End If
End Property

Friend Property Get fItem_Text(ByVal hItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the text of a node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_Text = pItem_Text(hItem)
    End If
End Property
Friend Property Let fItem_Text(ByVal hItem As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of a node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_Text(hItem) = sNew
    End If
End Property

Friend Property Get fItem_Cut(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the icon displays lighter, in a ghosted state.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        fItem_Cut = CBool(pItem_State(hItem, TVIS_CUT))
    End If
End Property
Friend Property Let fItem_Cut(ByVal hItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the icon displays lighter, in a ghosted state.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        pItem_State(hItem, TVIS_CUT) = TVIS_CUT * CLng(-bNew)
    End If
End Property

Friend Sub fItem_Sort(ByVal hItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Sort the children of a given node.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        SendMessage mhWnd, TVM_SORTCHILDREN, ZeroL, hItem
    End If
End Sub

Friend Property Get fItem_Left(ByVal hItem As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the left edge of the node in container coordinates.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        Dim tR As RECT
        tR.Left = hItem
        SendMessage mhWnd, TVM_GETITEMRECT, OneL, VarPtr(tR)
        fItem_Left = ScaleX(tR.Left, vbPixels, vbContainerPosition)
    End If
End Property
Friend Property Get fItem_Top(ByVal hItem As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the top edge of the node in container coordinates.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        Dim tR As RECT
        tR.Left = hItem
        SendMessage mhWnd, TVM_GETITEMRECT, OneL, VarPtr(tR)
        fItem_Top = ScaleX(tR.Top, vbPixels, vbContainerPosition)
    End If
End Property
Friend Property Get fItem_Width(ByVal hItem As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the width of the node in container coordinates.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        Dim tR As RECT
        tR.Left = hItem
        SendMessage mhWnd, TVM_GETITEMRECT, OneL, VarPtr(tR)
        fItem_Width = ScaleX(tR.Right - tR.Left, vbPixels, vbContainerSize)
    End If
End Property
Friend Property Get fItem_Height(ByVal hItem As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the height of the node in container coordinates.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        Dim tR As RECT
        tR.Left = hItem
        SendMessage mhWnd, TVM_GETITEMRECT, OneL, VarPtr(tR)
        fItem_Height = ScaleX(tR.bottom - tR.Top, vbPixels, vbContainerSize)
    End If
End Property

Friend Property Get fItem_HasChildren(ByVal hItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the item has any child nodes.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem, True) Then
        fItem_HasChildren = CBool(SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem))
    End If
End Property

Friend Sub fItem_Move(ByRef hItem As Long, ByRef vNodeRelative As Variant, ByVal iRelation As eTreeViewNodeRelation)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove the selected node and its children and insert identical nodes at
'             the new location.
'---------------------------------------------------------------------------------------
    If pItem_Verify(hItem) Then
        
        Dim lhItemRelative As Long
        lhItemRelative = pItem_hRelative(vNodeRelative, iRelation)
        
        If lhItemRelative <> TVI_ROOT Then
            Dim lhItemParent As Long
            lhItemParent = lhItemRelative
            Do While lhItemParent
                If lhItemParent = hItem Then gErr vbccInvalidProcedureCall, ucTreeView, "Cannot move a node to be a decendent of itself."
                lhItemParent = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_PARENT, lhItemParent)
            Loop
        End If
        
        mbIgnoreNotification = True
        
        Dim lhItemInserted As Long
        lhItemInserted = pItem_Move(hItem, lhItemRelative, iRelation)
        
        SendMessage mhWnd, TVM_SELECTITEM, TVGN_CARET, lhItemInserted
        SendMessage mhWnd, TVM_DELETEITEM, ZeroL, hItem
        hItem = lhItemInserted
        
        mbIgnoreNotification = False
        
    End If
End Sub

Private Function pItem_Move(ByVal hItem As Long, ByVal hItemNewParent As Long, ByVal hInsertAfter As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Recursively move the node and its children.
'---------------------------------------------------------------------------------------
    
    
    Dim lbExpanded As Boolean
    
    If hItem Then
        
        With mtInsert
            .itemex.hItem = hItem
            .itemex.Mask = TVIF_ALL
            
            If SendMessage(mhWnd, TVM_GETITEM, ZeroL, VarPtr(.itemex)) Then
                
                lbExpanded = .itemex.State And TVIS_EXPANDED
                
                mtInsert.hParent = hItemNewParent
                mtInsert.hInsertAfter = hInsertAfter
                
                pItem_Move = SendMessage(mhWnd, TVM_INSERTITEM, ZeroL, VarPtr(mtInsert))
                
                If pItem_Move Then
                    
                    Dim lhChild As Long
                    lhChild = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
                    
                    Do While CBool(lhChild)
                        pItem_Move lhChild, pItem_Move, TVI_LAST
                        lhChild = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_NEXT, lhChild)
                    Loop
                    
                    If lbExpanded Then SendMessage mhWnd, TVM_EXPAND, TVE_EXPAND, pItem_Move
                    
                End If
                
            End If
        End With
    End If
    
    If pItem_Move = ZeroL Then
        Debug.Assert False
        pcNode_Terminate mtInsert.itemex.lParam
    End If
    
End Function

Private Function pItem_hRelative(ByRef vNodeRelative As Variant, ByRef iRelation As eTreeViewNodeRelation) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a hItem and a TVI_* value from a node and relation value.
'---------------------------------------------------------------------------------------
    
    If IsMissing(vNodeRelative) Then
        pItem_hRelative = TVI_ROOT
    Else
        pItem_hRelative = pItem_hItem(vNodeRelative)
        If pItem_hRelative = ZeroL Then gErr vbccKeyOrIndexNotFound, ucTreeView
    End If
    
    If CBool(iRelation And tvwSibling) Then
        If pItem_hRelative <> TVI_ROOT Then
            iRelation = pItem_hRelative
            pItem_hRelative = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_PARENT, pItem_hRelative)
            If pItem_hRelative = ZeroL Then pItem_hRelative = TVI_ROOT
        End If
    End If
    
    If CBool(iRelation And tvwSorted) Then
        iRelation = TVI_SORT
    ElseIf CBool(iRelation And tvwLast) Then
        iRelation = TVI_LAST
    ElseIf CBool(iRelation And tvwFirst) Then
        iRelation = TVI_FIRST
    End If
End Function


Private Function pItem_Verify(ByVal hItem As Long, Optional ByVal bRootOK As Boolean) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Verify that a node still exists in the collection.
'---------------------------------------------------------------------------------------
    If hItem = TVI_ROOT Then
        If Not bRootOK Then gErr vbccInvalidProcedureCall, cNode, "Procedure call is not valid for the root node."
        pItem_Verify = True
    Else
        pItem_Verify = pItem_Info(hItem, TVIF_PARAM)
        pItem_Verify = pItem_VerifyResult()
    End If
End Function

Private Function pItem_VerifyResult() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : If the return of the last api is NULL, then the node is no longer part of the collection.
'---------------------------------------------------------------------------------------
    pItem_VerifyResult = CBool(lR)
    If Not pItem_VerifyResult Then gErr vbccItemDetached, cNode
End Function


Private Property Get pItem_Info(ByVal hItem As Long, ByVal iMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a member of the TVITEMEX structure.
'---------------------------------------------------------------------------------------
    lR = ZeroL
    With mtInsert.itemex
        .hItem = hItem
        .Mask = TVIF_HANDLE Or iMask
        If Not (CBool(iMask And TVIF_PARAM) And hItem = TVI_ROOT) Then
            If mhWnd Then lR = SendMessage(mhWnd, TVM_GETITEM, ZeroL, VarPtr(mtInsert.itemex)) Else lR = ZeroL
        Else
            lR = OneL
        End If
        
        If iMask = TVIF_IMAGE Then
            pItem_Info = .iImage
        ElseIf iMask = TVIF_SELECTEDIMAGE Then
            pItem_Info = .iSelectedImage
        ElseIf iMask = TVIF_PARAM Then
            pItem_Info = .lParam
        ElseIf iMask = TVIF_CHILDREN Then
            pItem_Info = .cChildren
        End If
    
    End With
    
End Property

Private Property Let pItem_Info(ByVal hItem As Long, ByVal iMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a member of the TVITEMEX structure.
'---------------------------------------------------------------------------------------
    lR = ZeroL
    
    With mtInsert.itemex
    
        .hItem = hItem
    
        .Mask = TVIF_HANDLE Or iMask
    
        If iMask = TVIF_IMAGE Then
            .iImage = iNew
        ElseIf iMask = TVIF_SELECTEDIMAGE Then
            .iSelectedImage = iNew
        ElseIf iMask = TVIF_PARAM Then
            .lParam = iNew
        ElseIf iMask = TVIF_CHILDREN Then
            .cChildren = iNew
        End If
        
        If mhWnd Then lR = SendMessage(mhWnd, TVM_SETITEM, ZeroL, VarPtr(mtInsert.itemex))
        
    End With
End Property

Private Property Get pItem_State(ByVal hItem As Long, ByVal iStateMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the state of a given item.
'---------------------------------------------------------------------------------------
    lR = ZeroL
    With mtInsert.itemex
        .hItem = hItem
    
        .Mask = TVIF_HANDLE Or TVIF_STATE
        .stateMask = iStateMask
        If mhWnd Then lR = SendMessage(mhWnd, TVM_GETITEM, ZeroL, VarPtr(mtInsert.itemex))
    
        pItem_State = .State And iStateMask
    End With
End Property

Private Property Let pItem_State(ByVal hItem As Long, ByVal iStateMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the state of a given item.
'---------------------------------------------------------------------------------------
    
    lR = ZeroL
    With mtInsert.itemex
        .hItem = hItem
    
        .Mask = TVIF_HANDLE Or TVIF_STATE
        .stateMask = iStateMask
        .State = iNew
    
        If mhWnd Then lR = SendMessage(mhWnd, TVM_SETITEM, ZeroL, VarPtr(mtInsert.itemex))
    End With
End Property

Private Property Get pItem_Text(ByVal hItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the text of a given item.
'---------------------------------------------------------------------------------------
    lR = ZeroL
    With mtInsert.itemex
        .hItem = hItem
        
        .Mask = TVIF_HANDLE Or TVIF_TEXT
        .pszText = StrPtr(msTextBuffer)
        .cchTextMax = LenB(msTextBuffer)
        
        If mhWnd Then lR = SendMessage(mhWnd, TVM_GETITEM, ZeroL, VarPtr(mtInsert.itemex))
    
        lstrToStringA .pszText, pItem_Text
    End With
End Property

Private Property Let pItem_Text(ByVal hItem As Long, ByVal sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of a given item.
'---------------------------------------------------------------------------------------
    lR = ZeroL
    With mtInsert.itemex
        .hItem = hItem
    
        .Mask = TVIF_HANDLE Or TVIF_TEXT
    
        msTextBuffer = sNew & vbNullChar
        MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = vbNullChar
        .pszText = StrPtr(msTextBuffer)
        .cchTextMax = LenB(msTextBuffer)
        
        If mhWnd Then lR = SendMessage(mhWnd, TVM_SETITEM, ZeroL, VarPtr(mtInsert.itemex))
    End With
End Property

Private Function pItem_hItem(ByRef vItem As Variant) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the hItem from a key, object or hitem.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    If VarType(vItem) = vbObject Then
        
        Dim loItem As cNode
        Set loItem = vItem
        
        If loItem.fIsMine(Me) Then pItem_hItem = loItem.fhItem
        
    ElseIf VarType(vItem) = vbString Then
        
       pItem_hItem = pItem_FindKey(CStr(vItem))
        
    Else
        
        pItem_hItem = CLng(vItem)
        
        If pItem_Verify(pItem_hItem) = False Then
            pItem_hItem = ZeroL
        End If
        
    End If
    On Error GoTo 0
End Function

Private Function pItem_FindKey(ByRef sKey As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the hItem from a key.
'---------------------------------------------------------------------------------------
    
    If LenB(sKey) Then
        Dim lsAnsi As String:   lsAnsi = StrConv(sKey & vbNullChar, vbFromUnicode)
        Dim lpAnsi As Long:     lpAnsi = StrPtr(lsAnsi)
        If Not moKeyMap Is Nothing Then
            pItem_FindKey = moKeyMap.Find(lpAnsi, Hash(lpAnsi, lstrlen(lpAnsi)))
            If pItem_FindKey Then pItem_FindKey = pcNode_hItem(pItem_FindKey)
        End If
    End If
End Function

Private Function pItem_Add( _
        ByVal hItem As Long, _
        ByVal hInsertAfter As Long, _
        ByRef sKey As String, _
        ByRef sText As String, _
        ByVal iIconIndex As Long, _
        ByVal iIconIndexSelected As Long, _
        ByVal iIconIndexState As Long, _
        ByVal iItemData As Long, _
        ByVal iItemNumber As Long, _
        ByVal bShowPlusMinus As Boolean) _
            As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add a node to the treeview.
'---------------------------------------------------------------------------------------
If mhWnd Then
    With mtInsert
        .hParent = hItem
        .hInsertAfter = hInsertAfter
        
        Dim lsAnsi As String
        Dim lpNode As Long
        
        lpNode = pcNode_Initialize(sKey, iItemData, iItemNumber)
        
        With .itemex
            .Mask = TVIF_CHILDREN Or TVIF_PARAM
            .stateMask = ZeroL
            .cChildren = CLng(-bShowPlusMinus)
            .lParam = lpNode
            
            If LenB(sText) Then
                .Mask = .Mask Or TVIF_TEXT
                sText = StrConv(sText & vbNullChar, vbFromUnicode)
                MidB$(msTextBuffer, OneL, LenB(sText)) = sText
                MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = vbNullChar
                .pszText = StrPtr(msTextBuffer)
                .cchTextMax = LenB(sText)
            End If
            
            If iIconIndex > NegOneL Then
                .Mask = .Mask Or TVIF_IMAGE
                .iImage = iIconIndex
            End If
            
            If iIconIndexSelected > NegOneL Then
                .Mask = .Mask Or TVIF_SELECTEDIMAGE
                .iSelectedImage = iIconIndexSelected
            ElseIf iIconIndex > NegOneL Then
                .Mask = .Mask Or TVIF_SELECTEDIMAGE
                .iSelectedImage = iIconIndex
            End If
            
            If iIconIndexState > NegOneL Then
                .stateMask = .stateMask Or TVIS_STATEIMAGEMASK
                .State = .State Or ((iIconIndexState And &HF) * &H1000&)
            End If
            
                
            hItem = SendMessage(mhWnd, TVM_INSERTITEM, ZeroL, VarPtr(mtInsert))
            
            If hItem Then
            
                pcNode_hItem(lpNode) = hItem

                Set pItem_Add = New cNode
                pItem_Add.fInit Me, hItem
                
            Else
                Debug.Assert False
                
                pcNode_Terminate lpNode
                
            End If
            
        End With
    End With
End If
End Function










Public Property Get ColorBack() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the backcolor of the treeview.
'---------------------------------------------------------------------------------------
    ColorBack = miBackColor
End Property
Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the backcolor of the treeview.
'---------------------------------------------------------------------------------------
    miBackColor = iNew
    If mhWnd Then SendMessage mhWnd, TVM_SETBKCOLOR, ZeroL, TranslateColor(iNew)
    pPropChanged PROP_BackColor
End Property

Public Property Get BorderStyle() As evbComCtlBorderStyle
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the borderstyle of the treeview.
'---------------------------------------------------------------------------------------
    BorderStyle = miBorderStyle
End Property
Public Property Let BorderStyle(ByVal iNew As evbComCtlBorderStyle)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the borderstyle of the treeview.
'---------------------------------------------------------------------------------------
    miBorderStyle = iNew
    pSetBorder
    pPropChanged PROP_BorderStyle
End Property


Public Property Get CheckBoxes() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether checkboxes are displayed next to the node icons.
'---------------------------------------------------------------------------------------
    CheckBoxes = CBool(miStyle And bpCheckBoxes)
End Property
Public Property Let CheckBoxes(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether checkboxes are displayed next to the node icons.
'---------------------------------------------------------------------------------------
    If CBool(mhWnd) And (bNew Xor CheckBoxes()) Then
        
        Dim lhImlState As Long
        lhImlState = SendMessage(mhWnd, TVM_GETIMAGELIST, TVSIL_STATE, ZeroL)
        pSetStyle bpCheckBoxes, bNew
        
        If bNew Then
            If mhImlCheckBoxes Then
                If mhImlCheckBoxes <> SendMessage(mhWnd, TVM_GETIMAGELIST, TVSIL_STATE, ZeroL) Then
                    ImageList_Destroy mhImlCheckBoxes
                    mhImlCheckBoxes = ZeroL
                End If
            End If
            mhImlCheckBoxes = SendMessage(mhWnd, TVM_GETIMAGELIST, TVSIL_STATE, ZeroL)
        Else
            SendMessage mhWnd, TVM_SETIMAGELIST, TVSIL_STATE, ZeroL
        End If
        
    End If
End Property

Public Property Get ShowNumbers() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether ItemNumbers can be displayed next to each node.
'---------------------------------------------------------------------------------------
    ShowNumbers = mbShowNumbers
End Property
Public Property Let ShowNumbers(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether ItemNumbers can be displayed next to each node.
'---------------------------------------------------------------------------------------
    mbShowNumbers = bNew
    pPropChanged PROP_ShowNumbers
    If mhWnd Then InvalidateRect mhWnd, ByVal ZeroL, OneL
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether the control is enabled.
'---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the control is enabled.
'---------------------------------------------------------------------------------------
    UserControl.Enabled = bNew
    If mhWnd Then
        EnableWindow mhWnd, -CLng(bNew)
        pSetStyle WS_DISABLED, Not bNew
        EnableWindow mhWnd, -CLng(bNew)
        InvalidateRect mhWnd, ByVal ZeroL, OneL
    End If
End Property

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the font used by this control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the font used by this control.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    pSetFont
    pPropChanged PROP_Font
End Property


Public Property Get ColorFore() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the color of the text used by the control.
'---------------------------------------------------------------------------------------
    ColorFore = miForeColor
End Property
Public Property Let ColorFore(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the color of the text used by the control.
'---------------------------------------------------------------------------------------
    miForeColor = iNew
    SendMessage mhWnd, TVM_SETTEXTCOLOR, ZeroL, TranslateColor(iNew)
    pPropChanged PROP_ForeColor
End Property

Public Property Get LinesAtRoot() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether lines are shown to the root node.
'---------------------------------------------------------------------------------------
    LinesAtRoot = (miStyle And bpLinesAtRoot)
End Property
Public Property Let LinesAtRoot(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether lines are shown to the root node.
'---------------------------------------------------------------------------------------
   pSetStyle bpLinesAtRoot, bNew
End Property

Public Property Get HasButtons() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get whether the treeview displays expand/collapse buttons next to the nodes.
'---------------------------------------------------------------------------------------
    HasButtons = (miStyle And bpHasButtons)
End Property
Public Property Let HasButtons(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the treeview displays expand/collapse buttons next to the nodes.
'---------------------------------------------------------------------------------------
   pSetStyle bpHasButtons, bNew
End Property

Public Property Get HasLines() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get whether the treeview has lines between parent and child nodes.
'---------------------------------------------------------------------------------------
    HasLines = (miStyle And bpHasLines)
End Property
Public Property Let HasLines(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the treeview has lines between parent and child nodes.
'---------------------------------------------------------------------------------------
   pSetStyle bpHasLines, bNew
End Property

Public Property Get HideSelection() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value indicating whether the selection is hidden when the control is
'             not in focus.
'---------------------------------------------------------------------------------------
    HideSelection = Not CBool(miStyle And bpShowSelAlways)
End Property
Public Property Let HideSelection(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the selection is hidden when the control is not in focus.
'---------------------------------------------------------------------------------------
    pSetStyle bpShowSelAlways, Not bNew
    If mhWnd Then InvalidateRect mhWnd, ByVal ZeroL, OneL
End Property

Public Property Get HotTrack() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value indicating whether nodes change color as the mouse passes over them.
'---------------------------------------------------------------------------------------
    HotTrack = CBool(miStyle And bpTrackSelect)
End Property
Public Property Let HotTrack(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether nodes change color as the mouse passes over them.
'---------------------------------------------------------------------------------------
    pSetStyle bpTrackSelect, bNew
End Property

Public Property Get HitTest(ByVal x As Single, ByVal y As Single, Optional ByRef iInfo As eTreeViewHitTest) As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a reference to the node at the given coordinates.
'---------------------------------------------------------------------------------------
   Dim tHT As TVHITTESTINFO
    
    With tHT
        With .pt
            .x = UserControl.ScaleX(x, vbContainerPosition, vbPixels)
            .y = UserControl.ScaleY(y, vbContainerPosition, vbPixels)
        End With
        
        If mhWnd Then SendMessage mhWnd, TVM_HITTEST, ZeroL, VarPtr(tHT)
        
        iInfo = .Flags
        If .hItem Then
            Set HitTest = New cNode
            HitTest.fInit Me, .hItem
        End If
        
    End With
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndTreeView() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hwnd to the treeview.
'---------------------------------------------------------------------------------------
    If mhWnd Then hWndTreeView = mhWnd
End Property

Public Property Get ImageList() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the imagelist used by this control.
'---------------------------------------------------------------------------------------
    Set ImageList = moImageList
End Property

Public Property Set ImageList(ByVal oNew As cImageList)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the imagelist used by this control.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Set moImageList = Nothing
    Set moImageListEvent = Nothing
    Set moImageList = oNew
    Set moImageListEvent = Nothing
    On Error GoTo 0
    If mhWnd Then
        If Not moImageList Is Nothing _
            Then SendMessage mhWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, moImageList.hIml _
            Else SendMessage mhWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, ZeroL
    End If
End Property

Public Property Get Indentation() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the indentation between each level of nodes.
'---------------------------------------------------------------------------------------
    Indentation = ScaleX(miIndent, vbPixels, vbContainerSize)
End Property
Public Property Let Indentation(ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the indentation between each level of nodes.
'---------------------------------------------------------------------------------------
    miIndent = ScaleX(fNew, vbContainerSize, vbPixels)
    If mhWnd Then SendMessage mhWnd, TVM_SETINDENT, miIndent, ZeroL
    pPropChanged PROP_Indent
End Property

Public Property Get ItemHeight() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the height of each node.
'---------------------------------------------------------------------------------------
   ItemHeight = miItemHeight
End Property
Public Property Let ItemHeight(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the height of each node.
'---------------------------------------------------------------------------------------
    miItemHeight = iNew
    If mhWnd Then SendMessage mhWnd, TVM_SETITEMHEIGHT, iNew, ZeroL
    pPropChanged PROP_ItemHeight
End Property

Public Property Get LabelEdit() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether nodes use automatic label editing.
'---------------------------------------------------------------------------------------
    LabelEdit = CBool(miStyle And bpEditLabels)
End Property
Public Property Let LabelEdit(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether nodes use automatic label editing.
'---------------------------------------------------------------------------------------
    pSetStyle bpEditLabels, bNew
End Property

Public Property Get ColorLine() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the color of the lines between parent and child nodes.
'---------------------------------------------------------------------------------------
    ColorLine = miLineColor
End Property
Public Property Let ColorLine(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the color of the lines between parent and child nodes.
'---------------------------------------------------------------------------------------
    miLineColor = iNew
    If mhWnd Then SendMessage mhWnd, TVM_SETLINECOLOR, 0, TranslateColor(iNew)
    pPropChanged PROP_LineColor
End Property


Public Property Get PathSeparator() As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the path separator used for the FullPath property of the nodes.
'---------------------------------------------------------------------------------------
    PathSeparator = msPathSeparator
End Property
Public Property Let PathSeparator(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the path separator used for the FullPath property of the nodes.
'---------------------------------------------------------------------------------------
    msPathSeparator = sNew
    pPropChanged PROP_PathSeparator
End Property

Public Property Get NoScrollBar() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value indicating whether the scrollbars are hidden.
'---------------------------------------------------------------------------------------
    NoScrollBar = CBool(miStyle And bpNoScroll)
End Property
Public Property Let NoScrollBar(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value indicating whether the scrollbars are hidden.
'---------------------------------------------------------------------------------------
   pSetStyle bpNoScroll, bNew
End Property

Public Property Get SingleExpand() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether only one node may be expanded at a time.
'---------------------------------------------------------------------------------------
    SingleExpand = CBool(miStyle And bpSingleExpand)
End Property
Public Property Let SingleExpand(ByVal bVal As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether only one node may be expanded at a time.
'---------------------------------------------------------------------------------------
    pSetStyle bpSingleExpand, bVal
End Property

Public Property Get DropHighlight() As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the node that is drop highlighted or Nothing otherwise.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Set DropHighlight = pItem(SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_DROPHILITE, TVI_ROOT))
    End If
End Property

Public Property Set DropHighlight(ByVal oNew As cNode)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the node that is drop highlighted.
'---------------------------------------------------------------------------------------
    SetDropHighlight oNew
End Property

Public Sub SetDropHighlight(ByVal vNode As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the node that is drop highlighted.
'---------------------------------------------------------------------------------------
    Dim lhItem As Long
    lhItem = pItem_hItem(vNode)
    If mhWnd Then
        If lhItem <> SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_DROPHILITE, TVI_ROOT) Then
            ImageDrag_Show False
            SendMessage mhWnd, TVM_SELECTITEM, TVGN_DROPHILITE, lhItem
            'pInvalidateItem lhItem
            ImageDrag_Show True
        End If
    End If
End Sub

Public Property Get InsertMark(Optional ByRef bAfter As Boolean) As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the node adjacent to the drop insert mark and the relative location of the
'             insert mark.
'---------------------------------------------------------------------------------------
    Set InsertMark = pItem(mhItemInsertMark)
    bAfter = mbInsertMarkAfter
End Property

Public Property Set InsertMark(Optional ByRef bAfter As Boolean, ByVal oNew As cNode)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the drop insert mark before or after the given node.
'---------------------------------------------------------------------------------------
    SetInsertMark oNew, bAfter
End Property

Public Sub SetInsertMark(ByVal vNode As Variant, Optional ByVal bAfter As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the drop insert mark before or after the given node.
'---------------------------------------------------------------------------------------
    Dim hItem As Long
    hItem = pItem_hItem(vNode)
    
    If CBool(mhWnd) And (hItem <> mhItemInsertMark Or bAfter <> mbInsertMarkAfter) Then
        mhItemInsertMark = hItem
        mbInsertMarkAfter = bAfter
        ImageDrag_Show False
        SendMessage mhWnd, TVM_SETINSERTMARK, CLng(-bAfter), hItem
        pInvalidateItem hItem
        ImageDrag_Show True
    End If
    
End Sub

Private Sub pInvalidateItem(ByVal hItem As Long, Optional ByVal iType As Long = ZeroL)
    Dim tR As RECT
    tR.Left = hItem
    SendMessage mhWnd, TVM_GETITEMRECT, iType, VarPtr(tR)
    InvalidateRect mhWnd, tR, ZeroL
    UpdateWindow mhWnd
End Sub

Public Sub Refresh()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Redraw the treeview.
'---------------------------------------------------------------------------------------
    InvalidateRect mhWnd, ByVal ZeroL, OneL
    UpdateWindow mhWnd
End Sub

Public Property Get Redraw() As Boolean
    Redraw = mbRedraw
End Property

Public Property Let Redraw(ByVal bNew As Boolean)
    mbRedraw = bNew
    SendMessage mhWnd, WM_SETREDRAW, Sgn(bNew), ZeroL
    Me.Refresh
End Property



Public Property Get Root() As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the root node.
'---------------------------------------------------------------------------------------
    Set Root = New cNode
    Root.fInit Me, TVI_ROOT
End Property

Public Property Get Nodes() As cNodes
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the collection of nodes.
'---------------------------------------------------------------------------------------
    Set Nodes = New cNodes
    Nodes.fInit Me
End Property

Public Property Get FocusedItem() As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the node that is in focus.
'---------------------------------------------------------------------------------------
Dim hItem As Long
    If mhWnd Then
        hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CARET, TVI_ROOT)
        If hItem Then
            Set FocusedItem = New cNode
            FocusedItem.fInit Me, hItem
        End If
    End If
End Property

Public Property Set FocusedItem(ByVal oNew As cNode)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the node that is in focus.
'---------------------------------------------------------------------------------------
    SetFocusedItem oNew
End Property

Public Sub SetFocusedItem(ByVal vNode As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the node that is in focus.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim hItem As Long
        hItem = pItem_hItem(vNode)
        If hItem = ZeroL Then gErr vbccKeyOrIndexNotFound, ucTreeView
        SendMessage mhWnd, TVM_SELECTITEM, TVGN_CARET, hItem
    End If
End Sub

Public Property Get FirstVisibleNode() As cNode
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the first visible node.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim hItem As Long
        hItem = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ZeroL)
        If hItem Then
            Set FirstVisibleNode = New cNode
            FirstVisibleNode.fInit Me, hItem
        End If
    End If
End Property

Public Property Set FirstVisibleNode(ByVal oNew As cNode)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the first visible node.
'---------------------------------------------------------------------------------------
    SetFirstVisibleNode oNew
End Property

Public Sub SetFirstVisibleNode(ByVal vNode As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the first visible node.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim hItem As Long
        hItem = pItem_hItem(vNode)
        If hItem = ZeroL Then gErr vbccKeyOrIndexNotFound, ucTreeView
        SendMessage mhWnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, hItem
    
    End If
End Sub

Public Sub StartLabelEdit(ByVal vNode As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Start a label edit on the given node.
'---------------------------------------------------------------------------------------
    Dim hItem As Long
    hItem = pItem_hItem(vNode)
    If hItem = ZeroL Then gErr vbccKeyOrIndexNotFound, ucTreeView
    If mhWnd Then
        If GetFocus() <> mhWnd Then vbComCtlTlb.SetFocus mhWnd
        SendMessage mhWnd, TVM_EDITLABEL, ZeroL, hItem
    End If
End Sub

Public Sub EndLabelEdit(Optional ByVal bDiscardText As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : End any pending label edit operation.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, TVM_ENDEDITLABELNOW, -(Not bDiscardText), ZeroL
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
        pPropChanged PROP_Themeable
        mbThemeable = bNew
        If mhWnd Then
            EnableWindowTheme mhWnd, mbThemeable
            SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
            RedrawWindow mhWnd, ByVal ZeroL, ZeroL, RDW_INVALIDATE
        End If
    End If
End Property

Public Property Get OLERegisterDrop() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 3/5/05
' Purpose   : Return a value indicating whether the control registers itself as an ole
'             drag-drop target.
'---------------------------------------------------------------------------------------
    OLERegisterDrop = -UserControl.OLEDropMode
End Property

Public Property Let OLERegisterDrop(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/5/05
' Purpose   : Set whether the control registers itself as an ole drag-drop target.
'---------------------------------------------------------------------------------------
    UserControl.OLEDropMode = -bNew
    pPropChanged PROP_OleDrop
End Property

Public Sub OLEDrag()
'---------------------------------------------------------------------------------------
' Date      : 3/5/05
' Purpose   : Initiate a drag operation through OLE.  The OLE Drag events are raised for further interaction.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lhNode As Long
        lhNode = SendMessage(mhWnd, TVM_GETNEXTITEM, TVGN_CARET, TVI_ROOT)
        If lhNode Then
        
            Dim tR As RECT
    
            tR.Left = lhNode
            SendMessage mhWnd, TVM_GETITEMRECT, OneL, VarPtr(tR)
            If Not moImageList Is Nothing Then
                tR.Left = tR.Left - moImageList.IconWidth - TwoL
            End If
            
            Dim x As Long, y As Long
            
            x = miLastXMouseDown - tR.Left
            y = miLastYMouseDown - tR.Top
            
            Dim lhIml As Long
            lhIml = SendMessage(mhWnd, TVM_CREATEDRAGIMAGE, ZeroL, lhNode)
            
            If lhIml Then
                
                Dim liIconWidth As Long
                Dim liIconHeight As Long
            
                ImageList_GetIconSize lhIml, liIconWidth, liIconHeight
                
                Dim lhDc As Long
                
                lhDc = CreateCompatibleDC(ZeroL)
                
                If lhDc Then
                
                    Dim loDib As pcDibSection
                    Set loDib = New pcDibSection
                    
                    If loDib.Create(liIconWidth, liIconHeight, lhDc, IIf(ImageDrag_Alpha, 32, 24)) Then
                        Dim lhBmpOld As Long
                        lhBmpOld = SelectObject(lhDc, loDib.hBitmap)
                        If lhBmpOld Then
                            
                            Dim lhBrush As Long
                            lhBrush = GdiMgr_CreateSolidBrush(ImageDrag_TransColor)
                            If lhBrush Then
                                tR.Left = ZeroL
                                tR.Top = ZeroL
                                tR.Right = liIconWidth
                                tR.bottom = liIconHeight
                                FillRect lhDc, tR, lhBrush
                                GdiMgr_DeletePen lhBrush
                            End If
                            
                            ImageList_Draw lhIml, ZeroL, lhDc, ZeroL, ZeroL, ZeroL
                            SelectObject lhDc, lhBmpOld
                            
                            If ImageDrag_Alpha Then pDragDib_FadeAlpha loDib
                            
                            ImageDrag_Start loDib, x, y
                        End If
                    End If
                    DeleteDC lhDc
                End If
                ImageList_Destroy lhIml
            End If
        End If
    End If
    UserControl.OLEDrag
End Sub

Private Sub pDragDib_FadeAlpha(ByVal oDib As pcDibSection)
'---------------------------------------------------------------------------------------
' Date      : 4/20/05
' Purpose   : Premultiply the rgb values with the alpha value that is used for the drag image.
'---------------------------------------------------------------------------------------
    
    Const Bkgnd As Long = ((ImageDrag_TransColor And &HFF0000) \ &H10000) Or _
                          (ImageDrag_TransColor And &HFF00&) Or _
                          ((ImageDrag_TransColor And &HFF&) * &H10000)
    Dim liBkgnd As Long
    liBkgnd = Bkgnd
    
    Dim x As Long
    Dim y As Long
    
    Dim liAlpha As Long
    
    liAlpha = 160&
    
    Dim lyBits() As Byte
    SAPtr(lyBits) = oDib.ArrPtr(1)
    
    For x = LBound(lyBits, 1) To UBound(lyBits, 1) Step 4
        For y = LBound(lyBits, 2) To UBound(lyBits, 2)
            If (MemOffset32(VarPtr(lyBits(x, y)), ZeroL) And &HFFFFFF) = liBkgnd Then
                lyBits(x, y) = ZeroY
                lyBits(x + 1, y) = ZeroY
                lyBits(x + 2, y) = ZeroY
                lyBits(x + 3, y) = ZeroY
            Else
                lyBits(x, y) = liAlpha * lyBits(x, y) \ &HFF&
                lyBits(x + 1, y) = liAlpha * lyBits(x + 1, y) \ &HFF&
                lyBits(x + 2, y) = liAlpha * lyBits(x + 2, y) \ &HFF&
                lyBits(x + 3, y) = liAlpha
            End If
        Next
    Next

    SAPtr(lyBits) = ZeroL
    
End Sub

Public Property Get FindItemData(ByVal iItemData As Long) As cNode
    If Not moItemDataMap Is Nothing Then
        Dim lpNode As Long
        lpNode = moItemDataMap.Find(iItemData, HashLong(iItemData))
        If lpNode Then
            Set FindItemData = New cNode
            FindItemData.fInit Me, pcNode_hItem(lpNode)
        End If
    End If
End Property
