VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucTreeViewTest 
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Align           =   3  'Align Left
      Height          =   6120
      Left            =   0
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   10795
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucTreeViewTest.ctx":0000
         Left            =   720
         List            =   "ucTreeViewTest.ctx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2940
         Width           =   1935
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Expand All"
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Collapse All"
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "HideSelection"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   4
         Top             =   1172
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2580
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Colors ..."
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Font ..."
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Fill List"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtItems 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "1000"
         Top             =   2580
         Width           =   1215
      End
      Begin vbComCtl.ucUpDown ucUpDown1 
         Height          =   300
         Left            =   1005
         Top             =   2580
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Value           =   5
         BProps          =   327718
         Buddy           =   "Text1"
         BuddyProp       =   "Text"
      End
      Begin vbComCtlTest.ucEvents evt 
         Height          =   2055
         Left            =   60
         TabIndex        =   16
         Top             =   3300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   3625
      End
      Begin VB.CheckBox chk 
         Caption         =   "Checkboxes"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   338
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "HasButtons"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   2
         Top             =   616
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "HasLines"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   3
         Top             =   894
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "HotTrack"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   5
         Top             =   1450
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "LabelEdit"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   6
         Top             =   1728
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "LinesAtRoot"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   7
         Top             =   2010
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Border:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Indentation:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "# of Items:"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   18
         Top             =   2340
         Width           =   855
      End
   End
   Begin vbComCtl.ucTreeView tvw 
      Height          =   1575
      Left            =   3000
      TabIndex        =   17
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Style           =   39
      Themeable       =   0   'False
      OleDrop         =   -1  'True
   End
End
Attribute VB_Name = "ucTreeViewTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucTreeViewTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the treeview control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Const OLEDRAG_Treeview As Long = &H1234

Private Enum eChk
    chkCheckboxes
    chkEnabled
    chkHasButtons
    chkHasLines
    chkHotTrack
    chkLabelEdit
    chkLinesAtRoot
    chkHideSelection
End Enum

Private Enum eCmd
    cmdColors
    cmdFill
    cmdFont
    cmdCollapse
    cmdExpand
End Enum

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the font or back color properties to match the container.
'---------------------------------------------------------------------------------------
    Select Case ZeroL
    Case StrComp(PropertyName, "Font")
        Set UserControl.Font = Ambient.Font
    Case StrComp(PropertyName, "BackColor")
        UserControl.BackColor = Ambient.BackColor
        vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
    End Select
End Sub

Private Sub Usercontrol_EnterFocus()
    vbComCtl.EnterFocus Controls
End Sub

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the constituent controls.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowAllUIStates hWnd
    Set tvw.ImageList = gImageListSmall
    With chk
        .Item(chkCheckboxes).Value = -tvw.CheckBoxes
        .Item(chkEnabled).Value = -tvw.Enabled
        .Item(chkHasButtons).Value = -tvw.HasButtons
        .Item(chkHasLines).Value = -tvw.HasLines
        .Item(chkHotTrack).Value = -tvw.HotTrack
        .Item(chkLabelEdit).Value = -tvw.LabelEdit
        .Item(chkLinesAtRoot).Value = -tvw.LinesAtRoot
    End With
    ucUpDown1.Value = tvw.Indentation
    cmb.ListIndex = tvw.BorderStyle
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    Set UserControl.Font = Ambient.Font
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    tvw.Move ucScrollBox1.Width, 0, ScaleWidth - ucScrollBox1.Width, ScaleHeight
    On Error GoTo 0
End Sub

Private Sub cmb_Click()
    tvw.BorderStyle = cmb.ListIndex
End Sub

Private Sub chk_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Change the corresponding property in the treeview.
'---------------------------------------------------------------------------------------
    Dim lbVal As Boolean
    lbVal = chk(Index).Value
    Select Case Index
    Case chkCheckboxes:     tvw.CheckBoxes = lbVal
    Case chkEnabled:        tvw.Enabled = lbVal
    Case chkHasButtons:     tvw.HasButtons = lbVal
    Case chkHasLines:       tvw.HasLines = lbVal
    Case chkHotTrack:       tvw.HotTrack = lbVal
    Case chkLabelEdit:      tvw.LabelEdit = lbVal
    Case chkLinesAtRoot:    tvw.LinesAtRoot = lbVal
    Case chkHideSelection:  tvw.HideSelection = lbVal
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Perform the action associated with each command button.
'---------------------------------------------------------------------------------------
    Select Case Index
    Case cmdColors:     fEditColors.EditColors Parent, tvw, "ColorFore", "ColorBack", "ColorLine"
    Case cmdFont:       tvw.Font.Browse hWnd, dlgFontScreenFonts
    Case cmdCollapse:   pExpand False
    Case cmdExpand:     pExpand True
    Case cmdFill:       pFill txtItems.Text
    End Select
End Sub

Private Sub tvw_AfterLabelEdit(ByVal oNode As vbComCtl.cNode, sNew As String, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "AfterLabelEdit " & sNew
End Sub

Private Sub tvw_BeforeCollapse(ByVal oNode As vbComCtl.cNode, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "BeforeCollapse " & oNode.Text
End Sub

Private Sub tvw_BeforeExpand(ByVal oNode As vbComCtl.cNode, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "BeforeExpand " & oNode.Text
End Sub

Private Sub tvw_BeforeLabelEdit(ByVal oNode As vbComCtl.cNode, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "BeforeLabelEdit " & oNode.Text
End Sub

Private Sub tvw_Click()
    evt.LogItem "Click"
End Sub

Private Sub tvw_Collapse(ByVal oNode As vbComCtl.cNode)
    evt.LogItem "Collapse " & oNode.Text
End Sub

Private Sub tvw_Expand(ByVal oNode As vbComCtl.cNode)
    evt.LogItem "Expand " & oNode.Text
End Sub

Private Sub tvw_KeyDown(ByVal iKeyCode As Integer, ByVal iState As vbComCtl.evbComCtlKeyboardState)
    evt.LogItem "KeyDown " & iKeyCode & ", " & iState
End Sub

Private Sub tvw_NodeCheck(ByVal oNode As vbComCtl.cNode)
    evt.LogItem "NodeCheck " & oNode.Text
End Sub

Private Sub tvw_NodeClick(ByVal oNode As vbComCtl.cNode, ByVal iHitTestCode As vbComCtl.eTreeViewHitTest)
    evt.LogItem "NodeClick " & oNode.Text
End Sub

Private Sub tvw_NodeDblClick(ByVal oNode As vbComCtl.cNode, ByVal iHitTestCode As vbComCtl.eTreeViewHitTest)
    evt.LogItem "NodeDblClick " & oNode.Text
End Sub

Private Sub tvw_NodeRightClick(ByVal oNode As vbComCtl.cNode, ByVal iHitTestCode As vbComCtl.eTreeViewHitTest)
    evt.LogItem "NodeRightClick " & oNode.Text
End Sub

Private Sub tvw_NodeRightDrag(ByVal oNode As vbComCtl.cNode)
    evt.LogItem "NodeRightDrag " & oNode.Text
    tvw.OLEDrag
End Sub

Private Sub tvw_NodeSelect(ByVal oNode As vbComCtl.cNode)
    If oNode Is Nothing _
        Then evt.LogItem "NodeSelect Nothing" _
        Else evt.LogItem "NodeSelect " & oNode.Text
End Sub

Private Sub tvw_RightClick()
    evt.LogItem "RightClick"
End Sub

Private Sub tvw_NodeDrag(ByVal oNode As vbComCtl.cNode)
    evt.LogItem "NodeDrag " & oNode.Text
    tvw.OLEDrag
End Sub

Private Sub tvw_OLECompleteDrag(Effect As vbComCtl.evbComCtlOleDropEffect)
    Set tvw.InsertMark = Nothing
    Set tvw.DropHighlight = Nothing
End Sub

Private Sub tvw_OLEDragDrop(Data As DataObject, Effect As vbComCtl.evbComCtlOleDropEffect, Button As vbComCtl.evbComCtlMouseButton, Shift As vbComCtl.evbComCtlKeyboardState, x As Single, y As Single)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Move the node that is in focus to the insertion mark.
'---------------------------------------------------------------------------------------
    If Not tvw.FocusedItem Is Nothing Then
        
        Dim loInsert As cNode, lbAfter As Boolean
        Set loInsert = tvw.InsertMark(lbAfter)
        Set tvw.InsertMark = Nothing
        
        Dim liRelation As eTreeViewNodeRelation
        liRelation = tvwSibling
        
        If Not lbAfter Then
            If loInsert.GetNode(tvwGetNodePreviousSibling) Is Nothing Then
                Set loInsert = loInsert.GetNode(tvwGetNodeParent)
                liRelation = tvwFirst
            Else
                Set loInsert = loInsert.GetNode(tvwGetNodePreviousSibling)
            End If
        End If
            
        On Error Resume Next
        tvw.Redraw = False
        With tvw.FocusedItem.GetNode(tvwGetNodeParent)
            'this will raise an error if you try to move the focused item to
            'be a decendent of itself.  This should not happen here though
            'since the insertion point is not set unless it is a valid drop site.
            tvw.FocusedItem.Move loInsert, liRelation
            
            'This will raise an error if the old parent node is the root node
            'since the ShowPlusMinus property is not available.
            If Not .HasChildren Then .ShowPlusMinus = False
        End With
        tvw.Redraw = True
        On Error GoTo 0
    Else
        Debug.Assert False
        Set tvw.InsertMark = Nothing
    End If
End Sub

Private Sub tvw_OLEDragOver(Data As DataObject, Effect As vbComCtl.evbComCtlOleDropEffect, Button As vbComCtl.evbComCtlMouseButton, Shift As vbComCtl.evbComCtlKeyboardState, x As Single, y As Single, State As vbComCtl.evbComCtlOleDragOverState)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the insertion mark if the drag is over a valid drop site.
'---------------------------------------------------------------------------------------
    Effect = IIf(Data.GetFormat(OLEDRAG_Treeview), vbccOleDropNone, vbccOleDropMove)
    
    Dim loNode As cNode
    Dim liInfo As eTreeViewHitTest
    Set loNode = tvw.HitTest(x, y, liInfo)
    
    If (liInfo = tvwHitTestItemPlusMinus Or liInfo = tvwHitTestItemText Or liInfo = tvwHitTestItemIcon) And State <> vbccOleDragLeave Then
    
        Dim lbAfter As Boolean
        lbAfter = y > (loNode.Top + loNode.Height \ TwoL)
    
        If Not pIsDecendent(loNode, tvw.FocusedItem) _
           And Not (pIsDecendent(loNode.GetNode(tvwGetNodePreviousSibling), tvw.FocusedItem) And Not lbAfter) _
           And Not (pIsDecendent(loNode.GetNode(tvwGetNodeNextSibling), tvw.FocusedItem) And lbAfter) Then
            Effect = vbccOleDropMove
            tvw.SetInsertMark loNode, lbAfter
            Exit Sub
        End If
    End If
    
    Effect = vbccOleDropNone
    Set tvw.InsertMark = Nothing
End Sub

Private Sub tvw_OLEStartDrag(Data As DataObject, AllowedEffects As vbComCtl.evbComCtlOleDropEffect)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initiate the drag operation.
'---------------------------------------------------------------------------------------
    Dim ly() As Byte
    ly = "Test Data"
    Data.SetData ly, OLEDRAG_Treeview
    AllowedEffects = vbccOleDropMove
End Sub

Private Sub ucScrollBox1_ScrollBarChange()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Adjust the position of the constituent controls.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    ucScrollBox1.Width = 181 + ucScrollBox1.ScrollBarWidth
    tvw.Move ucScrollBox1.Width, 0, ScaleWidth - ucScrollBox1.Width, ScaleHeight
    On Error GoTo 0
End Sub

Private Sub ucUpDown1_Change(ByVal iValue As Long)
    tvw.Indentation = iValue
End Sub

Private Sub pExpand(ByVal bVal As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Expand or collapse all nodes.
'---------------------------------------------------------------------------------------
    Dim loNode As cNode
    tvw.Redraw = False
    For Each loNode In tvw.Nodes
        loNode.Expanded = bVal
    Next
    tvw.Redraw = True
End Sub

Private Sub pFill(ByVal iItems As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Add nodes containing random numbers of children until the total number
'             of nodes is the specified value.
'---------------------------------------------------------------------------------------
    Dim liCount As Long
    Dim liTotalCount As Long
    Dim loNode As cNode
    
    liTotalCount = iItems
    
    Randomize
    tvw.Redraw = False
    tvw.Nodes.Clear
    
    Do While liCount < liTotalCount
        liCount = liCount + OneL
        Set loNode = tvw.Nodes.Add(, , "This is a key " & liCount, "Node Item " & liCount, RandIcon(gImageListSmall), , , liCount)
        pFillSub loNode, liCount, liTotalCount
    Loop
    tvw.Redraw = True
End Sub

Private Sub pFillSub(ByVal oNode As cNode, ByRef iCount As Long, ByVal iLimit As Long, Optional ByVal iNesting As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Recursively add a random number of nodes as children of the given
'             node, not exceeding a certain number of total nodes in the treeview.
'---------------------------------------------------------------------------------------
    Dim i As Long
    If iNesting < 10 Then
        If Int(Rnd * iNesting) = 0 Then
            For i = 0 To Rnd * 5
                If iCount < iLimit Then
                    If Not oNode.ShowPlusMinus Then oNode.ShowPlusMinus = True
                    iCount = iCount + OneL
                    pFillSub oNode.AddChildNode(, "This is a key " & iCount, "Node Item " & iCount, RandIcon(gImageListSmall), , , iCount), iCount, iLimit, iNesting + OneL
                End If
            Next
        End If
    End If
End Sub

Private Function pIsDecendent(ByVal oNode As cNode, ByVal oNodePotentialParent As cNode)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a value indicating whether the given node's parent chain can
'             be traced to the given node.
'---------------------------------------------------------------------------------------
    Do Until oNode Is Nothing
        pIsDecendent = CBool(oNode.hItem = oNodePotentialParent.hItem)
        If pIsDecendent Then Exit Do
        Set oNode = oNode.GetNode(tvwGetNodeParent)
    Loop
End Function

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Public Sub StressTest()
    Dim i As Long
    
    Dim liItems As Long
    liItems = 100
    For i = 1 To 2
        pFill liItems
        
        Dim j As Long
        For j = liItems To OneL Step NegOneL
            Debug.Assert tvw.FindItemData(j).Key = "This is a key " & j
            tvw.Nodes.Remove "This is a key " & j
        Next
        Debug.Assert tvw.Nodes.Count = ZeroL
        
        If KeyIsDown(VK_ESCAPE) Then Exit For
    Next
    
End Sub
