VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucListViewTest 
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   7635
   Begin vbComCtl.ucListView lvw 
      Height          =   1815
      Left            =   4680
      TabIndex        =   43
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Style           =   520
      StyleEx         =   1024
      IconSpaceX      =   100
      IconSpaceY      =   60
      BackColor       =   16777215
      ForeColor       =   0
      BackX           =   50
      BackY           =   50
      TileLines       =   3
   End
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   6795
      Left            =   0
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   11986
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   0
         ScaleHeight     =   6735
         ScaleWidth      =   4575
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Width           =   4575
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   39
            Text            =   "60"
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   38
            Text            =   "100"
            Top             =   3960
            Width           =   495
         End
         Begin VB.CheckBox chk 
            Caption         =   "Use Sys Imagelist"
            Height          =   255
            Index           =   25
            Left            =   0
            TabIndex        =   25
            Top             =   6495
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "BackPictureTile"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   0
            Top             =   60
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "BorderSelect"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   1
            Top             =   317
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "CheckBoxes"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   2
            Top             =   574
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "DoubleBuffer"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   3
            Top             =   831
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "Enabled"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   4
            Top             =   1088
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "FlatScrollBar"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   5
            Top             =   1345
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "FullRowSelect"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   6
            Top             =   1602
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "Gridlines"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   7
            Top             =   1859
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "HeaderButtons"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   8
            Top             =   2116
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "HeaderDragDrop"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   9
            Top             =   2373
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "HeaderHotTrack"
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   10
            Top             =   2630
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "HeaderTrackSize"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   11
            Top             =   2887
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "HideColumnHdrs"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   12
            Top             =   3144
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "HideSelection"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   13
            Top             =   3401
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "InfoTips"
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   14
            Top             =   3658
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "LabelEdit"
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   15
            Top             =   3915
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "LabelTips"
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   16
            Top             =   4172
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "LabelWrap"
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   17
            Top             =   4429
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "MultiSelect"
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   18
            Top             =   4686
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "OneClickActivate"
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   19
            Top             =   4943
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "ShowSortArrow"
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   20
            Top             =   5200
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "SubItemImages"
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   21
            Top             =   5457
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "TrackSelect"
            Height          =   255
            Index           =   22
            Left            =   0
            TabIndex        =   22
            Top             =   5714
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "UnderlineHot"
            Height          =   255
            Index           =   23
            Left            =   0
            TabIndex        =   23
            Top             =   5971
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Colors ..."
            Height          =   375
            Index           =   0
            Left            =   3240
            TabIndex        =   40
            Top             =   3720
            Width           =   1335
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Font ..."
            Height          =   375
            Index           =   1
            Left            =   3240
            TabIndex        =   41
            Top             =   4200
            Width           =   1335
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Fill List"
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   34
            Top             =   2640
            Width           =   1335
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   0
            ItemData        =   "ucListViewTest.ctx":0000
            Left            =   2400
            List            =   "ucListViewTest.ctx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1440
            Width           =   2175
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   1
            ItemData        =   "ucListViewTest.ctx":0043
            Left            =   2400
            List            =   "ucListViewTest.ctx":0050
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1800
            Width           =   2175
         End
         Begin VB.ComboBox cmb 
            Height          =   315
            Index           =   2
            ItemData        =   "ucListViewTest.ctx":0083
            Left            =   2400
            List            =   "ucListViewTest.ctx":0096
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "Use Groups"
            Height          =   255
            Index           =   24
            Left            =   0
            TabIndex        =   24
            Top             =   6228
            Width           =   1215
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   35
            Text            =   "50"
            Top             =   3300
            Width           =   495
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   36
            Text            =   "50"
            Top             =   3300
            Width           =   495
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   2
            Left            =   3240
            TabIndex        =   37
            Text            =   "1000"
            Top             =   3300
            Width           =   1335
         End
         Begin VB.CommandButton cmd 
            Caption         =   "BackPicture ..."
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   33
            Top             =   2640
            Width           =   1335
         End
         Begin vbComCtl.ucFrame ucFrame1 
            Height          =   1335
            Left            =   1680
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2355
            BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
            EndProperty
            Caption         =   "Selected Item:"
            Begin VB.TextBox txtItem 
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   840
               TabIndex        =   27
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox txtItem 
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   840
               TabIndex        =   26
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtItem 
               Enabled         =   0   'False
               Height          =   300
               Index           =   2
               Left            =   840
               TabIndex        =   28
               Top             =   960
               Width           =   735
            End
            Begin VB.PictureBox picDropTarget 
               HasDC           =   0   'False
               Height          =   300
               Left            =   1920
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   240
               ScaleWidth      =   795
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   960
               Width           =   855
            End
            Begin vbComCtl.ucUpDown udIcon 
               Height          =   285
               Left            =   1560
               Top             =   960
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               BProps          =   327719
               Buddy           =   "txtItem(2)"
               Enabled         =   0   'False
               BuddyProp       =   "Text"
            End
            Begin VB.Label Label1 
               Caption         =   "Text:"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Icon:"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   46
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "ToolTip:"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   45
               Top             =   600
               Width           =   615
            End
         End
         Begin vbComCtlTest.ucEvents evt 
            Height          =   2055
            Left            =   1680
            TabIndex        =   42
            Top             =   4680
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   3625
         End
         Begin VB.Label Label1 
            Caption         =   "Icon space x, y"
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   53
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Border:"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   52
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Arrange:"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   51
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "View:"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   50
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "BackPicture X, Y"
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   49
            Top             =   3060
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "# of Items:"
            Height          =   255
            Index           =   7
            Left            =   3240
            TabIndex        =   48
            Top             =   3060
            Width           =   855
         End
      End
   End
   Begin vbComCtl.ucComDlg dlg 
      Left            =   4980
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "ucListViewTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucListViewTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the listview control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Const OLEDRAG_Listview As Long = &H1234 'DataObject format value

Private Enum eChk       'checkbox corresponding to listview properties
    chkBackPictureTile
    chkBorderSelect
    chkCheckboxes
    chkDoubleBuffer
    chkEnabled
    chkFlatScrollBar
    chkFullRowSelect
    chkGridLines
    chkHeaderButtons
    chkHeaderDragDrop
    chkHeaderHotTrack
    chkHeaderTrackSize
    chkHideColumnHeaders
    chkHideSelection
    chkInfoTips
    chkLabelEdit
    chkLabelTips
    chkLabelWrap
    chkMultiSelect
    chkOneClickActivate
    chkShowSortArrow
    chkSubItemImages
    chkTrackSelect
    chkUnderlineHot
    chkUseGroups
    chkUseSystemImageList
End Enum

Private Enum eTxt       'textboxes
    txtBackPictureX
    txtBackPictureY
    txtItems
    txtIconSpaceX
    txtIconSpaceY
End Enum

Private Enum eCmb       'comboboxes
    cmbBorderStyle
    cmbAutoArrange
    cmbView
End Enum

Private Enum eCmd       'commandbuttons
    cmdColors
    cmdFont
    cmdFillList
    cmdBackPicture
End Enum

Private Enum eTxtItem   'more textboxes
    txtItemText
    txtItemToolTip
    txtItemIcon
End Enum

Private mbChanging As Boolean   'updating the controls, ignore the change events.

Private Sub ucScrollBox1_ScrollBarChange()
    On Error Resume Next
    With ucScrollBox1
        .Width = 4635 + .ScrollBarWidth
        lvw.Move .Width, 0, Width - .Width, Height
    End With
    On Error GoTo 0
End Sub

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
    
    With lvw
        .Columns.Add("Col 1", "Column1", 1).ImageOnRight = True
        .Columns.Add "Col 2", "Column2", 2, lvwSortCurrency, lvwAlignLeft, , "currency"
        .Columns.Add "Col 3", "Column3", 3
        .Columns.Add "Col 4", "Column4", 4, lvwSortDate, lvwAlignRight, , "MM/dd/YYYY"
        
        Set .ImageList(lvwImageHeaderImages) = gImageListSmall
        Set .ImageList(lvwImageLargeIcon) = gImageListLarge
        Set .ImageList(lvwImageSmallIcon) = gImageListSmall
        
        Dim lbCC6 As Boolean: lbCC6 = Not .ItemGroups Is Nothing
        
        chk(chkUseGroups).Enabled = lbCC6
        chk(chkShowSortArrow).Enabled = lbCC6
        chk(chkBorderSelect).Enabled = lbCC6
        
        If Not lbCC6 Then cmb(cmbView).RemoveItem lvwTile
    End With
    
    pShowInfo
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
    ucScrollBox1.Height = Height
    lvw.Move ucScrollBox1.Width, lvw.Top, Width - ucScrollBox1.Width, Height
    On Error GoTo 0
End Sub

Private Sub pShowInfo()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the ui to match the current properties of the listview.
'---------------------------------------------------------------------------------------
    mbChanging = True
    On Error Resume Next
    
    With lvw
        chk(chkBackPictureTile).Value = Abs(.BackPictureTile)
        chk(chkBorderSelect).Value = Abs(.BorderSelect)
        chk(chkCheckboxes).Value = Abs(.CheckBoxes)
        chk(chkDoubleBuffer).Value = Abs(.DoubleBuffer)
        chk(chkEnabled).Value = Abs(.Enabled)
        chk(chkFlatScrollBar).Value = Abs(.FlatScrollBar)
        chk(chkFullRowSelect).Value = Abs(.FullRowSelect)
        chk(chkGridLines).Value = Abs(.GridLines)
        chk(chkHeaderButtons).Value = Abs(.HeaderButtons)
        chk(chkHeaderDragDrop).Value = Abs(.HeaderDragDrop)
        chk(chkHeaderHotTrack).Value = Abs(.HeaderHotTrack)
        chk(chkHeaderTrackSize).Value = Abs(.HeaderTrackSize)
        chk(chkHideColumnHeaders).Value = Abs(.HideColumnHeaders)
        chk(chkHideSelection).Value = Abs(.HideSelection)
        chk(chkInfoTips).Value = Abs(.InfoTips)
        chk(chkLabelEdit).Value = Abs(.LabelEdit)
        chk(chkLabelTips).Value = Abs(.LabelTips)
        chk(chkLabelWrap).Value = Abs(.LabelWrap)
        chk(chkMultiSelect).Value = Abs(.MultiSelect)
        chk(chkOneClickActivate).Value = Abs(.OneClickActivate)
        chk(chkShowSortArrow).Value = Abs(.ShowSortArrow)
        chk(chkSubItemImages).Value = Abs(.SubItemImages)
        chk(chkTrackSelect).Value = Abs(.TrackSelect)
        chk(chkUnderlineHot).Value = Abs(.UnderlineHot)
        
        txt(txtBackPictureX).Text = .BackPictureXOffset
        txt(txtBackPictureY).Text = .BackPictureYOffset
        
        cmb(cmbBorderStyle).ListIndex = .BorderStyle
        cmb(cmbAutoArrange).ListIndex = .AutoArrange
        cmb(cmbView).ListIndex = .View
        
    End With
    
    On Error GoTo 0
    mbChanging = False
End Sub

Private Sub chk_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Change the appropriate property of the listview.
'---------------------------------------------------------------------------------------
    If mbChanging Then Exit Sub
    
    Dim lbVal As Boolean
    Dim loItem As cListItem
    
    lbVal = chk(Index).Value
    With lvw
        Select Case Index
        Case chkBackPictureTile:        .BackPictureTile = lbVal
        Case chkBorderSelect:           .BorderSelect = lbVal
        Case chkCheckboxes:             .CheckBoxes = lbVal
        Case chkDoubleBuffer:           .DoubleBuffer = lbVal
        Case chkEnabled:                .Enabled = lbVal
        Case chkFlatScrollBar:          .FlatScrollBar = lbVal
        Case chkFullRowSelect:          .FullRowSelect = lbVal
        Case chkGridLines:              .GridLines = lbVal
        Case chkHeaderButtons:          .HeaderButtons = lbVal
        Case chkHeaderDragDrop:         .HeaderDragDrop = lbVal
        Case chkHeaderHotTrack:         .HeaderHotTrack = lbVal
        Case chkHeaderTrackSize:        .HeaderTrackSize = lbVal
        Case chkHideColumnHeaders:      .HideColumnHeaders = lbVal
        Case chkHideSelection:          .HideSelection = lbVal
        Case chkInfoTips:               .InfoTips = lbVal
        Case chkLabelEdit:              .LabelEdit = lbVal
        Case chkLabelTips:              .LabelTips = lbVal
        Case chkLabelWrap:              .LabelWrap = lbVal
        Case chkMultiSelect:            .MultiSelect = lbVal
        Case chkOneClickActivate:       .OneClickActivate = lbVal
        Case chkShowSortArrow:          .ShowSortArrow = lbVal
        Case chkSubItemImages:          .SubItemImages = lbVal
        Case chkTrackSelect:            .TrackSelect = lbVal
        Case chkUnderlineHot:           .UnderlineHot = lbVal
        Case chkUseGroups
            If Not .ItemGroups Is Nothing Then
                .Redraw = False
                If lbVal Then
                    With .ItemGroups
                        .Clear
                        .Add "Group 1", "Group 1"
                        .Add "Group 2", "Group 2"
                        .Add "Group 3", "Group 3"
                        .Add "Group 4", "Group 4"
                        .Add "Group 5", "Group 5"
                    End With
                    
                    For Each loItem In .ListItems
                        Set loItem.Group = .ItemGroups("Group " & ((loItem.Index Mod 5) + 1))
                    Next
                End If
                
                .ItemGroups.Enabled = lbVal
                .Redraw = True
            End If
        Case chkUseSystemImageList
            lvw.Redraw = False
            Dim liCount As Long
            If lbVal Then
                Set lvw.ImageList(lvwImageLargeIcon) = gSysImageListLarge
                Set lvw.ImageList(lvwImageSmallIcon) = gSysImageListSmall
            Else
                Set lvw.ImageList(lvwImageLargeIcon) = gImageListLarge
                Set lvw.ImageList(lvwImageSmallIcon) = gImageListSmall
            End If
            liCount = lvw.ImageList(lvwImageLargeIcon).IconCount - OneL
            udIcon.Max = liCount
            For Each loItem In lvw.ListItems
                loItem.IconIndex = Rnd * liCount
            Next
            lvw.Redraw = True
        End Select
    End With
End Sub

Private Sub cmb_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Change the appropriate property of the listview.
'---------------------------------------------------------------------------------------
    If mbChanging Then Exit Sub
    Dim liIndex As Long
    liIndex = cmb(Index).ListIndex
    With lvw
        Select Case Index
        Case cmbBorderStyle:    .BorderStyle = liIndex
        Case cmbAutoArrange:    .AutoArrange = liIndex
        Case cmbView:           .View = liIndex
        End Select
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Perform the action associated with each command button.
'---------------------------------------------------------------------------------------
    With lvw
        Select Case Index
        Case cmdColors: fEditColors.EditColors Parent, lvw, "ColorBack", "ColorFore"
        Case cmdFont:   .Font.Browse hWnd
        Case cmdFillList: pFill txt(txtItems).Text

        Case cmdBackPicture
            Dim lsFile As String
            If dlg.ShowFileOpen(lsFile, dlgFileExplorerStyle Or dlgFileHideReadOnly, dlg.FileGetFilter("Bitmap Files (*.bmp)", "*.bmp", "JPG Files (*.jpg;*.jpeg)", "*.jpg;*.jpeg", "GIF Files (*.gif)", "*.gif", "All Files (*.*)|*.*"), , , , , "Choose a background picture") Then
                .BackPictureURL = lsFile
            Else
                .BackPictureURL = vbNullString
            End If
        End Select
    End With
End Sub

Private Sub lvw_Click(ByVal iButton As evbComCtlMouseButton)
    evt.LogItem "Click " & iButton
End Sub

Private Sub lvw_ColumnAfterDrag(ByVal oColumn As vbComCtl.cColumn, iNewPosition As Long, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "AfterDrag " & oColumn.Text & " Position " & oColumn.Position
End Sub

Private Sub lvw_ColumnAfterSize(ByVal oColumn As vbComCtl.cColumn)
    evt.LogItem "AfterSize " & oColumn.Text & " Size " & oColumn.Width
End Sub

Private Sub lvw_ColumnBeforeDrag(ByVal oColumn As vbComCtl.cColumn, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "BeforeDrag " & oColumn.Text & " Position " & oColumn.Position
End Sub

Private Sub lvw_ColumnBeforeSize(ByVal oColumn As vbComCtl.cColumn, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "BeforeSize " & oColumn.Text & " Size " & oColumn.Width
End Sub

Private Sub lvw_ColumnClick(ByVal oColumn As vbComCtl.cColumn)
    evt.LogItem "Click " & oColumn.Text
    oColumn.Sort
    Set lvw.SelectedColumn = oColumn
End Sub

Private Sub lvw_ContextMenu(ByVal x As Single, ByVal y As Single)
    evt.LogItem "ContextMenu " & x & ", " & y
End Sub

Private Sub lvw_ItemActivate(ByVal oItem As vbComCtl.cListItem)
    evt.LogItem "Activate " & oItem.Text
End Sub

Private Sub lvw_ItemAfterEdit(ByVal oItem As vbComCtl.cListItem, bCancel As stdole.OLE_CANCELBOOL, sNew As String)
    evt.LogItem "AfterEdit " & oItem.Text
End Sub

Private Sub lvw_ItemBeforeEdit(ByVal oItem As vbComCtl.cListItem, bCancel As stdole.OLE_CANCELBOOL)
    evt.LogItem "BeforeEdit " & oItem.Text
End Sub

Private Sub lvw_ItemCheck(ByVal oItem As vbComCtl.cListItem, ByVal bCheck As Boolean)
    evt.LogItem "Check " & oItem.Checked & " Value " & bCheck
End Sub

Private Sub lvw_ItemClick(ByVal oItem As vbComCtl.cListItem, ByVal iButton As vbComCtl.evbComCtlMouseButton)
    evt.LogItem "Click " & oItem.Text & " Button " & iButton
End Sub

Private Sub lvw_ItemDrag(ByVal oItem As vbComCtl.cListItem, ByVal iButton As vbComCtl.evbComCtlMouseButton)
    evt.LogItem "Drag " & oItem.Text & " Button " & iButton
    lvw.OLEDrag
End Sub

Private Sub lvw_ItemFocus(ByVal oItem As vbComCtl.cListItem)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the ui to match the current selected item.
'---------------------------------------------------------------------------------------
    
    With txtItem
        .Item(txtItemText).Tag = vbNullString
        
        Dim lbEnabled As Boolean: lbEnabled = Not oItem Is Nothing
        
        If lbEnabled Then
            udIcon.Value = oItem.IconIndex
            .Item(txtItemText).Text = oItem.Text
            .Item(txtItemToolTip).Text = oItem.ToolTipText
            evt.LogItem "Focus " & oItem.Text
            .Item(txtItemText).Tag = oItem.Index
        Else
            .Item(txtItemText).Text = vbNullString
            .Item(txtItemToolTip).Text = vbNullString
            .Item(txtItemIcon).Text = vbNullString
            evt.LogItem "Focus Nothing"
        End If
        
        .Item(txtItemText).Enabled = lbEnabled
        .Item(txtItemToolTip).Enabled = lbEnabled
        .Item(txtItemIcon).Enabled = lbEnabled
        udIcon.Enabled = lbEnabled
    End With
End Sub

Private Sub lvw_KeyDown(iKeyCode As Integer, ByVal iState As vbComCtl.evbComCtlKeyboardState, ByVal bRepeat As Boolean)
    evt.LogItem "KeyDown " & iKeyCode & " State " & iState & " Repeat " & bRepeat
End Sub

Private Sub lvw_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Dim lyData() As Byte
    lyData = "Test Data"
    Data.SetData lyData, OLEDRAG_Listview
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
End Sub

Private Sub picDropTarget_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Data.GetFormat(OLEDRAG_Listview) Then evt.LogItem "OLE Drop " & Data.GetData(OLEDRAG_Listview)
End Sub

Private Sub picDropTarget_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(OLEDRAG_Listview) Then Effect = vbDropEffectCopy Else Effect = vbDropEffectNone
End Sub

Private Sub picDropTarget_Paint()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Print "Drop Here" on the picturebox used as the drop target.
'             A textbox is not used for this purpose because the drop cursor
'             caused some redrawing problems.
'---------------------------------------------------------------------------------------
    Dim lbThemed As Boolean
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    
    bInHere = True
    lbThemed = IsAppThemed() And lvw.Themeable
    
    Dim liWidth As Long
    Dim liHeight As Long
    Const PicText As String = "Drop Here"
    
    With picDropTarget
        If .Appearance <> (lbThemed + OneL) Then .Appearance = lbThemed + OneL
        If .BorderStyle <> .Appearance Then .BorderStyle = .Appearance
        If lbThemed Then
            picDropTarget.Line (0, 0)-(.Width - .ScaleX(1, vbPixels, .ScaleMode), .Height - .ScaleY(1, vbPixels, .ScaleMode)), RGB(165, 172, 178), B
        End If
        .CurrentX = (.Width \ TwoL - .TextWidth(PicText) \ TwoL) - .ScaleX(.Appearance * 2, vbPixels, .ScaleMode)
        .CurrentY = (.Height \ TwoL - .TextHeight(PicText) \ TwoL) - .ScaleY(.Appearance * 2, vbPixels, .ScaleMode)
    End With
    picDropTarget.Print PicText
    bInHere = False
End Sub

Private Sub txt_Change(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the appropriate property of the listview.
'---------------------------------------------------------------------------------------
    Select Case Index
    Case txtBackPictureX: lvw.BackPictureXOffset = Val(txt(Index).Text)
    Case txtBackPictureY: lvw.BackPictureYOffset = Val(txt(Index).Text)
    Case txtIconSpaceX: lvw.IconSpaceX = Val(txt(Index).Text)
    Case txtIconSpaceY: lvw.IconSpaceY = Val(txt(Index).Text)
    End Select
End Sub

Private Sub txtItem_Change(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the appropriate property of the listview item that is in focus.
'---------------------------------------------------------------------------------------
    If LenB(txtItem(txtItemText).Tag) Then
        With lvw.ListItems.Item(CLng(txtItem(txtItemText).Tag))
            Select Case Index
            Case txtItemText:       .Text = txtItem(Index).Text
            Case txtItemToolTip:    .ToolTipText = txtItem(Index).Text
            Case txtItemIcon:       .IconIndex = Val(txtItem(Index).Text)
            End Select
        End With
    End If
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Private Sub pFill(ByVal iItems As Long)
    Dim liCount As Long
    Randomize
    liCount = iItems
    
    Dim liTileItems(0 To 2) As Long
    liTileItems(0) = 1
    liTileItems(1) = 2
    liTileItems(2) = 3
    
    Dim liIconCount As Long
    liIconCount = lvw.ImageList(lvwImageLargeIcon).IconCount - OneL
    
    With lvw
        udIcon.Max = liIconCount
        .Redraw = False
        With .ListItems
            .Clear
            .InitStorage liCount
            For liCount = 1 To liCount
                With .Add("This is a key " & liCount, "This is an item " & liCount, , Rnd * liIconCount, , , , "Group " & ((liCount Mod 5) + 1), Array(Rnd * 50000, "Sub Item" & liCount, Now + Rnd * 10000))
                    If liCount Mod 2 Then .SubItem(3).IconIndex = 4
                    .SetTileViewItems liTileItems
                End With
            Next
        End With
        .Redraw = True
    End With
End Sub

Public Sub StressTest()
    
    Const ItemCount As Long = 100
    
    pFill ItemCount
  
    Dim i As Long
    For i = ItemCount To 1 Step NegOneL
        lvw.ListItems.Remove i
        lvw.ListItems.Add "This is a key " & i, "This is an item " & i, , Rnd * udIcon.Max, i, , IIf(i < ItemCount, i, pGetMissing), "Group " & ((i Mod 5) + 1), Array(Rnd * 50000, "Sub Item" & i, Now + Rnd * 10000)
    Next
    
    For i = 1 To ItemCount
        Debug.Assert lvw.ListItems(i).Key = "This is a key " & i
        Debug.Assert lvw.ListItems(i).ItemData = i
        Debug.Assert lvw.FindItemData(i).Index = i
    Next
    
End Sub

Private Function pGetMissing(Optional v)
    pGetMissing = v
End Function
