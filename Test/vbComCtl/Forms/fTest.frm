VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#111.0#0"; "vbComCtl.ocx"
Begin VB.Form fTest 
   Caption         =   "Test"
   ClientHeight    =   6180
   ClientLeft      =   390
   ClientTop       =   510
   ClientWidth     =   8355
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin vbComCtl.ucTabStrip tabstrip 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   4020
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   2990
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      MultiLine       =   -1  'True
      RightJustify    =   -1  'True
   End
   Begin vbComCtl.ucToolbar tbar 
      Height          =   495
      Index           =   2
      Left            =   120
      Top             =   2580
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
   End
   Begin vbComCtl.ucStatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      Top             =   5835
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   529
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
   End
   Begin vbComCtl.ucRebar rbar 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   1508
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
         Source          =   2
      EndProperty
   End
   Begin vbComCtl.ucToolbar tbar 
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      MenuStyle       =   -1  'True
      Themeable       =   0   'False
   End
   Begin vbComCtl.ucComboBoxEx cmb 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   3420
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   423
      BeginProperty Fnt {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      ExtUI           =   -1  'True
   End
   Begin vbComCtl.ucToolbar tbar 
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   1860
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      ImageSource     =   0
   End
   Begin vbComCtl.ucPopupMenus pop 
      Left            =   5820
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
   End
   Begin vbComCtl.ucComDlg dlg 
      Left            =   6360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'fTest.frm           3/28/05
'
'            PURPOSE:
'               IE clone gui with rebar, toolbar, combobox and statusbar.
'               Tabstrip to allow navigation through tests of remaining controls.
'
'---------------------------------------------------------------------------------------

Option Explicit

'The controls provided by vbComCtl that we will test on this form.
Private Enum eControl
    ctlAnimation
    ctlComboBox
    ctlComDlg
    ctlDateTimePicker
    ctlFrame
    ctlHotKey
    ctlListView
    ctlMaskedEdit
    ctlMonthCalendar
    ctlPopupMenus
    ctlProgressBar
    ctlRebar
    ctlRichEdit
    ctlScrollBox
    ctlStatusBar
    ctlTabstrip
    ctlToolbar
    ctlTrackbar
    ctlTreeview
    ctlUpDown
End Enum

'Root and sub menu items
Private Enum eMenu
    mnuBarWindow = 1
    mnuBarRebar
    mnuBarMenu
    mnuBarTabstrip
    mnuBarHelp
    
    mnuWindowStressTest
    mnuWindowFont
    mnuWindowThemeable
    mnuWindowResourceUsage
    mnuWindowNew
    mnuWindowClose
    
    mnuRebarTop
    mnuRebarBottom
    mnuRebarLeft
    mnuRebarRight
    mnuRebarLocked
    
    mnuPopupOfficeXP
    mnuPopupBackgroundBitmap
    mnuPopupImageProcessBitmap
    mnuPopupSidebar
    mnuPopupButtonHighlight
    mnuPopupGradientHighlight
    mnuPopupTitleHeaders
    mnuPopupShowInfrequent
    mnuPopupColors
    
    mnuTabstripTabs
    mnuTabstripButtons
    mnuTabstripFlatButtons
    mnuTabstripButtonSeparators
    mnuTabstripMultiline
    
    mnuHelpContents
    mnuHelpSearch
                       
    mnuHelpAnimation = 1000         'help context ids
    mnuHelpComboBox = 1010
    mnuHelpComDlg = 1020
    mnuHelpDateTimePicker = 1030
    mnuHelpFrame = 1040
    mnuHelpHotKey = 1050
    mnuHelpListView = 1060
    mnuHelpMaskedEdit = 1070
    mnuHelpMonthCalendar = 1080
    mnuHelpPopupMenus = 1090
    mnuHelpProgressBar = 1100
    mnuHelpRebar = 1110
    mnuHelpRichEdit = 1120
    mnuHelpScrollBox = 1130
    mnuHelpStatusBar = 1140
    mnuHelpTabstrip = 1150
    mnuHelpToolbar = 1160
    mnuHelpTrackbar = 1170
    mnuHelpTreeview = 1180
    mnuHelpUpDown = 1190
End Enum

'Indexes in the tbar control array.
Private Enum eTbar
    tbarMenu
    tbarButtons
    tbarCorner
End Enum

Private WithEvents moTestControl    As VBControlExtender    'The current control on the tabstrip.
Attribute moTestControl.VB_VarHelpID = -1
Private mbSideBar                   As Boolean              'Flag to indicate whether the menus display a picture along the left side.
Private mbLoaded                    As Boolean              'Flag to indicate whether the form is loaded.
Private miStressTesting             As Long                 'Flag to indicate whether we are in the process of stress testing.

Private Sub Form_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the form.
'---------------------------------------------------------------------------------------
    InitComCtl  'initialize the common controls.  This ensures correct
                'operation when linked to cc version 6.
    vbComCtl.ShowAllUIStates hWnd 'ensure that we show mnemonics and the focus rectangle
End Sub

Private Sub Form_Load()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Add some tabstrip tabs, toolbar buttons, comboboxex items, rebar bands,
'             and status bar panels.
'---------------------------------------------------------------------------------------

    Randomize
    
    pSyncBackColor
    
    dlg.HelpFile = App.Path & "\..\vbComCtl.chm" 'the help file that will be displayed
    
    Dim i As Long
    
    With cmb
        Set .ImageList = gImageListSmall    'set the combo imagelist
        For i = 1 To 100                    'add 100 test items
            .AddItem "Test ComboBoxEx Item " & i, RandIcon(gImageListSmall), , , i Mod 6
        Next
    End With
    
    With tabstrip
        Set .ImageList = gImageListHelp     'set the tabstrip imagelist
        With .Tabs                          'add the test tabs
            For i = ZeroL To pTestControlCount - OneL
                If pTestControlTab(i) Then .Add pTestControlName(i), i + TwoL
            Next
        End With
        .SetSelectedTab OneL
    End With
    
    With tbar
        With .Item(tbarMenu).Buttons
            For i = 1 To 5                  'add five items to the menu bar
                .Add , Choose(i, "&Window", "&Rebar", "&Menu", "&Tabstrip", "&Help")
            Next
        End With
        
        With .Item(tbarButtons).Buttons
            For i = 1 To 10                 'add 10 test buttons
                .Add , "Button " & i, Switch(i = 1, tbarButtonNormal, _
                                             i = 2, tbarButtonWholeDropDown, _
                                             i = 3, tbarButtonDropDown, _
                                    i = 4 Or i = 8, tbarButtonSeparator, _
                                             i < 8, tbarButtonCheckGroup, _
                                              True, tbarButtonCheck), i
            Next
        End With
        
        With .Item(tbarCorner)
            Set .ImageList = gImageListHelp
            .Buttons.Add , , , ZeroL, , , mnuHelpContents
        End With
        
    End With
    
    With rbar.Bands                         'add some rebar bands
        .Add tbar(tbarMenu), , , True
        .Add tbar(tbarCorner), , , , , False, , True
        .Add tbar(tbarButtons), , , True, True
        .Add cmb, , "Address", , True
    End With
    
    With sbar
        Set .ImageList = gImageListSmall
        With .Panels                        'add some test statusbar panels
            .Add "Test Panel", , "Test Panel Tooltip", sbarStandard, , RandIcon(gImageListSmall), ScaleX(100, vbPixels, ScaleMode)
            .Add "Test Spring Panel", , "Test Spring Panel Tooltip", sbarStandard, , RandIcon(gImageListSmall), , , True
            .Add "Test Autofit Panel", , "Test Autofit Panel Tooltip", sbarStandard, , RandIcon(gImageListSmall), , , , True
            .Add , , , sbarCaps, , , , , , True
            .Add , , , sbarIns, , , , , , True
            .Add , , , sbarDateTime, , RandIcon(gImageListSmall), , , , True
        End With
    End With
    
    mbLoaded = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Disallow closing of the form if we are in the middle of a stress test.
'---------------------------------------------------------------------------------------
    If UnloadMode = vbFormControlMenu And miStressTesting > ZeroL Then Cancel = OneL
End Sub

Private Sub Form_Resize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Move the constituent controls to their proper position.
'---------------------------------------------------------------------------------------
    pResize
    sbar.SizeGrip = CBool(Me.WindowState <> vbMaximized)
End Sub

Private Sub moTestControl_ObjectEvent(Info As EventInfo)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Handle events from the usercontrol that is currently on the tabstrip.
'---------------------------------------------------------------------------------------
    Select Case Info.Name
    Case "MenuItemHighlight"
        pop_ItemHighlight Info.EventParameters(0).Value
    Case "GetPopupSettings"
        Dim loPop As ucPopupMenus:  Set loPop = Info.EventParameters(0).Value
        Dim loMenu As cPopupMenu:   Set loMenu = Info.EventParameters(1).Value
        
        With pop
            If loPop.BackgroundPictureExists Xor .BackgroundPictureExists Then
                If .BackgroundPictureExists _
                    Then loPop.SetBackgroundPicture mResources.PopupBackPicture _
                    Else loPop.SetBackgroundPicture Nothing
            End If
            
            loPop.ButtonHighlight = .ButtonHighlight
            loPop.ColorActiveBack = .ColorActiveBack
            loPop.ColorActiveFore = .ColorActiveFore
            loPop.ColorInactiveBack = .ColorInactiveBack
            loPop.ColorInactiveFore = .ColorInactiveFore
            loPop.GradientHighlight = .GradientHighlight
            loPop.ImageProcessBitmap = .ImageProcessBitmap
            loPop.OfficeXPStyle = .OfficeXPStyle
            loPop.ShowInfrequent = .ShowInfrequent
            loPop.ShowInfrequentHoverDelay = .ShowInfrequentHoverDelay
            loPop.TitleSeparators = .TitleSeparators
            
            If mbSideBar Then loMenu.SetSidebar mResources.PopupSideBar
        End With
        
    End Select
End Sub

Private Sub rbar_Resize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : The rebar's size has changed, possibly due to the user dragging one of the bands.
'             Move the constituent controls to their proper position on the form.
'---------------------------------------------------------------------------------------
    If mbLoaded Then pResize
End Sub

Private Sub sbar_PanelClick(ByVal oPanel As vbComCtl.cPanel, ByVal iButton As vbComCtl.evbComCtlMouseButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Notify the user that we received a click event on the statusbar.
'---------------------------------------------------------------------------------------
    MsgBox "You clicked " & oPanel.Text
End Sub

Private Sub tabstrip_Click(ByVal oTab As vbComCtl.cTab)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Load and display the selected control.
'---------------------------------------------------------------------------------------
    If Not moTestControl Is Nothing Then Controls.Remove moTestControl
    Set moTestControl = Controls.Add(App.EXEName & "." & oTab.Text & "Test", oTab.Text & "TestControl")
    pUpdateTestControl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Allow the toolbar menu to interpret Alt and F10 keypresses
'             to activate the menu bar.
'             Allow toolbars to raise events for keyboard mnemonics.
'---------------------------------------------------------------------------------------
    vbComCtl.TbarContainerKeyDown KeyCode, Shift, tbar
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Correctly terminate the help window if it was shown.
'
'             In the IDE, this does not work ideally.  It's kind of a catch-22.
'                   If you DO call dlgHelpCloseAll:
'                       If the help window was previously shown and had been closed manually
'                       then you may crash after returning to design mode.  To prevent this,
'                       don't close the help file if you display it unless you are running
'                       the compiled version.
'
'                       This is a very quirky error.  It has produced both 'unknown software
'                       exception' and 'access violation'.  If after returning to design mode
'                       as described above you keep VB as the active task, the crash
'                       does not immediately appear.  It is only upon starting up another task,
'                       expecially IE or Explorer, that the IDE crashes.
'
'                   If you DO NOT call dlgHelpCloseAll:
'                       If the help window is still visible when returning to design mode,
'                       the IDE will almost certainly crash.
'---------------------------------------------------------------------------------------
    mbLoaded = False
    
    If Forms.Count = 2 And IsLoaded(fSystemResources) _
        Then Unload fSystemResources
    
    If Forms.Count = 1 Then dlg.ShowHelp dlgHelpCloseAll
    
    If Not moTestControl Is Nothing Then
        Controls.Remove moTestControl
        Set moTestControl = Nothing
    End If
    
    'Clearing the rebar bands is necessary to avoid a possible memory leak.
    rbar.Bands.Clear
End Sub

Private Sub pop_Click(ByVal oItem As vbComCtl.cPopupMenuItem)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Perform the action associated with each popup menu item.
'---------------------------------------------------------------------------------------
    tbar(tbarMenu).HidePopup                                    'If a menu is being pressed from a
    tbar(tbarButtons).HidePopup                                 'a chevron toolbar, hide it.
    
    CallMeBack Me, "OnMenuItemClick", VbMethod, oItem           'Return execution so that the pressed menu button or pressed
                                                                'chevron are unpressed while we process the button click.
End Sub

Private Sub pop_ItemHighlight(ByVal oItem As vbComCtl.cPopupMenuItem)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Display the item help string in the statusbar.
'---------------------------------------------------------------------------------------
    If oItem Is Nothing _
        Then sbar.SimpleText = vbNullString _
        Else sbar.SimpleText = oItem.HelpString
    sbar.Simple = LenB(sbar.SimpleText)
End Sub

Private Sub rbar_ChevronPushed(ByVal oBand As vbComCtl.cBand, ByVal fLeft As Single, ByVal fTop As Single, ByVal fWidth As Single, ByVal fHeight As Single)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the toolbar chevron.
'---------------------------------------------------------------------------------------
    Dim loToolbar As ucToolbar: Set loToolbar = oBand.Child
    loToolbar.ShowPopup fLeft + fWidth, fTop, fLeft, fTop, fWidth, fHeight, rbar.Align
End Sub

Private Sub tabstrip_RightClick(ByVal oTab As vbComCtl.cTab)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the menu that will allow the user to change the display of the tabs.
'---------------------------------------------------------------------------------------
    pMenu(mnuBarTabstrip).ShowAtCursor mnuRightButton
End Sub

Private Sub tbar_ButtonClick(Index As Integer, ByVal oButton As vbComCtl.cButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Hide the toolbar in case the button was clicked from a chevron.
'             Show the help dialog or a message indicating the button that was clicked.
'---------------------------------------------------------------------------------------
    tbar(Index).HidePopup                                       'Hide the chevron toolbar if it is visible.
    CallMeBack Me, "OnToolbarButtonClick", VbMethod, oButton    'Return execution so that any pressed chevron is
                                                                'unpressed while we process the button click.
End Sub

Private Sub tbar_ButtonDropDown(Index As Integer, ByVal oButton As vbComCtl.cButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show a dropdown menu next to a button.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowMenuAtButton pMenu(oButton.Index), oButton
    If Index = tbarButtons And pop.ChevronWasClicked Then pop.ShowInfrequent = False
End Sub

Private Sub tbar_ExitMenuTrack(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Reset the visibility of the infrequent items if they were shown.
'---------------------------------------------------------------------------------------
    If pop.ChevronWasClicked Then pop.ShowInfrequent = False
End Sub

Private Sub tbar_RightButtonUp(Index As Integer, ByVal x As Single, ByVal y As Single)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the rebar popup menu if the user right clicks a toolbar.
'---------------------------------------------------------------------------------------
    pMenu(mnuBarRebar).ShowAtCursor mnuRightButton
End Sub

Private Property Get pMenu(ByVal iIndex As Long) As vbComCtl.cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Build and return the menu for the given menu bar button.
'---------------------------------------------------------------------------------------
    Dim i As Long
    
    Set pMenu = pop.NewMenu()
    
    Set pop.ImageList = IIf(iIndex <> mnuBarHelp, gImageListSmall, gImageListHelp)
    With pMenu
        Select Case iIndex
        Case mnuBarWindow
            .Add "&Stress Test ...", "Continually load and unload controls and other memory hungry resources to check for leaks.", , mnuWindowStressTest, , , , mnuWindowStressTest
            .Add "&Font ...", "Change the font for the controls on this window.", , mnuWindowFont, , , , mnuWindowFont
            If vbComCtl.IsAppThemed Then
                .Add "&Themeable", "Toggle whether the default theme is applied to this window.", , , mnuChecked * -rbar.Themeable, , , mnuWindowThemeable
            End If
            .Add "&Resource Usage ...", "Show OS resource counters.", , mnuWindowResourceUsage, , , , mnuWindowResourceUsage
            .Add , , , , mnuSeparator
            .Add "&New ...", "Open a new window.", , mnuWindowNew, , , , mnuWindowNew
            .Add "&Close", "Close this window.", , mnuWindowClose, , , , mnuWindowClose
        Case mnuBarRebar
            .Add "&Left", "Align the rebar at the left of this window.", , , mnuChecked * -(rbar.Align = vbccAlignLeft), , , mnuRebarLeft
            .Add "&Right", "Align the rebar at the right of this window.", , , mnuChecked * -(rbar.Align = vbccAlignRight), , , mnuRebarRight
            .Add "&Top", "Align the rebar at the top of this window.", , , mnuChecked * -(rbar.Align = vbccAlignTop), , , mnuRebarTop
            .Add "&Bottom", "Align the rebar at the bottom of this window.", , , mnuChecked * -(rbar.Align = vbccAlignBottom), , , mnuRebarBottom
            .Add , , , , mnuSeparator
            .Add "Loc&ked", "Lock or unlock the rebar bands.", , , mnuChecked * (rbar.Bands.Item(1).Gripper + OneL), , , mnuRebarLocked
        Case mnuBarMenu
            .Add "Drawing Options", , , , mnuSeparator
            If SystemColorDepth > imlColor8 Then
                .Add "&Office XP Style", "Paint the menus similar to the office xp theme.", , , mnuChecked * -pop.OfficeXPStyle, , , mnuPopupOfficeXP
                .Add "&Gradient Highlight", "Show highlight using a gradient from inactive to active forecolors.", , , mnuChecked * -pop.GradientHighlight, , , mnuPopupGradientHighlight
                .Add "&Title Separators", "Draw separators to be distinct from menu items.", , , mnuChecked * -pop.TitleSeparators, , , mnuPopupTitleHeaders
            End If
            .Add "B&utton Highlight", "Show highlight using a raised button.", , , mnuChecked * -pop.ButtonHighlight, , , mnuPopupButtonHighlight
            If SystemColorDepth > imlColor8 Then
                .Add "Picture Options", , , , mnuSeparator
                .Add "&Background Bitmap", "Use a bitmap for the background insteam of the In/ActiveBackColor properties.", , , mnuChecked * -pop.BackgroundPictureExists, , , mnuPopupBackgroundBitmap
                .Add "&Image Process Bitmap", "Lighten the background bitmap for infrequent items and highlight.", , , mnuChecked * -pop.ImageProcessBitmap, , , mnuPopupImageProcessBitmap
                .Add "Si&de Bar", "Show a picture along the left edge of the menus.", , , mnuChecked * -mbSideBar, , , mnuPopupSidebar
            End If
            .Add "Other", , , , mnuSeparator
            .Add "&Show Infrequent", "Show infrequently used items.", , , mnuChecked * -pop.ShowInfrequent, , , mnuPopupShowInfrequent
            .Add "&Colors", "Change the different colors of the menu.", , RandIcon(gImageListSmall), , , , mnuPopupColors
            With .Add("Test Infrequent").SubMenu
                .ShowCheckAndIcon = True
                For i = 1 To 10
                    .Add "Test Item " & i, "Test Infrequent Item " & i, , RandIcon(gImageListSmall), (-CBool((Rnd * 10) > 5) * mnuInfrequent) Or (-CBool(Rnd > 0.5) * mnuChecked)
                Next
            End With
        Case mnuBarTabstrip
            If Not vbComCtl.IsAppThemed Or Not rbar.Themeable Then
                .Add "&Tabs", "Display the default tabs.", , , mnuRadioChecked * (tabstrip.Buttons + OneL), , , mnuTabstripTabs
                .Add "&Buttons", "Display 3D buttons instead of tabs.", , , mnuRadioChecked * -(tabstrip.Buttons And Not tabstrip.FlatButtons), , , mnuTabstripButtons
                .Add "&Flat Buttons", "Display flat buttons instead of tabs.", , , mnuRadioChecked * -(tabstrip.Buttons And tabstrip.FlatButtons), , , mnuTabstripFlatButtons
                .Add , , , , mnuSeparator
                .Add "Flat Button &Separators", "Display separators between flat buttons.", , , mnuChecked * -tabstrip.FlatSeparators, , , mnuTabstripButtonSeparators
            End If
            .Add "Multiline", "Wrap tabs or buttons instead of scrolling.", , , mnuChecked * -tabstrip.MultiLine, , , mnuTabstripMultiline
        Case mnuBarHelp
            .Add "Contents", "Display the help file contents.", , 0, , , , mnuHelpContents
            .Add "Search", "Search the help file.", , 1, , , , mnuHelpSearch
            For i = ZeroL To pTestControlCount - OneL
                .Add pTestControlName(i), "Display help on the " & pTestControlName(i) & " control.", , i + TwoL, , , , mnuHelpAnimation + (i * 10)
            Next
        End Select
        If mbSideBar Then .SetSidebar mResources.PopupSideBar
    End With
End Property

Private Sub pAlignRebar(ByVal iAlign As AlignConstants)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Align the rebar to the specified edge and ensure that the combobox is
'             only visible when the rebar is in horizontal alignment.
'---------------------------------------------------------------------------------------
    
    rbar.Redraw = False
   
    rbar.SetAlignment iAlign
    rbar.Bands.Item(cmb).Visible = CBool(iAlign < vbAlignLeft)
    
    rbar.Redraw = True
    
End Sub

Private Sub pSetTabstrip(ByVal bButtons As Boolean, ByVal bFlat As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the Buttons, FlatButtons and HotTrack properties of the tabstrip.
'---------------------------------------------------------------------------------------
    With tabstrip
        .Buttons = bButtons
        .FlatButtons = bFlat
        .HotTrack = bFlat
    End With
    pResize
End Sub

Private Sub pResize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Move the non-aligned controls to their position on the form.
'---------------------------------------------------------------------------------------
    Static bInHere As Boolean
    If bInHere Or Not mbLoaded Then Exit Sub
    bInHere = True
    
    On Error Resume Next
    tabstrip.Align = vbccAlignNone
    Select Case rbar.Align
    Case vbccAlignLeft
        tabstrip.Move rbar.Width, ZeroL, ScaleWidth - rbar.Width, ScaleHeight - sbar.Height
    Case vbccAlignRight
        tabstrip.Move ZeroL, ZeroL, ScaleWidth - rbar.Width, ScaleHeight - sbar.Height
    Case vbccAlignTop
        tabstrip.Move ZeroL, rbar.Height, ScaleWidth, ScaleHeight - sbar.Height - rbar.Height
    Case vbccAlignBottom
        tabstrip.Move ZeroL, ZeroL, ScaleWidth, ScaleHeight - sbar.Height - rbar.Height
    End Select
    
    If Not moTestControl Is Nothing Then tabstrip.MoveToClient moTestControl
    
    'This to ensure that the rebar stays above the status bar
    'while the bands are being dragged around.
    If rbar.Align = vbAlignBottom Then pAlignRebar rbar.Align
    bInHere = False
    On Error GoTo 0
End Sub

Private Sub pShowHelp(ByVal iCmdShow As vbComCtl.eHelpDialog, Optional ByVal vTopicNameOrID As Variant)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Let the user know that showing the help dialog from the IDE is at their own risk!
'---------------------------------------------------------------------------------------
    Static bConfirmed As Boolean
    
    If Not bConfirmed Then
        bConfirmed = Not InIDE()
        If Not bConfirmed Then bConfirmed = CBool(MsgBox("Using the HTML Help Api from the IDE can sometimes cause an access violation seconds or minutes after you return to design mode.  Are you sure you want to do this?" & vbNewLine & vbNewLine & vbNewLine & "If you choose yes, ensuring that the HTML Help window is still visible when closing the last test form to return to design mode *seems* to solve the problem.", vbYesNo Or vbDefaultButton2) = vbYes)
    End If
    
    If bConfirmed Then dlg.ShowHelp iCmdShow, vTopicNameOrID
    
End Sub

Private Sub pSyncBackColor()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Change the back color to match either the classic windows theme or white.
'             for the default 'Luna' theme. Constituent usercontrols catch the ambient change
'             of the backcolor property.
'---------------------------------------------------------------------------------------
    Me.BackColor = IIf(vbComCtl.IsAppThemed And rbar.Themeable, vbWhite, vbButtonFace)
End Sub

Private Property Get pTestControlName(ByVal iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the name of a type of control that is being tested on this form.
'---------------------------------------------------------------------------------------
    Select Case iIndex
    Case ctlAnimation:      pTestControlName = "ucAnimation"
    Case ctlComboBox:       pTestControlName = "ucComboBoxEx"
    Case ctlComDlg:         pTestControlName = "ucComDlg"
    Case ctlDateTimePicker: pTestControlName = "ucDateTimePicker"
    Case ctlFrame:          pTestControlName = "ucFrame"
    Case ctlHotKey:         pTestControlName = "ucHotKey"
    Case ctlListView:       pTestControlName = "ucListView"
    Case ctlMaskedEdit:     pTestControlName = "ucMaskedEdit"
    Case ctlMonthCalendar:  pTestControlName = "ucMonthCalendar"
    Case ctlPopupMenus:     pTestControlName = "ucPopupMenus"
    Case ctlProgressBar:    pTestControlName = "ucProgressBar"
    Case ctlRebar:          pTestControlName = "ucRebar"
    Case ctlRichEdit:       pTestControlName = "ucRichEdit"
    Case ctlScrollBox:      pTestControlName = "ucScrollBox"
    Case ctlStatusBar:      pTestControlName = "ucStatusBar"
    Case ctlTabstrip:       pTestControlName = "ucTabstrip"
    Case ctlToolbar:        pTestControlName = "ucToolbar"
    Case ctlTrackbar:       pTestControlName = "ucTrackbar"
    Case ctlTreeview:       pTestControlName = "ucTreeview"
    Case ctlUpDown:         pTestControlName = "ucUpDown"
    End Select
End Property

Private Property Get pTestControlCount() As Long
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the number of controls being tested by this form.
'---------------------------------------------------------------------------------------
    pTestControlCount = ctlUpDown + OneL
End Property

Private Property Get pTestControlTab(ByVal iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a value indicating whether the given test control gets its own tab
'             on the tabstrip.
'---------------------------------------------------------------------------------------
    Select Case iIndex
    Case ctlAnimation, ctlComDlg, ctlDateTimePicker, ctlFrame, ctlHotKey, ctlListView, ctlMaskedEdit, _
         ctlMonthCalendar, ctlProgressBar, ctlRichEdit, ctlTrackbar, ctlTreeview, ctlUpDown
        pTestControlTab = True
    End Select
End Property

Private Sub pUpdateTestControl()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Ensure that the test control is displayed in the appropriate fashion to
'             match the rest of the form.
'---------------------------------------------------------------------------------------
    If Not moTestControl Is Nothing Then
        With moTestControl
            .Visible = False
            pSyncBackColor
            .object.Themeable = rbar.Themeable
            .ZOrder
            tabstrip.MoveToClient moTestControl
            .Visible = True
        End With
    End If
End Sub

Private Sub pStressTest()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Continually load, show, hide, and unload fTest forms.
'---------------------------------------------------------------------------------------
    Dim loTest As fTest
    
    If MsgBox("A copy of this form will now be loaded, shown and unloaded continually. Each test control will be loaded and unloaded for each form.  Each test control will possibly load/unload additional resources, like listview items or treeview nodes. Press Escape to stop the madness." & vbNewLine & vbNewLine & vbNewLine & "Ready?", vbYesNo Or vbDefaultButton2) = vbYes Then
        On Error GoTo handler
        miStressTesting = miStressTesting + OneL
        Do Until KeyIsDown(VK_ESCAPE)
            Set loTest = New fTest
            Load loTest
            loTest.Visible = True
            loTest.fStressTest
            loTest.Visible = False
            Unload loTest
            Set loTest = Nothing
        Loop
handler:
        Debug.Assert Err.Number = False
        If Err.Number Then MsgBox "Error: " & Err.Number & vbNewLine & Err.Description
        On Error GoTo 0
        miStressTesting = miStressTesting - OneL
    End If
End Sub

Friend Sub fStressTest()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Load each test control and ask each to load/unload its own resources.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    miStressTesting = miStressTesting + OneL
    Dim loTab As cTab
    For Each loTab In tabstrip.Tabs
        tabstrip.SetSelectedTab loTab
        
        Select Case True
            Case TypeOf moTestControl Is ucAnimationTest, _
                 TypeOf moTestControl Is ucListViewTest, _
                 TypeOf moTestControl Is ucRichEditTest, _
                 TypeOf moTestControl Is ucTreeViewTest
                 moTestControl.StressTest
        End Select
        
        
        If KeyIsDown(VK_ESCAPE) Then Exit For
        DoEvents
        If KeyIsDown(VK_ESCAPE) Or Not mbLoaded Then Exit For
    Next
    miStressTesting = miStressTesting - OneL
    On Error GoTo 0
End Sub

Public Sub OnMenuItemClick(ByVal oItem As cPopupMenuItem)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Process menu item clicks.
'---------------------------------------------------------------------------------------
    Select Case oItem.ItemData
    Case mnuWindowNew:                  Dim oWindow As New fTest: oWindow.Show
    Case mnuWindowClose:                Unload Me
    Case mnuPopupOfficeXP:              pop.OfficeXPStyle = Not oItem.Checked
    Case mnuPopupImageProcessBitmap:    pop.ImageProcessBitmap = Not oItem.Checked
    Case mnuPopupSidebar:               mbSideBar = Not oItem.Checked
    Case mnuPopupButtonHighlight:       pop.ButtonHighlight = Not oItem.Checked
    Case mnuPopupGradientHighlight:     pop.GradientHighlight = Not oItem.Checked
    Case mnuPopupTitleHeaders:          pop.TitleSeparators = Not oItem.Checked
    Case mnuPopupShowInfrequent:        pop.ShowInfrequent = Not oItem.Checked
    Case mnuPopupColors:                fEditColors.EditColors Me, pop, "ColorActiveFore", "ColorActiveBack", "ColorInactiveFore", "ColorInactiveBack"
    Case mnuTabstripTabs:               pSetTabstrip False, False
    Case mnuTabstripButtons:            pSetTabstrip True, False
    Case mnuTabstripFlatButtons:        pSetTabstrip True, True
    Case mnuTabstripButtonSeparators:   tabstrip.FlatSeparators = Not oItem.Checked
    Case mnuTabstripMultiline:          tabstrip.MultiLine = Not oItem.Checked: pResize
    Case mnuWindowStressTest:           pStressTest
    Case mnuWindowResourceUsage:        fSystemResources.Show

    Case mnuHelpContents:               pShowHelp dlgHelpContents
    Case mnuHelpSearch:                 pShowHelp dlgHelpSearch
    Case Is > mnuHelpContents:          pShowHelp dlgHelpContext, oItem.ItemData
    
    Case mnuRebarLeft, mnuRebarRight, _
         mnuRebarTop, mnuRebarBottom
                                        pAlignRebar oItem.ItemData - mnuRebarTop + OneL
    
    Case mnuWindowFont:                 If pop.Font.Browse(hWnd) Then Set Me.Font = pop.Font.FontData(fntDataTypeStdFont): pResize
    
    Case mnuRebarLocked
        Dim loBand As cBand
        For Each loBand In rbar.Bands
            loBand.Gripper = Not loBand.Gripper
        Next
    
    Case mnuPopupBackgroundBitmap
        If pop.BackgroundPictureExists _
            Then pop.SetBackgroundPicture Nothing _
            Else pop.SetBackgroundPicture mResources.PopupBackPicture
    
    Case mnuWindowThemeable
        vbComCtl.ThemeControls Controls, Not oItem.Checked
        tbar(tbarMenu).Themeable = False
        If Not oItem.Checked Then pSetTabstrip False, False
        pUpdateTestControl
        pResize
    
    Case Else:                          MsgBox "You clicked: " & oItem.Caption
    End Select
End Sub

Public Sub OnToolbarButtonClick(ByVal oButton As cButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Process toolbar button clicks.
'---------------------------------------------------------------------------------------
    If oButton.ItemData = mnuHelpContents _
        Then pShowHelp dlgHelpContents _
        Else MsgBox "You clicked " & oButton.Text
End Sub
