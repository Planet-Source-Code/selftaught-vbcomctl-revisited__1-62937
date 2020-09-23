VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#109.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucRichEditTest 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CheckBox chkModifyProtected 
      Caption         =   "Allow Modify Protected"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin vbComCtl.ucToolbar tbar 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   1296
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Style           =   2560
      TextRows        =   0
   End
   Begin vbComCtl.ucRichEdit rtxt 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3413
      BooleanProps    =   1613762756
      BackColor       =   -2147483643
      MaxLen          =   32767
      LMargin         =   7
      RMargin         =   7
   End
   Begin vbComCtl.ucPopupMenus pop 
      Left            =   2520
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Flags           =   56
   End
   Begin vbComCtl.ucComDlg dlg 
      Left            =   3060
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "ucRichEditTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucRichEditTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the rich edit control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eButton    'toolbar buttons
    btnBold
    btnItalic
    btnUnderline
    btnCut
    btnCopy
    btnPaste
    btnLeft
    btnCenter
    btnRight
    btnJustify
    btnLink
    btnProtected
    btnUndo
    btnFontColor
    btnFont
    btnNew
    btnOpen
    btnSave
    btnPrint
    btnProps
    btnFind
End Enum

Private Enum eProp  'rich edit properties displayed in the popup menu
    propAutoHScroll
    propAutoVScroll
    propAutoURLDetect
    propDisableNoScroll
    propEnabled
    propHideSelection
    propModified
    propMultiLine
    propPlainText
    propReadOnly
    propSelectionBar
    propWantTab
    propWordWrap
    propBorderNone
    propBorderSingle
    propBorderThin
    propBorder3DSunken
    propLeftMargin
    propRightMargin
    propMaxLength
    propScrollNone
    propScrollHorz
    propScrollVert
    propScrollBoth
    propRecreate
End Enum

Event MenuItemHighlight(ByVal oItem As cPopupMenuItem)
Event GetPopupSettings(ByVal oPop As ucPopupMenus, ByVal oMenu As cPopupMenu)

Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long   'delete the dc from the print dialog after
                                                                                'printing the contents of the rich edit.

Private WithEvents mfFind As fFind
Attribute mfFind.VB_VarHelpID = -1

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
    Set tbar.ImageList = gImageListRichEdit
    With tbar.Buttons
        .Add , "New", , btnNew, , True
        .Add , "Open", , btnOpen, , True
        .Add , "Save", , btnSave, , True
        .Add , "Print", , btnPrint, , True
        .Add , , tbarButtonSeparator
        .Add , "Bold", tbarButtonCheck, btnBold, , True
        .Add , "Italic", tbarButtonCheck, btnItalic, , True
        .Add , "Underline", tbarButtonCheck, btnUnderline, , True
        .Add , , tbarButtonSeparator
        .Add , "Link", tbarButtonCheck, btnLink, , True
        .Add , "Protected", tbarButtonCheck, btnProtected, , True
        .Add , , tbarButtonSeparator
        .Add , , tbarButtonSeparator, chkModifyProtected.Width \ Screen.TwipsPerPixelX
        .Add , , tbarButtonSeparator
        .Add , "Left", tbarButtonCheckGroup, btnLeft, , True
        .Add , "Center", tbarButtonCheckGroup, btnCenter, , True
        .Add , "Right", tbarButtonCheckGroup, btnRight, , True
        .Add , "Justify", tbarButtonCheckGroup, btnJustify, , True
        .Add , , tbarButtonSeparator
        .Add , "Cut", , btnCut, , True
        .Add , "Copy", , btnCopy, , True
        .Add , "Paste", , btnPaste, , True
        .Add , "Undo", , btnUndo, , True
        .Add , , tbarButtonSeparator
        .Add , "Font", , btnFont, , True
        .Add , "Font Color", , btnFontColor, , True
        .Add , "Find", , btnFind, , True
        .Add , "Properties", tbarButtonWholeDropDown, btnProps, , True
    End With
    rtxt_SelectionChange 0, 0, 0
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
    tbar_Resize
End Sub

Private Sub UserControl_Terminate()
    If Not mfFind Is Nothing Then Unload mfFind
End Sub

Private Sub mfFind_FindText(sText As String, ByVal bWholeWord As Boolean, ByVal bMatchCase As Boolean, bPassedEnd As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Search the rich edit control, starting over at the top if necessary.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    liIndex = rtxt.FindText(sText, bWholeWord, bMatchCase, rtxt.SelEnd)
    
    If liIndex = NegOneL Then
        liIndex = rtxt.FindText(sText, bWholeWord, bMatchCase)
        bPassedEnd = (liIndex > NegOneL)
    End If
    
    If liIndex > NegOneL Then
        rtxt.HideSelection = False
        rtxt.SetSelection liIndex, Len(sText)
    Else
        MsgBox "The specified text was not found."
    End If
End Sub

Private Sub pop_Click(ByVal oItem As vbComCtl.cPopupMenuItem)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Perform the action of the menu item.
'---------------------------------------------------------------------------------------
    Dim lbVal As Boolean: lbVal = Not oItem.Checked
    Dim lsText As String
    
    Select Case oItem.ItemData
    Case propAutoHScroll:       rtxt.AutoHScroll = lbVal
    Case propAutoVScroll:       rtxt.AutoVScroll = lbVal
    Case propAutoURLDetect:     rtxt.AutoURLDetect = lbVal
    Case propDisableNoScroll:   rtxt.DisableNoScroll = lbVal
    Case propEnabled:           rtxt.Enabled = lbVal
    Case propHideSelection:     rtxt.HideSelection = lbVal
    Case propModified:          rtxt.Modified = lbVal
    Case propMultiLine:         rtxt.MultiLine = lbVal
    Case propPlainText:         rtxt.PlainText = lbVal
    Case propReadOnly:          rtxt.ReadOnly = lbVal
    Case propSelectionBar:      rtxt.SelectionBar = lbVal
    Case propWantTab:           rtxt.WantTab = lbVal
    Case propWordWrap:          rtxt.WordWrap = lbVal
    Case propRecreate:          rtxt.Recreate True
    Case propBorderNone To propBorder3DSunken
                                rtxt.BorderStyle = oItem.ItemData - propBorderNone
    Case propScrollNone To propScrollBoth
                                rtxt.ScrollBars = oItem.ItemData - propScrollNone + OneL
    Case propLeftMargin, propRightMargin
        lsText = InputBox("Enter the new margin in twips:")
        If LenB(lsText) Then
            If oItem.ItemData = propLeftMargin _
                Then rtxt.LeftMargin = Val(lsText) _
                Else rtxt.RightMargin = Val(lsText)
        End If
    Case propMaxLength
        lsText = InputBox("Enter the new maximum length:")
        If LenB(lsText) Then rtxt.MaxLength = Val(lsText)
    
    Case Else:                  Debug.Assert False
    End Select
End Sub

Private Sub pop_ItemHighlight(ByVal oItem As vbComCtl.cPopupMenuItem)
    RaiseEvent MenuItemHighlight(oItem)
End Sub

Private Sub rtxt_ContextMenu(ByVal x As Single, ByVal y As Single)
    pMenu.ShowAtCursor mnuRightButton
End Sub

Private Sub rtxt_KeyDown(iKeyCode As Integer, ByVal iState As vbComCtl.evbComCtlKeyboardState, ByVal bRepeat As Boolean)
    If iKeyCode = vbKeyF And iState = vbccControlMask Then pShowFind
End Sub

Private Sub rtxt_ModifyProtected(bModify As Boolean, ByVal iMin As Long, ByVal iMax As Long)
    bModify = CBool(chkModifyProtected.Value)
End Sub

Private Sub rtxt_SelectionChange(ByVal iMin As Long, ByVal iMax As Long, ByVal iSelType As vbComCtl.eRichEditSelectionType)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the toolbar buttons to match the selection.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    With rtxt.CharFormat(rtfFormatSelection)
        pUpdateButton 6, .Bold, .Consistent
        pUpdateButton 7, .Italic, .Consistent
        pUpdateButton 8, .Underline, .Consistent
        pUpdateButton 10, .Link, .Consistent
        pUpdateButton 11, .Protected, .Consistent
    End With
    With rtxt.ParaFormat
        Dim liAlign As Long:            liAlign = .Alignment
        Dim lbConsistent As Boolean:    lbConsistent = .Consistent
        pUpdateButton 15, liAlign = rtfParaLeft, lbConsistent
        pUpdateButton 16, liAlign = rtfParaCenter, lbConsistent
        pUpdateButton 17, liAlign = rtfParaRight, lbConsistent
        pUpdateButton 18, liAlign = rtfParaJustify, lbConsistent
    End With
    With tbar.Buttons
        .Item(20).Enabled = rtxt.CanCopy
        .Item(21).Enabled = rtxt.CanCopy
        .Item(22).Enabled = rtxt.CanPaste
        .Item(23).Enabled = rtxt.CanUndo
    End With
    On Error GoTo 0
End Sub

Private Sub pUpdateButton(ByVal vButton As Variant, ByVal bVal As Boolean, ByVal bConsistent As Boolean)
    With tbar.Buttons(vButton)
        .Checked = bVal
        .Grayed = Not bConsistent
    End With
End Sub

Private Sub tbar_ButtonClick(ByVal oButton As vbComCtl.cButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Perform the action associated with each button.
'---------------------------------------------------------------------------------------
    Dim lsFile As String
    On Error GoTo handler
    Select Case oButton.IconIndex
    Case btnNew:        rtxt.Text = vbNullString
    Case btnBold:       rtxt.CharFormat(rtfFormatSelection).Bold = oButton.Checked
    Case btnItalic:     rtxt.CharFormat(rtfFormatSelection).Italic = oButton.Checked
    Case btnUnderline:  rtxt.CharFormat(rtfFormatSelection).Underline = oButton.Checked
    Case btnLink:       rtxt.CharFormat(rtfFormatSelection).Link = oButton.Checked
    Case btnProtected:  rtxt.CharFormat(rtfFormatSelection).Protected = oButton.Checked
    Case btnCut:        rtxt.Cut
    Case btnCopy:       rtxt.Copy
    Case btnPaste:      rtxt.Paste
    Case btnLeft:       rtxt.ParaFormat.Alignment = rtfParaLeft
    Case btnCenter:     rtxt.ParaFormat.Alignment = rtfParaCenter
    Case btnRight:      rtxt.ParaFormat.Alignment = rtfParaRight
    Case btnJustify:    rtxt.ParaFormat.Alignment = rtfParaJustify
    Case btnUndo:       rtxt.Undo
    Case btnFontColor:  fEditColors.EditColors Parent, rtxt.CharFormat(rtfFormatSelection), "ColorFore", "ColorBack"
    Case btnOpen:       If dlg.ShowFileOpen(lsFile, dlgFileExplorerStyle Or dlgFileHideReadOnly Or dlgFilePathMustExist, dlg.FileGetFilter("Rich Text Files (*.rtf)", "*.rtf", "Text Files (*.txt)", "*.txt", "All Files (*.*)", "*.*"), , , , , "Open a text file") Then rtxt.LoadFile lsFile, Right$(lsFile, 4) = ".rtf"
    Case btnSave:       If dlg.ShowFileSave(lsFile, dlgFileExplorerStyle Or dlgFileHideReadOnly Or dlgFilePathMustExist, dlg.FileGetFilter("Rich Text Files (*.rtf)", "*.rtf", "Text Files (*.txt)", "*.txt", "All Files (*.*)", "*.*"), , , , , "Save text file") Then rtxt.SaveFile lsFile, Right$(lsFile, 4) = ".rtf"
    Case btnFind:       pShowFind
    Case btnPrint
        'If rtxt.CharacterCount Then
            Dim lhDc As Long
            Dim liRange As ePrintRange
            If dlg.ShowPrint(lhDc, dlgPrintReturnDc Or dlgPrintNoPageNums Or dlgPrintNoSelection, liRange) Then
                Debug.Assert lhDc
                If lhDc Then
                    With rtxt.NewPrintJob(lhDc, "Rich Edit Sample Print Job")
                        '1440 twips = 1 inch margins
                        Do While .DoPrint(1440, 1440, 1440, 1440): Loop
                    End With
                    DeleteDC lhDc
                End If
            End If
        'End If
    Case btnFont
        Dim loFont As cFont
        Set loFont = New cFont
        With rtxt.CharFormat(rtfFormatSelection)
            loFont.FaceName = .FaceName
            loFont.Weight = IIf(.Bold, fntWeightBold, fntWeightNormal)
            loFont.Height = .Height
            loFont.Italic = .Italic
            loFont.Underline = .Underline
            loFont.Strikeout = .Strikeout
            If loFont.Browse(ContainerHwnd) Then
                .FaceName = loFont.FaceName
                .Bold = loFont.Weight > fntWeightNormal
                .Height = loFont.Height
                .Italic = loFont.Italic
                .Underline = loFont.Underline
                .Strikeout = loFont.Strikeout
            End If
        End With
    
    Case Else:              Debug.Assert False
    End Select
handler:
    If Err.Number Then MsgBox "Error: " & Err.Number & vbNewLine & Err.Description
    rtxt_SelectionChange 0, 0, 0
    On Error GoTo 0
End Sub

Private Sub tbar_ButtonDropDown(ByVal oButton As vbComCtl.cButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show a menu next to the dropped button.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowMenuAtButton pMenu, oButton
End Sub

Private Sub tbar_Resize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Resize the richedit to fill the portion of the control not taken by the toolbar.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    rtxt.Move 0, tbar.Height, Width, Height - tbar.Height
    If tbar.Buttons.Count > 13 Then
        With tbar.Buttons(13)
            chkModifyProtected.Move .Left, .Top, .Width, .Height
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub pShowFind()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the find dialog.
'---------------------------------------------------------------------------------------
    If mfFind Is Nothing Then Set mfFind = New fFind
    mfFind.ShowFind Parent, rtxt.TextRange(rtxt.SelStart, rtxt.SelEnd)
End Sub

Private Property Get pMenu() As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a menu object for editing the richedit properties.
'---------------------------------------------------------------------------------------
    
    Set pMenu = pop.NewMenu()
    With pMenu
        .Add "Effective Immediately", , , , mnuSeparator
        .Add "AutoURLDetect", "Toggles whether link references are recognized during stream in operations.", , , mnuChecked * -CBool(rtxt.AutoURLDetect), , , propAutoURLDetect
        .Add "Enabled", "Toggles whether the control accepts user input.", , , mnuChecked * -CBool(rtxt.Enabled), , , propEnabled
        .Add "HideSelection", "Toggles whether the selection is shown when the control is not in focus", , , mnuChecked * -CBool(rtxt.HideSelection), , , propHideSelection
        .Add "Modified", "Toggles the flag that indicates whether the text has been modified since last loaded.", , , mnuChecked * -CBool(rtxt.Modified), , , propModified
        .Add "ReadOnly", "Toggles whether the control can be modified by the user.", , , mnuChecked * -CBool(rtxt.ReadOnly), , , propReadOnly
        .Add "WantTab", "Toggles whether the Tab key inserts a tab character or moves focus.", , , mnuChecked * -CBool(rtxt.WantTab), , , propWantTab
        .Add "WordWrap", "Toggles whether the text wraps at the right edge of the control.", , , mnuChecked * -CBool(rtxt.WordWrap), , , propWordWrap
        .Add "MaxLength (" & rtxt.MaxLength & ")", "Sets the maximum number of characters that can be held in the control.", , , , , , propMaxLength
        
        With .Add("Border").SubMenu
            .Add "None", "No Border.", , , mnuRadioChecked * -CBool(rtxt.BorderStyle = vbccBorderNone), , , propBorderNone
            .Add "Fixed Single", "One line border.", , , mnuRadioChecked * -CBool(rtxt.BorderStyle = vbccBorderSingle), , , propBorderSingle
            .Add "Thin", "Thin 3D border.", , , mnuRadioChecked * -CBool(rtxt.BorderStyle = vbccBorderThin), , , propBorderThin
            .Add "3D Sunken", "Sunken 3D border.", , , mnuRadioChecked * -CBool(rtxt.BorderStyle = vbccBorderSunken), , , propBorder3DSunken
        End With
        
        With .Add("Margins").SubMenu
            .Add "Left (" & rtxt.LeftMargin & ")", "Sets the left margin of the control.", , , , , , propLeftMargin
            .Add "Right (" & rtxt.RightMargin & ")", "Sets the right margin of the control.", , , , , , propRightMargin
        End With
        
        .Add "Need Recreation", , , , mnuSeparator
        .Add "AutoHScroll", "Toggles whether the control scrolls horizontally to keep the cursor in view.", , , mnuChecked * -CBool(rtxt.AutoHScroll), , , propAutoHScroll
        .Add "AutoVScroll", "Toggles whether the control scrolls vertically to keep the cursor in view.", , , mnuChecked * -CBool(rtxt.AutoVScroll), , , propAutoVScroll
        .Add "DisableNoScroll", "Toggles whether the scrollbars are disabled instead of hidden when not in use.", , , mnuChecked * -CBool(rtxt.DisableNoScroll), , , propDisableNoScroll
        .Add "MultiLine", "Toggles whether the control accepts text on multiple lines.", , , mnuChecked * -CBool(rtxt.MultiLine), , , propMultiLine
        .Add "PlainText", "Toggles whether the control accepts formatted text.", , , mnuChecked * -CBool(rtxt.PlainText), , , propPlainText
        .Add "SelectionBar", "Toggles whether the user can select a line of text at a time by clicking the left edge of the control.", , , mnuChecked * -CBool(rtxt.SelectionBar), , , propSelectionBar
        
        With .Add("Scrollbars").SubMenu
            .Add "Both", "Vertical and horizontal scrollbars.", , , mnuRadioChecked * -CBool(rtxt.ScrollBars = rtfScrollBarsBoth), , , propScrollBoth
            .Add "Vertical", "Vertical scrollbar only.", , , mnuRadioChecked * -CBool(rtxt.ScrollBars = rtfScrollBarsVertical), , , propScrollVert
            .Add "Horizontal", "Horizontal scrollbar only.", , , mnuRadioChecked * -CBool(rtxt.ScrollBars = rtfScrollBarsHorizontal), , , propScrollHorz
            .Add "None", "No scrollbars.", , , mnuRadioChecked * -CBool(rtxt.ScrollBars = rtfScrollBarsNone), , , propScrollNone
        End With
        .Add , , , , mnuSeparator
        .Add "Recreate the control", "Recreate the richedit control, putting into effect all new property settings.", , , , , , propRecreate
    End With
    
    RaiseEvent GetPopupSettings(pop, pMenu)
    
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Public Sub StressTest()
    Dim lsText As String
    lsText = "This is some testing text"
    
    Dim i As Long
    For i = 1 To 50
        rtxt.Text = rtxt.Text & vbNewLine & lsText
    Next
End Sub
