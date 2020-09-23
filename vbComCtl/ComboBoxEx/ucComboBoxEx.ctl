VERSION 5.00
Begin VB.UserControl ucComboBoxEx 
   BackColor       =   &H80000005&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   HasDC           =   0   'False
   PropertyPages   =   "ucComboBoxEx.ctx":0000
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   ToolboxBitmap   =   "ucComboBoxEx.ctx":000D
End
Attribute VB_Name = "ucComboBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucComboBoxEx.ctl                  12/15/04
'
'           PURPOSE:
'               Implement WC_COMBOBOXEX from ComCtl32.dll.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Combo_and_List_Boxes/ComboBoxEx/VB6_ComboBoxEx_Full_Source.asp
'               vbalCboEx6.ctl
'
'==================================================================================================

Option Explicit

Public Enum eComboBoxExStyle
    cboSimple
    cboDropDownCombo
    cboDropDownList
End Enum

Public Enum eComboBoxExEndEditReason
    cboEndEditKillFocus = CBENF_KILLFOCUS
    cboEndEditReturn = CBENF_RETURN
    cboEndEditEscape = CBENF_ESCAPE
    cboEndEditDropDown = CBENF_DROPDOWN
End Enum

Public Event DropDown()
Public Event CloseUp()
Public Event EditChange()
Public Event ListIndexChange()
Public Event BeginEdit()
Public Event EndEdit(ByVal bEditChanged As Boolean, ByRef iNewIndex As Long, ByRef sText As String, ByVal iWhy As eComboBoxExEndEditReason)

Implements iSubclass
Implements iOleInPlaceActiveObjectVB

Private Const PROP_Font             As String = "Fnt"
Private Const PROP_Style            As String = "Sty"
Private Const PROP_DroppedHeight    As String = "DHt"
Private Const PROP_DroppedWidth     As String = "DWd"
Private Const PROP_Enabled          As String = "Enbld"
Private Const PROP_MaxLength        As String = "MaxLen"
Private Const PROP_ExtendedUI       As String = "ExtUI"
Private Const PROP_Themeable        As String = "Them"

Private Const DEF_Style             As Long = cboDropDownCombo
Private Const DEF_DroppedHeight     As Long = 160
Private Const DEF_DroppedWidth      As Long = 0
Private Const DEF_Enabled           As Boolean = True
Private Const DEF_MaxLength         As Long = ZeroL
Private Const DEF_ExtendedUI        As Boolean = False
Private Const DEF_Themeable         As Boolean = True

Private WithEvents moFont           As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage       As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1
Private WithEvents moImageListEvent As cImageList
Attribute moImageListEvent.VB_VarHelpID = -1
Private moImageList                 As cImageList

Private mhWnd                       As Long
Private mhWndCombo                  As Long
Private mhWndEdit                   As Long

Private mhFont                      As Long

Private miStyle                     As eComboBoxExStyle
Private miNewIndex                  As Long

Private miDroppedHeight             As Long
Private miDroppedWidth              As Long

Private miMaxLength                 As Long
Private mbRedraw                    As Boolean
Private mbExtendedUI                As Boolean
Private mbThemeable                 As Boolean

Private mtItem                      As COMBOBOXEXITEM
Private msTextBuffer                As String

Private Const ucComboBoxEx          As String = "ucComboBoxEx"

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Update the font if it is set to use the ambient font.
'---------------------------------------------------------------------------------------
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Initialize modular vars, load the shell module to prevent crashes when linked
'             to CC version 6 and install vtable subclassing.
'---------------------------------------------------------------------------------------
    msTextBuffer = Space$(MAX_PATH \ TwoL)
    LoadShellMod
    InitCC ICC_USEREX_CLASSES
    Set moFontPage = New pcSupportFontPropPage
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Initialize property values to the defaults.
'---------------------------------------------------------------------------------------
    Set moFont = Font_CreateDefault(Ambient.Font)
    miStyle = DEF_Style
    miDroppedHeight = DEF_DroppedHeight
    miDroppedWidth = DEF_DroppedWidth
    UserControl.Enabled = DEF_Enabled
    miMaxLength = DEF_MaxLength
    mbExtendedUI = DEF_ExtendedUI
    mbRedraw = True
    mbThemeable = DEF_Themeable
    pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Read property values from a previously persisted instance.
'---------------------------------------------------------------------------------------
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    miStyle = PropBag.ReadProperty(PROP_Style, DEF_Style)
    miDroppedHeight = PropBag.ReadProperty(PROP_DroppedHeight, DEF_DroppedHeight)
    miDroppedWidth = PropBag.ReadProperty(PROP_DroppedWidth, DEF_DroppedWidth)
    UserControl.Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    miMaxLength = PropBag.ReadProperty(PROP_MaxLength, DEF_MaxLength)
    mbExtendedUI = PropBag.ReadProperty(PROP_ExtendedUI, DEF_ExtendedUI)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    mbRedraw = True
    pCreate
End Sub

Private Sub UserControl_Terminate()
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Destroy the comboboxex window, remove subclassing and release the shell
'             module handle.
'---------------------------------------------------------------------------------------
    Set moFontPage = Nothing
    pDestroy
    If mhFont Then moFont.ReleaseHandle mhFont
    ReleaseShellMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Persist our property values.
'---------------------------------------------------------------------------------------
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Style, miStyle, DEF_Style
    PropBag.WriteProperty PROP_DroppedHeight, miDroppedHeight, DEF_DroppedHeight
    PropBag.WriteProperty PROP_DroppedWidth, miDroppedWidth, DEF_DroppedWidth
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_MaxLength, miMaxLength, DEF_MaxLength
    PropBag.WriteProperty PROP_ExtendedUI, mbExtendedUI, DEF_ExtendedUI
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub moFont_Changed()
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Update the control's font when it changes.
'---------------------------------------------------------------------------------------
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    pPropChanged PROP_Font
End Sub

Private Sub moImageListEvent_Changed()
'---------------------------------------------------------------------------------------
' Date      : 9/09/05
' Purpose   : Update the control's imagelist when it changes.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If Not moImageList Is Nothing _
            Then SendMessage mhWnd, CBEM_SETIMAGELIST, ZeroL, moImageList.hIml _
            Else SendMessage mhWnd, CBEM_SETIMAGELIST, ZeroL, ZeroL
    End If
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Intercept arrow keys, home/end/pageup/pagedown and return keys.
'---------------------------------------------------------------------------------------
    
    Select Case uMsg
    Case WM_KEYDOWN, WM_KEYUP
        Select Case wParam And &HFFFF&
        Case vbKeyPageUp To vbKeyDown, vbKeyReturn, vbKeyEscape
            If (((wParam And &HFFFF&) = vbKeyReturn) Or ((wParam And &HFFFF&) = vbKeyEscape)) Then
                'only eat the return/esc keys if the combo is dropped
                If SendMessage(mhWnd, CB_GETDROPPEDSTATE, ZeroL, ZeroL) = ZeroL Then Exit Sub
            End If
            
            Dim lhWndFocus As Long
            lhWndFocus = GetFocus()
            
            If lhWndFocus Then
                
                Select Case lhWndFocus
                Case mhWnd, mhWndCombo, mhWndEdit
                    SendMessage lhWndFocus, uMsg, wParam, lParam
                    bHandled = True
                End Select
                
            End If
           
        End Select
    End Select
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Subclass message processing after the original procedure.
'---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_SIZE
        pResize
    Case WM_SETFOCUS
        If hWnd = UserControl.hWnd _
            Then vbComCtlTlb.SetFocus mhWnd _
            Else ActivateIPAO Me
    Case WM_KILLFOCUS
        DeActivateIPAO Me
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Subclass message processing before the original procedure.
'---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_COMMAND
        bHandled = True
        lReturn = ZeroL
        If mhWnd Then
            If lParam = mhWnd Then
                Select Case ((wParam And &HFFFF0000) \ &H10000)
                Case CBN_DROPDOWN
                    RaiseEvent DropDown
                Case CBN_CLOSEUP
                    RaiseEvent CloseUp
                Case CBN_SELCHANGE
                    RaiseEvent EditChange
                    RaiseEvent ListIndexChange
                Case CBN_EDITCHANGE
                    If mhWndCombo Then
                        If (wParam And &HFFFF&) = GetWindowLong(mhWndCombo, GWL_ID) Then
                            RaiseEvent EditChange
                        End If
                    Else
                        RaiseEvent EditChange
                    End If
                End Select
            End If
        End If
    Case WM_NOTIFY
        bHandled = True
        Const NMHDR_hwndFrom As Long = ZeroL
        Const NMHDR_code As Long = 8
        If mhWnd Then
            If MemOffset32(lParam, NMHDR_hwndFrom) = mhWnd Then
                Select Case MemOffset32(lParam, NMHDR_code)
                Case CBEN_BEGINEDIT
                    RaiseEvent BeginEdit
                Case CBEN_ENDEDIT
                    pOnEndEdit lParam
                End Select
            End If
        End If
    Case WM_SETFOCUS
        If hWnd = UserControl.hWnd _
            Then vbComCtlTlb.SetFocus mhWnd _
            Else ActivateIPAO Me
        
    Case WM_MOUSEACTIVATE
        Dim lhWnd As Long
        lhWnd = GetFocus()
        
        If lhWnd <> mhWnd And lhWnd <> mhWndCombo And lhWnd <> mhWndEdit Then
            bHandled = True
            lReturn = MA_NOACTIVATE
            vbComCtlTlb.SetFocus UserControl.hWnd
        End If
   End Select
End Sub

Private Sub pOnEndEdit(ByRef lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Raise the EndEdit Event with data from the NMCBEENDEDIT Structure.
'---------------------------------------------------------------------------------------
    Const NMCBEENDEDIT_fChanged As Long = 12
    Const NMCBEENDEDIT_iNewSelection As Long = 16
    Const NMCBEENDEDIT_szText As Long = 20
    Const NMCBEENDEDIT_iWhy As Long = 280
    
    Dim sText As String
    Dim liSelection As Long
    
    lstrToStringA UnsignedAdd(lParam, NMCBEENDEDIT_szText), sText
    liSelection = MemOffset32(lParam, NMCBEENDEDIT_iNewSelection)
    RaiseEvent EndEdit(CBool(MemOffset32(lParam, NMCBEENDEDIT_fChanged)), _
                             liSelection, _
                             sText, _
                             MemOffset16(lParam, NMCBEENDEDIT_iWhy))
    MemOffset32(lParam, NMCBEENDEDIT_iNewSelection) = liSelection
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return a proxy object to receive notifications from the font property page.
'---------------------------------------------------------------------------------------
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Tell the ppFont page which font properties this control implements.
'---------------------------------------------------------------------------------------
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Provide the property page with an ambient font.
'---------------------------------------------------------------------------------------
    Set o = Ambient.Font
End Sub

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Create the comboex window and install the necessary subclasses.
'---------------------------------------------------------------------------------------
    pDestroy
    
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_COMBOBOXEX & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, (miStyle + OneL) Or WS_CHILD Or CBS_AUTOHSCROLL, ZeroL, ZeroL, UserControl.ScaleWidth, UserControl.ScaleHeight + miDroppedHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        
        If Not moImageList Is Nothing Then SendMessage mhWnd, CBEM_SETIMAGELIST, ZeroL, moImageList.hIml
        SendMessage mhWnd, CB_SETDROPPEDWIDTH, miDroppedWidth, ZeroL
        SendMessage mhWnd, CB_LIMITTEXT, miMaxLength, ZeroL
        SendMessage mhWnd, CB_SETEXTENDEDUI, -mbExtendedUI, ZeroL
        SendMessage mhWnd, WM_SETREDRAW, -mbRedraw, ZeroL
        
        mhWndCombo = SendMessage(mhWnd, CBEM_GETCOMBOCONTROL, ZeroL, ZeroL)
        mhWndEdit = SendMessage(mhWnd, CBEM_GETEDITCONTROL, ZeroL, ZeroL)
        
        If mhWndEdit = ZeroL And miStyle <> cboDropDownList Then
            mhWndEdit = FindWindowExW(mhWndCombo, ZeroL, "Edit", vbNullString)
            Debug.Assert mhWndEdit
        End If
        
        If Ambient.UserMode Then
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_COMMAND, WM_NOTIFY), Array(WM_SETFOCUS, WM_SIZE)
            Subclass_Install Me, mhWnd, WM_MOUSEACTIVATE, Array(WM_SETFOCUS, WM_KILLFOCUS)
            Subclass_Install Me, mhWndCombo, Array(WM_MOUSEACTIVATE, WM_SETFOCUS), WM_KILLFOCUS
            
            If mhWndEdit Then Subclass_Install Me, mhWndEdit, Array(WM_MOUSEACTIVATE, WM_SETFOCUS), WM_KILLFOCUS
            
            VTableSubclass_IPAO_Install Me
            
        End If
        
        ShowWindow mhWnd, SW_SHOWNORMAL
        
        EnableWindowTheme mhWnd, mbThemeable
        If mhWndCombo Then EnableWindowTheme mhWndCombo, mbThemeable
        If mhWndEdit Then EnableWindowTheme mhWndEdit, mbThemeable
        
        pResize
        pSetFont
        
    End If
    
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Destroy the comboex window and subclasses.
'---------------------------------------------------------------------------------------
    
    If mhWnd Then
    
        VTableSubclass_IPAO_Remove
    
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        Subclass_Remove Me, mhWndCombo
        If mhWndEdit Then Subclass_Remove Me, mhWndEdit
        
        DestroyWindow mhWnd
        
        mhWnd = ZeroL
        mhWndCombo = ZeroL
        mhWndEdit = ZeroL
    
    End If
    
End Sub

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Update the handle stored in mhFont and notify the comboex window.
'---------------------------------------------------------------------------------------
    
    Dim hFont As Long: hFont = moFont.GetHandle
    If mhWnd Then SendMessage mhWnd, WM_SETFONT, hFont, OneL
    If mhFont Then moFont.ReleaseHandle mhFont
    mhFont = hFont
    pAutoSizeHeight
    
End Sub

Private Sub pAutoSizeHeight()
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Ensure that the usercontrol is the correct height.
'             If the combobox is contained on a rebar band, ask it to resize.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        
        UserControl.SIZE Width, ScaleY(SendMessage(mhWnd, CB_GETITEMHEIGHT, NegOneL, ZeroL), vbPixels, vbTwips)
        
        Dim lhWndParent As Long
        lhWndParent = GetParent(UserControl.hWnd)
        
        Dim lsAnsi As String
        Dim lsClassName As String
        lsClassName = Space$(MAX_PATH \ 2)
        
        If lhWndParent Then
            lsAnsi = StrConv(WC_REBAR & vbNullChar, vbFromUnicode)
            GetClassName lhWndParent, ByVal StrPtr(lsClassName), LenB(lsClassName)
            If lstrcmp(StrPtr(lsAnsi), StrPtr(lsClassName)) = ZeroL Then
                lhWndParent = GetParent(lhWndParent)
                If lhWndParent Then
                    lsAnsi = StrConv("ThunderUserControl" & vbNullChar, vbFromUnicode)
                    GetClassName lhWndParent, ByVal StrPtr(lsClassName), LenB(lsClassName)
                    
                    If lstrcmp(StrPtr(lsAnsi), StrPtr(lsClassName)) = ZeroL Then
                        PostMessage lhWndParent, UM_SIZEBAND, UserControl.hWnd, MakeLong(0, ScaleHeight And &HFFFF&)
                    Else
                        lsAnsi = StrConv("ThunderRT6UserControl" & vbNullChar, vbFromUnicode)
                        If lstrcmp(StrPtr(lsAnsi), StrPtr(lsClassName)) = ZeroL Then
                            PostMessage lhWndParent, UM_SIZEBAND, UserControl.hWnd, MakeLong(0, ScaleHeight And &HFFFF&)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub pResize()
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Try to size the comboex to the same size as the usercontrol.
'             Size the usercontrol to the resulting size of the comboex window.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
        
        Dim tR As RECT
        If GetWindowRect(mhWnd, tR) Then
            UserControl.SIZE ScaleX(tR.Right - tR.Left, vbPixels, vbTwips), ScaleY(tR.bottom - tR.Top, vbPixels, vbTwips)
        End If
        
    End If
End Sub

Private Sub pPropChanged(ByRef s As String)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Determine whether a given hwnd is one of our comboex or child windows.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode = False Then PropertyChanged s
End Sub

Private Property Get pItem_Text(ByVal iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the text of a combo list item.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .Mask = CBEIF_TEXT
            .pszText = StrPtr(msTextBuffer)
            .cchTextMax = LenB(msTextBuffer)
            .iItem = iIndex
            If SendMessage(mhWnd, CBEM_GETITEM, ZeroL, VarPtr(mtItem)) _
                Then lstrToStringA .pszText, pItem_Text
        End With
    End If
End Property
Private Property Let pItem_Text(ByVal iIndex As Long, ByRef sText As String)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the text of a combo list item.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .Mask = CBEIF_TEXT
            .iItem = iIndex
            Dim ls As String
            ls = StrConv(sText & vbNullChar, vbFromUnicode)
            .pszText = StrPtr(ls)
            .cchTextMax = LenB(ls)
            SendMessage mhWnd, CBEM_SETITEM, ZeroL, VarPtr(mtItem)
        End With
    End If
End Property

Private Property Get pItem_Info(ByVal iIndex As Long, ByVal iMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get a 32 bit value in the COMBOBOXEXITEM structure.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .Mask = iMask
            .iItem = iIndex
            If SendMessage(mhWnd, CBEM_GETITEM, ZeroL, VarPtr(mtItem)) Then
                If iMask = CBEIF_LPARAM Then
                    pItem_Info = .lParam
                ElseIf iMask = CBEIF_IMAGE Then
                    pItem_Info = .iImage
                ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                    pItem_Info = .iSelectedImage
                ElseIf iMask = CBEIF_INDENT Then
                    pItem_Info = .iIndent
                End If
            End If
        End With
    End If
End Property
Private Property Let pItem_Info(ByVal iIndex As Long, ByVal iMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set a 32 bit value in the COMBOBOXEXITEM structure.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtItem
            .Mask = iMask
            .iItem = iIndex
            If iMask = CBEIF_LPARAM Then
                .lParam = iNew
            ElseIf iMask = CBEIF_IMAGE Then
                .iImage = iNew
            ElseIf iMask = CBEIF_SELECTEDIMAGE Then
                .iSelectedImage = iNew
            ElseIf iMask = CBEIF_INDENT Then
                .iIndent = iNew
            End If
            SendMessage mhWnd, CBEM_SETITEM, ZeroL, VarPtr(mtItem)
        End With
    End If
End Property

Public Property Get Style() As eComboBoxExStyle
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the style of combo box.
'---------------------------------------------------------------------------------------
    Style = miStyle
End Property

Public Property Let Style(ByVal iNew As eComboBoxExStyle)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the style of combo box.
'---------------------------------------------------------------------------------------
    miStyle = iNew And 3&
    pCreate
    pPropChanged PROP_Style
End Property

Public Function AddItem( _
            ByRef sText As String, _
   Optional ByVal iIconIndex As Long = NegOneL, _
   Optional ByVal iIconIndexSelected As Long = NegOneL, _
   Optional ByVal iItemData As Long, _
   Optional ByVal iIndent As Long, _
   Optional ByVal iIndexInsertBefore As Long = NegOneL) _
                As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Add a combobox item.
'---------------------------------------------------------------------------------------
    
    If iIconIndexSelected < ZeroL Then iIconIndexSelected = iIconIndex
    
    If mhWnd Then
        With mtItem
            .Mask = CBEIF_TEXT Or CBEIF_LPARAM Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_INDENT Or CBEIF_LPARAM
            Dim ls As String
            ls = StrConv(sText & vbNullChar, vbFromUnicode)
            MidB$(msTextBuffer, OneL, LenB(ls)) = ls
            MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = vbNullChar
            
            .pszText = StrPtr(msTextBuffer)
            .cchTextMax = LenB(msTextBuffer)
            .iItem = iIndexInsertBefore
            .iImage = iIconIndex
            .iSelectedImage = iIconIndexSelected
            .iOverlay = NegOneL
            .iIndent = iIndent
            .lParam = iItemData
            If iIndexInsertBefore < NegOneL Then iIndexInsertBefore = NegOneL
            miNewIndex = SendMessage(mhWnd, CBEM_INSERTITEM, iIndexInsertBefore, VarPtr(mtItem))
            AddItem = CBool(miNewIndex > NegOneL)
        End With
    End If
End Function

Public Function RemoveItem(ByVal iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Remove a combobox item.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        RemoveItem = CBool(SendMessage(mhWnd, CBEM_DELETEITEM, iIndex, ZeroL) > NegOneL)
        If RemoveItem Then
            If iIndex = miNewIndex Then
                miNewIndex = NegOneL
            ElseIf iIndex < miNewIndex Then
                miNewIndex = miNewIndex - OneL
            End If
        End If
    End If
End Function

Public Property Get ImageList() As cImageList
Attribute ImageList.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the imagelist in use by this control.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode = False Then gErr vbccGetNoDesignTime, ucComboBoxEx
    Set ImageList = moImageList
End Property
Public Property Set ImageList(ByVal oNew As cImageList)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the imagelist used by this control.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode = False Then gErr vbccLetSetNoDesignTime, ucComboBoxEx
    On Error Resume Next
    Set moImageList = Nothing
    Set moImageListEvent = Nothing
    Set moImageList = oNew
    Set moImageListEvent = oNew
    On Error GoTo 0
    moImageListEvent_Changed
    pResize
End Property

Public Property Get DroppedWidth() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the width of the dropdown portion of the control.
'---------------------------------------------------------------------------------------
    DroppedWidth = ScaleX(miDroppedWidth, vbPixels, vbContainerSize)
End Property
Public Property Let DroppedWidth(ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the width of the dropdown portion of the control.
'---------------------------------------------------------------------------------------
    miDroppedWidth = ScaleX(fNew, vbContainerSize, vbPixels)
    If mhWnd Then
        SendMessage mhWnd, CB_SETDROPPEDWIDTH, miDroppedWidth, ZeroL
    End If
    pPropChanged PROP_DroppedWidth
End Property

Public Property Get DroppedHeight() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the Height of the dropdown portion of the control.
'---------------------------------------------------------------------------------------
    DroppedHeight = ScaleY(miDroppedHeight, vbPixels, vbContainerSize)
End Property
Public Property Let DroppedHeight(ByRef fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the Height of the dropdown portion of the control.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode Then gErr vbccLetSetNoRunTime, ucComboBoxEx
    miDroppedHeight = ScaleY(fNew, vbContainerSize, vbPixels)
    pPropChanged PROP_DroppedHeight
End Property

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the font object for this control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the font object for this control.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    pSetFont
    pPropChanged PROP_Font
End Property

Public Property Get NewIndex() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the index of the last item added.
'---------------------------------------------------------------------------------------
    NewIndex = miNewIndex
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return whether the control is enabled.
'---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether the control is enabled.
'---------------------------------------------------------------------------------------
    UserControl.Enabled = bNew
    pPropChanged PROP_Enabled
End Property

Public Property Get Dropped() As Boolean
Attribute Dropped.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return whether the drop down is visible.
'---------------------------------------------------------------------------------------
    If miStyle > cboSimple Then
        If Ambient.UserMode = False Then gErr vbccGetNoDesignTime, ucComboBoxEx
        If mhWnd Then Dropped = CBool(SendMessage(mhWnd, CB_GETDROPPEDSTATE, ZeroL, ZeroL))
    End If
End Property

Public Property Let Dropped(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether the drop down is visible.
'---------------------------------------------------------------------------------------
    If miStyle > cboSimple Then
        If Ambient.UserMode = False Then gErr vbccLetSetNoDesignTime, ucComboBoxEx
        If mhWnd Then SendMessage mhWnd, CB_SHOWDROPDOWN, -bNew, ZeroL
    End If
End Property


Public Property Get ListCount() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the number of items in the combobox.
'---------------------------------------------------------------------------------------
    If mhWnd Then ListCount = SendMessage(mhWnd, CB_GETCOUNT, ZeroL, ZeroL)
End Property
Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the index of the selected item in the list.
'---------------------------------------------------------------------------------------
    If mhWnd Then ListIndex = SendMessage(mhWnd, CB_GETCURSEL, ZeroL, ZeroL)
End Property
Public Property Let ListIndex(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the index of the selected item in the list.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If SendMessage(mhWnd, CB_SETCURSEL, iNew, ZeroL) > CB_ERR Then RaiseEvent ListIndexChange
    End If
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the selstart of the edit portion (if any) of the combobox.
'---------------------------------------------------------------------------------------
    Dim i As Long
    If mhWndEdit Then
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(SelStart), VarPtr(i)
    End If
End Property
Public Property Let SelStart(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the selstart of the edit portion (if any) of the combobox.
'---------------------------------------------------------------------------------------
    If mhWndEdit Then
        Dim iStart As Long, iEnd As Long
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(iStart), VarPtr(iEnd)
        If iEnd > iNew _
            Then SendMessage mhWndEdit, EM_SETSEL, iNew, iEnd _
            Else SendMessage mhWndEdit, EM_SETSEL, iNew, iNew
    End If
End Property

Public Property Get SelEnd() As Long
Attribute SelEnd.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the selend of the edit portion (if any) of the combobox.
'---------------------------------------------------------------------------------------
    Dim i As Long
    If mhWndEdit Then
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(i), VarPtr(SelEnd)
    End If
End Property
Public Property Let SelEnd(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the selend of the edit portion (if any) of the combobox.
'---------------------------------------------------------------------------------------
    If mhWndEdit Then
        Dim iStart As Long, iEnd As Long
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(iStart), VarPtr(iEnd)
        If iStart < iNew _
            Then SendMessage mhWndEdit, EM_SETSEL, iStart, iNew _
            Else SendMessage mhWndEdit, EM_SETSEL, iNew, iNew
    End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the sellength of the edit portion (if any) of the combobox.
'---------------------------------------------------------------------------------------
    If mhWndEdit Then
        Dim iStart As Long, iEnd As Long
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(iStart), VarPtr(iEnd)
        SelLength = iEnd - iStart
    End If
End Property
Public Property Let SelLength(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the sellength of the edit portion (if any) of the combobox.
'---------------------------------------------------------------------------------------
    If mhWndEdit Then
        Dim iStart As Long, iEnd As Long
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(iStart), VarPtr(iEnd)
        SendMessage mhWndEdit, EM_SETSEL, iStart, iStart + iNew
    End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the seltext of the edit portion (if any) of the combobox.
'             If there is no edit box, return the text of the control.
'---------------------------------------------------------------------------------------
    If mhWndEdit Then
        Dim iStart As Long
        Dim iEnd As Long
        SendMessage mhWndEdit, EM_GETSEL, VarPtr(iStart), VarPtr(iEnd)
        On Error Resume Next
        SelText = Mid$(Text, iStart, (iEnd - iStart))
        On Error GoTo 0
    Else
        SelText = Text
    
    End If
End Property

Public Property Get Text() As String
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the selected list item or the text in the edit portion of the control.
'---------------------------------------------------------------------------------------
    If mhWndEdit Then
        
        Dim liTextLength As Long
        liTextLength = GetWindowTextLength(hWnd)
        
        If liTextLength > ZeroL Then
            Dim lpText As String
            lpText = MemAllocFromString(ZeroL, liTextLength)
            
            If lpText Then
                liTextLength = GetWindowText(hWnd, lpText, liTextLength)
                If liTextLength Then lstrToStringA lpText, Text, liTextLength
                MemFree lpText
            End If
                
        End If
        
        
    ElseIf mhWnd Then
        
        Text = pItem_Text(SendMessage(mhWnd, CB_GETCURSEL, ZeroL, ZeroL))
        
    End If
End Property

Public Property Let Text(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : If there is an edit box, set the text.  Otherwise, search the list for an
'             item that matches sNew.  If found, set the listindex.  If not, raise an error.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lsAnsi As String
        lsAnsi = StrConv(sNew & vbNullChar, vbFromUnicode)
        
        If mhWndEdit Then
            SetWindowText mhWndEdit, StrPtr(lsAnsi)
            
        Else
            Dim liIndex As Long
            
            liIndex = SendMessage(mhWnd, CB_FINDSTRINGEXACT, ZeroL, StrPtr(lsAnsi))
            
            If liIndex > NegOneL _
                Then SendMessage mhWnd, CB_SETCURSEL, liIndex, ZeroL _
                Else gErr vbccInvalidProcedureCall, ucComboBoxEx
            
        End If
    End If
End Property

Public Property Get FindItem(ByRef sText As String, Optional ByVal bExact As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Search the list for an item and return the index if found.  -1 otherwise.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim ls As String
        ls = StrConv(sText & vbNullChar, vbFromUnicode)
        FindItem = SendMessage(mhWnd, IIf(bExact, CB_FINDSTRINGEXACT, CB_FINDSTRING), ZeroL, StrPtr(ls))
    End If
End Property

Public Sub Clear()
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Delete all items in the list.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, CB_RESETCONTENT, ZeroL, ZeroL
        miNewIndex = NegOneL
    End If
End Sub

Public Property Get ItemIndent(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the indentation level of the specified item.
'---------------------------------------------------------------------------------------
    ItemIndent = pItem_Info(iIndex, CBEIF_INDENT)
End Property
Public Property Let ItemIndent(ByVal iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the indentation level of the specified item.
'---------------------------------------------------------------------------------------
    pItem_Info(iIndex, CBEIF_INDENT) = iNew
End Property

Public Property Get ItemIconIndex(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the iconindex of the specified item.
'---------------------------------------------------------------------------------------
    ItemIconIndex = pItem_Info(iIndex, CBEIF_IMAGE)
End Property
Public Property Let ItemIconIndex(ByVal iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the iconindex of the specified item.
'---------------------------------------------------------------------------------------
    pItem_Info(iIndex, CBEIF_IMAGE) = iNew
End Property

Public Property Get ItemIconIndexSelected(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the selected iconindex of the specified item.
'---------------------------------------------------------------------------------------
    ItemIconIndexSelected = pItem_Info(iIndex, CBEIF_SELECTEDIMAGE)
End Property
Public Property Let ItemIconIndexSelected(ByVal iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the selected iconindex of the specified item.
'---------------------------------------------------------------------------------------
    pItem_Info(iIndex, CBEIF_SELECTEDIMAGE) = iNew
End Property

Public Property Get ItemData(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the itemdata of the specified item.
'---------------------------------------------------------------------------------------
    ItemData = pItem_Info(iIndex, CBEIF_LPARAM)
End Property
Public Property Let ItemData(ByVal iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the itemdata of the specified item.
'---------------------------------------------------------------------------------------
    pItem_Info(iIndex, CBEIF_LPARAM) = iNew
End Property

Public Property Get ItemText(ByVal iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Get the text of the specified item.
'---------------------------------------------------------------------------------------
    ItemText = pItem_Text(iIndex)
End Property
Public Property Let ItemText(ByVal iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the text of the specified item.
'---------------------------------------------------------------------------------------
    pItem_Text(iIndex) = sNew
End Property

Public Property Get MaxLength() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the maximum allowable number of characters in the edit portion of the window.
'---------------------------------------------------------------------------------------
    MaxLength = miMaxLength
End Property
Public Property Let MaxLength(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set the maximum allowable number of characters in the edit portion of the window.
'---------------------------------------------------------------------------------------
    If iNew > &H7FFFFFFF Then iNew = &H7FFFFFFF
    miMaxLength = iNew
    If mhWnd Then SendMessage mhWnd, CB_LIMITTEXT, miMaxLength, ZeroL
    pPropChanged PROP_MaxLength
End Property

Public Property Get ExtendedUI() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return whether we are using the extended ui features available through comctl32.dll.
'---------------------------------------------------------------------------------------
    ExtendedUI = mbExtendedUI
End Property
Public Property Let ExtendedUI(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether we are using the extended ui features available through comctl32.dll.
'---------------------------------------------------------------------------------------
    mbExtendedUI = bNew
    pPropChanged PROP_ExtendedUI
    If mhWnd Then SendMessage mhWnd, CB_SETEXTENDEDUI, -mbExtendedUI, ZeroL
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return whether redrawing items is enabled.
'---------------------------------------------------------------------------------------
    Redraw = mbRedraw
End Property
Public Property Let Redraw(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether redrawing items is enabled.  This can increase performance
'             when adding multiple items to a simple style combobox.
'---------------------------------------------------------------------------------------
    mbRedraw = bNew
    If mhWnd Then SendMessage mhWnd, WM_SETREDRAW, -mbRedraw, ZeroL
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property
Public Property Get hWndComboEx() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the ComboBoxEx.
'---------------------------------------------------------------------------------------
    hWndComboEx = mhWnd
End Property
Public Property Get hWndCombo() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the ComboBox.
'---------------------------------------------------------------------------------------
    hWndCombo = mhWndCombo
End Property
Public Property Get hWndEdit() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return the hwnd of the Edit control.
'---------------------------------------------------------------------------------------
    hWndEdit = mhWndEdit
End Property

Public Property Get Themeable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Return a value indicating whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property
Public Property Let Themeable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/9/05
' Purpose   : Set whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        mbThemeable = bNew
        If mhWnd Then EnableWindowTheme mhWnd, mbThemeable
        If mhWndCombo Then EnableWindowTheme mhWndCombo, mbThemeable
        If mhWndEdit Then EnableWindowTheme mhWndEdit, mbThemeable
        pPropChanged PROP_Themeable
    End If
End Property

