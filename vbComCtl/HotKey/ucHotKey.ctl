VERSION 5.00
Begin VB.UserControl ucHotKey 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   HasDC           =   0   'False
   MousePointer    =   3  'I-Beam
   PropertyPages   =   "ucHotKey.ctx":0000
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ToolboxBitmap   =   "ucHotKey.ctx":000D
End
Attribute VB_Name = "ucHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucHotKey.ctl                9/10/05
'
'           PURPOSE:
'               Implement the Win32 hotkey control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Utility_Controls/HotKey_Control/VB5_HotKey_Control.asp
'               HotKey.ctl
'
'==================================================================================================

Option Explicit

Public Enum eHotKeyModifier
    hotNone = 0
    hotAlt = &H4&
    hotControl = &H2&
    hotShift = &H1&
    hotControlAlt = hotControl Or hotAlt
    hotControlShift = hotShift Or hotControl
    hotControlShiftAlt = hotControl Or hotAlt Or hotShift
    hotShiftAlt = hotAlt Or hotShift
End Enum

Public Enum eHotKeySetAppHotKeyResult
    hotInvalidHotKey
    hotInvalidWindow
    hotSuccess
    hotSuccessWithDuplicate
End Enum

Public Event Change()

Implements iSubclass
Implements iOleInPlaceActiveObjectVB
Implements iOleControlVB

Private Const PROP_Enabled      As String = "Enabled"
Private Const PROP_Font         As String = "Font"
Private Const PROP_Themeable    As String = "Themeable"

Private Const DEF_Enabled       As Boolean = True
Private Const DEF_Themeable     As Boolean = True

Private WithEvents moFont       As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage   As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private mhWnd                   As Long
Private mhFont                  As Long
Private mbThemeable             As Boolean

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the font object if it is set to use the ambient font.
'---------------------------------------------------------------------------------------
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Initialize the common control library and load a shell handle to prevent
'             crashes when linked to CC 6.0.
'---------------------------------------------------------------------------------------
    Set moFontPage = New pcSupportFontPropPage
    LoadShellMod
    InitCC ICC_HOTKEY_CLASS
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Initialize property values.
'---------------------------------------------------------------------------------------
   mbThemeable = DEF_Themeable
   Set moFont = Font_CreateDefault(UserControl.Ambient.Font)
   UserControl.Enabled = DEF_Enabled
   pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Read property values from a previously saved instance.
'---------------------------------------------------------------------------------------
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    pCreate
End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Move the hotkey control to the same size as the usercontrol.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
    End If
End Sub

Private Sub UserControl_Terminate()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Destroy the hotkey, release the shell module and font.
'---------------------------------------------------------------------------------------
    pDestroy
    ReleaseShellMod
    If mhFont Then moFont.ReleaseHandle mhFont
    Set moFontPage = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Save property values between instances.
'---------------------------------------------------------------------------------------
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Tell the property page which font properties we support.
'---------------------------------------------------------------------------------------
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Provide the property page with the ambient font.
'---------------------------------------------------------------------------------------
    Set o = Ambient.Font
End Sub

Private Sub moFont_Changed()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the font in the hotkey control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        moFont.OnAmbientFontChanged Ambient.Font
        Dim lhFont As Long
        lhFont = mhFont
        
        mhFont = moFont.GetHandle()
        SendMessage mhWnd, WM_SETFONT, mhFont, OneL
        If lhFont Then moFont.ReleaseHandle lhFont
        
        If Not Ambient.UserMode Then PropertyChanged PROP_Font
    End If
End Sub

Private Sub iOleControlVB_OnMnemonic(bHandled As Boolean, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long):  End Sub
Private Sub iOleControlVB_GetControlInfo(bHandled As Boolean, iAccelCount As Long, hAccelTable As Long, iFlags As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Inform VB that we eat these characters so that default and cancel buttons will
'             display appropriately.
'---------------------------------------------------------------------------------------
    iAccelCount = ZeroL
    hAccelTable = ZeroL
    iFlags = vbccEatsEscape Or vbccEatsReturn
    bHandled = True
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Forward arrow keys and such to the hotkey control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If (wParam And &HFFFF&) <> VK_TAB Then
            If uMsg = WM_KEYDOWN Or uMsg = WM_SYSKEYDOWN Or uMsg = WM_KEYUP Or uMsg = WM_SYSKEYUP Then
                SendMessage mhWnd, uMsg, wParam, lParam
                lReturn = OneL
                bHandled = True
            End If
        End If
    End If
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Handle subclass notifications after the original procedure.
'---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_SETFOCUS
        vbComCtlTlb.SetFocus mhWnd
    Case WM_KILLFOCUS
        DeActivateIPAO Me
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Handle subclass notifications before the original procedure.
'---------------------------------------------------------------------------------------
    Select Case uMsg
    Case WM_SETFOCUS
        ActivateIPAO Me
    Case WM_MOUSEACTIVATE
        If GetFocus() <> mhWnd Then
            vbComCtlTlb.SetFocus UserControl.hWnd
            lReturn = MA_NOACTIVATE
            bHandled = True
        End If
      
    Case WM_COMMAND
        If ((wParam And &H7FFF0000) \ &H10000) = &H300& Then RaiseEvent Change
        
    End Select

End Sub

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Create the hotkey window and necessary subclasses.
'---------------------------------------------------------------------------------------
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_HOTKEY & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, WS_CHILD Or WS_VISIBLE Or WS_DISABLED * (UserControl.Enabled + OneL), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        If Ambient.UserMode Then
            Subclass_Install Me, UserControl.hWnd, WM_COMMAND, WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_SETFOCUS, WM_MOUSEACTIVATE), WM_KILLFOCUS
            VTableSubclass_OleControl_Install Me
            VTableSubclass_IPAO_Install Me
        End If
        EnableWindowTheme mhWnd, mbThemeable
        moFont_Changed
    End If
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Destroy the hotkey window and subclasses.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    
    If mhWnd Then
    
        VTableSubclass_OleControl_Remove
        VTableSubclass_IPAO_Remove
    
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
        
    End If
    On Error GoTo 0
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a proxy object to receive notifications from the font property page.
'---------------------------------------------------------------------------------------
    Set fSupportFontPropPage = moFontPage
End Property

Public Function SetApplicationHotKey(Optional ByVal hWnd As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the hotkey of the specified window to the current hotkey.
'             return a success code.
'---------------------------------------------------------------------------------------
    If hWnd = ZeroL Then hWnd = UserControl.hWnd
    SetApplicationHotKey = SendMessage(RootParent(hWnd), WM_SETHOTKEY, HotKeyAndModifier(), ZeroL) + OneL
End Function

Public Property Let InvalidHotKeyOperation(ByVal iInvalidModifier As eHotKeyModifier, ByVal iAlternateModifier As eHotKeyModifier, ByVal bState As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the ctrl/alt/shift masks that are disallowed.
'---------------------------------------------------------------------------------------
    iInvalidModifier = pInvalid(iInvalidModifier)
    
    Static liInvalidModifiers As Long
    Static liAlternateModifiers As Long
    
    If bState Then
        liInvalidModifiers = liInvalidModifiers Or (iInvalidModifier And &HFF&)
        liAlternateModifiers = liAlternateModifiers Or (iAlternateModifier And &HFF&)
    Else
        liInvalidModifiers = liInvalidModifiers And Not (iInvalidModifier And &HFF&)
        liAlternateModifiers = liAlternateModifiers And Not (iAlternateModifier And &HFF&)
    End If
   
    If mhWnd Then
        SendMessage mhWnd, HKM_SETRULES, liInvalidModifiers, liAlternateModifiers
    End If
    
End Property

Private Property Get pInvalid(ByVal i As eHotKeyModifier) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the invalid modifier key value from the modifier key value.
'---------------------------------------------------------------------------------------
   Const hotInvalidNone = &H1
   Const hotInvalidShift = &H2
   Const hotInvalidControl = &H4
   Const hotInvalidAlt = &H8
   Const hotInvalidShiftControl = &H10
   Const hotInvalidShiftAlt = &H20
   Const hotInvalidControlAlt = &H40
   Const hotInvalidControlAltShift = &H80
   
   Select Case i
   Case hotAlt:             pInvalid = hotInvalidAlt
   Case hotControl:         pInvalid = hotInvalidControl
   Case hotShift:           pInvalid = hotInvalidShift
   Case hotControlAlt:      pInvalid = hotInvalidControlAlt
   Case hotControlShiftAlt: pInvalid = hotInvalidControlAltShift
   Case hotShiftAlt:        pInvalid = hotInvalidShiftAlt
   Case hotControlShift:    pInvalid = hotInvalidShiftControl
   Case Else:               pInvalid = hotInvalidNone: Debug.Assert False
   End Select
End Property

Public Property Get HotKey() As Long
Attribute HotKey.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Get the current key displayed in the control.
'---------------------------------------------------------------------------------------

    If mhWnd Then
        HotKey = (SendMessage(mhWnd, HKM_GETHOTKEY, 0, 0) And &HFF&)
    End If

End Property

Public Property Let HotKey(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the current key displayed in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        iNew = iNew And &HFF&
        
        Dim liKeyAndMod As Long
        
        liKeyAndMod = SendMessage(mhWnd, HKM_GETHOTKEY, 0, 0)
        
        If iNew <> (liKeyAndMod And &HFF&) Then
            SendMessage mhWnd, HKM_SETHOTKEY, (liKeyAndMod And &HFF00&) Or iNew, 0
        End If
        
    End If

End Property

Public Property Get HotKeyModifier() As eHotKeyModifier
Attribute HotKeyModifier.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Get the current modifiers displayed in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        HotKeyModifier = ((SendMessage(mhWnd, HKM_GETHOTKEY, 0, 0) And &HFF00&) \ &H100&)
    End If
    
End Property

Public Property Let HotKeyModifier(ByVal iNew As eHotKeyModifier)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the current modifiers displayed in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        
        iNew = (iNew And &HFF&) * &H100&
        
        Dim liKeyAndMod As Long
        
        liKeyAndMod = (SendMessage(mhWnd, HKM_GETHOTKEY, 0, 0) And &HFF&)
        
        If iNew <> (liKeyAndMod And &HFF00&) Then
            SendMessage mhWnd, HKM_SETHOTKEY, (liKeyAndMod And &HFF&) Or iNew, 0
        End If
        
    End If
    
End Property

Public Property Get HotKeyAndModifier() As Long
Attribute HotKeyAndModifier.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Get the current modifiers and the current key displayed in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        HotKeyAndModifier = SendMessage(mhWnd, HKM_GETHOTKEY, 0, 0)
    End If
    
End Property

Public Property Let HotKeyAndModifier(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the current modifiers and the current key displayed in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, HKM_SETHOTKEY, iNew, 0
    End If
    
End Property

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the font displayed by the control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the font displayed by the control.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    moFont_Changed
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the enabled status of the control.
'---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the enabled status of the control.
'---------------------------------------------------------------------------------------
    UserControl.Enabled = bNew
    If mhWnd Then
        If bNew _
            Then SetWindowStyle mhWnd, ZeroL, WS_DISABLED _
            Else SetWindowStyle mhWnd, WS_DISABLED, ZeroL
    End If
End Property

Public Property Get Themeable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a value indicating whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        PropertyChanged PROP_Themeable
        mbThemeable = bNew
        If mhWnd Then
            EnableWindowTheme mhWnd, mbThemeable
            SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
            RedrawWindow mhWnd, ByVal ZeroL, ZeroL, RDW_INVALIDATE
        End If
    End If
End Property
