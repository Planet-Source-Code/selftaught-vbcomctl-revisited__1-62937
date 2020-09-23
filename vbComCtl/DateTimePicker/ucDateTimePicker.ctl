VERSION 5.00
Begin VB.UserControl ucDateTimePicker 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   HasDC           =   0   'False
   PropertyPages   =   "ucDateTimePicker.ctx":0000
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   ToolboxBitmap   =   "ucDateTimePicker.ctx":001D
End
Attribute VB_Name = "ucDateTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DateTimePicker control"
'==================================================================================================
'ucDateTimePicker.ctl        9/10/05
'
'           PURPOSE:
'               Implement the Win32 DateTimePicker.
'
'           LINEAGE:
'               http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=9549&lngWId=1
'               DateTimePick.ctl
'
'==================================================================================================

Option Explicit

Public Enum eDateTimePickerFormats
    dtpTime = DTS_TIMEFORMAT
    dtpLongDate = DTS_LONGDATEFORMAT
    dtpShortDate = DTS_SHORTDATEFORMAT
End Enum

Public Event DropDown(ByVal hWndMonthCal As Long)
Public Event CloseUp()
Public Event DateTimeChange()

Implements iSubclass
Implements iOleInPlaceActiveObjectVB

'These props are on/off boolean type props.  They are stored in one long value.
'This means you can't see the property names in the .frm file, but it saves
'numerous properties both in this module's runtime storage and in the .frm file.
Private Enum eBooleanProps
    
    'store the dtp styles in the loword
    bpShortDateFormat = DTS_SHORTDATEFORMAT
    bpUpDown = DTS_UPDOWN
    bpShowCheckBox = DTS_SHOWNONE
    bpLongDateFormat = DTS_LONGDATEFORMAT
    bpTimeFormat = DTS_TIMEFORMAT
    bpRightAligned = DTS_RIGHTALIGN
    
    'store the monthcal styles in the hiword
    bpMCDivisor = &H10000
    
    bpNoToday = MCS_NOTODAY * bpMCDivisor
    bpNoTodayCircle = MCS_NOTODAYCIRCLE * bpMCDivisor
    bpWeekNumbers = MCS_WEEKNUMBERS * bpMCDivisor
    
    'masks for extracting different values from the properties
    bpValidFormats = bpLongDateFormat Or bpTimeFormat
    bpValidDTPStyles = bpUpDown Or bpShowCheckBox Or bpRightAligned Or bpValidFormats
    bpValidMCStyles = bpNoToday Or bpNoTodayCircle Or bpWeekNumbers
    
End Enum

Private Const PROP_Font                     As String = "Font"
Private Const PROP_CalFont                  As String = "CalFont"
Private Const PROP_BooleanProps             As String = "BooleanProps"
Private Const PROP_FormatString             As String = "FormatString"
Private Const PROP_MaxDate                  As String = "MaxDate"
Private Const PROP_MinDate                  As String = "MinDate"
Private Const PROP_Value                    As String = "Value"
Private Const PROP_Enabled                  As String = "Enabled"

Private Const PROP_ColorText                As String = "TextColor"
Private Const PROP_ColorTitleBackground     As String = "TitleBack"
Private Const PROP_ColorTitleText           As String = "TitleText"
Private Const PROP_ColorBackground          As String = "BackColor"
Private Const PROP_ColorTrailingText        As String = "TrailingText"
Private Const PROP_Themeable                As String = "Themeable"

Private Const DEF_Enabled                   As Boolean = True
Private Const DEF_MaxDate                   As Date = NoDate
Private Const DEF_MinDate                   As Date = NoDate
Private Const DEF_FormatString              As String = vbNullString

Private Const DEF_TrailingBackColor         As Long = vbWindowBackground
Private Const DEF_TextColor                 As Long = vbButtonText
Private Const DEF_TitleBackColor            As Long = vbActiveTitleBar
Private Const DEF_TitleTextColor            As Long = vbTitleBarText
Private Const DEF_Backcolor                 As Long = vbWindowBackground
Private Const DEF_TrailingTextColor         As Long = vbGrayText
Private Const DEF_Themeable                 As Boolean = True

Private Const DEF_BooleanProps              As Long = bpShortDateFormat

Private WithEvents moFont                   As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moCalFont                As cFont
Attribute moCalFont.VB_VarHelpID = -1

Private WithEvents moFontPage               As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private mhWnd                               As Long
Private mhWndCal                            As Long

Private mbNoPropChange                      As Boolean
Private mbThemeable                         As Boolean

Private miBooleanProps                      As eBooleanProps

Private mdMinDate                           As Date
Private mdMaxDate                           As Date
Private mdValue                             As Date
Private miWheelDelta                        As Long
Private mhCalFont                           As Long
Private mhFont                              As Long
Private msFormat                            As String

Private mbInCallback                        As Boolean
Private miColors(MCSC_BACKGROUND To MCSC_TRAILINGTEXT) As OLE_COLOR

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the fonts if they are set to use the ambient font.
'---------------------------------------------------------------------------------------
    If StrComp(PropertyName, "Font") = ZeroL Then
        moFont.OnAmbientFontChanged Ambient.Font
        moCalFont.OnAmbientFontChanged Ambient.Font
    End If
End Sub

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Initialize the common control library load a shell module handle to prevent
'             crashes on CC 6.0 and install vtable subclassing.
'---------------------------------------------------------------------------------------
    LoadShellMod
    InitCC ICC_DATE_CLASSES
    Set moFontPage = New pcSupportFontPropPage
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Initialize property values.
'---------------------------------------------------------------------------------------
    mbNoPropChange = True
    Set moFont = Font_CreateDefault(Ambient.Font)
    Set moCalFont = Font_CreateDefault(Ambient.Font)
    msFormat = DEF_FormatString
    miBooleanProps = DEF_BooleanProps
    pCreate
    mdValue = Now
    msFormat = DEF_FormatString
    mdMinDate = DEF_MinDate
    mdMaxDate = DEF_MaxDate
    miColors(MCSC_BACKGROUND) = DEF_TrailingBackColor
    miColors(MCSC_MONTHBK) = DEF_Backcolor
    miColors(MCSC_TITLEBK) = DEF_TitleBackColor
    miColors(MCSC_TEXT) = DEF_TextColor
    miColors(MCSC_TITLETEXT) = DEF_TitleTextColor
    miColors(MCSC_TRAILINGTEXT) = DEF_TrailingTextColor
    mbThemeable = DEF_Themeable
    mbNoPropChange = False
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Read property values from a previously saved instance.
'---------------------------------------------------------------------------------------
    mbNoPropChange = True
    On Error Resume Next
    
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    Set moCalFont = Font_Read(PropBag, PROP_CalFont, Ambient.Font)
    
    miBooleanProps = PropBag.ReadProperty(PROP_BooleanProps, DEF_BooleanProps)
    
    ColorBackground = PropBag.ReadProperty(PROP_ColorBackground, DEF_Backcolor)
    ColorTitleBackground = PropBag.ReadProperty(PROP_ColorTitleBackground, DEF_TitleBackColor)
    ColorText = PropBag.ReadProperty(PROP_ColorText, DEF_TextColor)
    ColorTitleText = PropBag.ReadProperty(PROP_ColorTitleText, DEF_TitleTextColor)
    ColorTrailingText = PropBag.ReadProperty(PROP_ColorTrailingText, DEF_TrailingTextColor)
    
    Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    mdMaxDate = PropBag.ReadProperty(PROP_MaxDate, DEF_MaxDate)
    mdMinDate = PropBag.ReadProperty(PROP_MinDate, DEF_MinDate)
    mdValue = PropBag.ReadProperty(PROP_Value)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    
    pCreate
    pSetMaxMinDate
    
    FormatString = PropBag.ReadProperty(PROP_FormatString, DEF_FormatString)
    On Error GoTo 0
    mbNoPropChange = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Save property values between instances.
'---------------------------------------------------------------------------------------
    PropBag.WriteProperty PROP_MaxDate, mdMaxDate, DEF_MaxDate
    PropBag.WriteProperty PROP_MinDate, mdMinDate, DEF_MinDate
    PropBag.WriteProperty PROP_FormatString, msFormat, DEF_FormatString
    PropBag.WriteProperty PROP_ColorBackground, miColors(MCSC_MONTHBK), DEF_Backcolor
    PropBag.WriteProperty PROP_ColorTitleBackground, miColors(MCSC_TITLEBK), DEF_TitleBackColor
    PropBag.WriteProperty PROP_ColorText, miColors(MCSC_TEXT), DEF_TextColor
    PropBag.WriteProperty PROP_ColorTitleText, miColors(MCSC_TITLETEXT), DEF_TitleTextColor
    PropBag.WriteProperty PROP_ColorTrailingText, miColors(MCSC_TRAILINGTEXT), DEF_TrailingTextColor
    PropBag.WriteProperty PROP_BooleanProps, miBooleanProps, DEF_BooleanProps
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_Value, mdValue
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    Font_Write moFont, PropBag, PROP_Font
    Font_Write moCalFont, PropBag, PROP_CalFont
End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the position of the dtp.
'---------------------------------------------------------------------------------------
    If mhWnd Then MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
End Sub

Private Sub UserControl_Terminate()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Destroy the dtp, remove vtable subclassing and release the shell module handle.
'---------------------------------------------------------------------------------------
    Set moFontPage = Nothing
    pDestroy
    If mhCalFont Then moCalFont.ReleaseHandle mhCalFont
    If mhFont Then moFont.ReleaseHandle mhFont
    mhCalFont = ZeroL
    mhFont = ZeroL
    ReleaseShellMod
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Forward arrow keys, home, end, pageup, pagedown to the dtp.
'---------------------------------------------------------------------------------------
    Select Case wParam And &HFFFF&
    Case vbKeyEnd To vbKeyDown
        If mhWnd Then SendMessage mhWnd, uMsg, wParam, lParam
        bHandled = True
    End Select
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
    Dim lbOldInCallback As Boolean
    lbOldInCallback = mbInCallback
    mbInCallback = True
    Select Case uMsg
    Case WM_NOTIFY
        pOnNotify lParam
   
   Case WM_SETFOCUS
        ActivateIPAO Me
        
   Case WM_MOUSEACTIVATE
        
        If mhWnd Then
            If GetFocus() <> mhWnd Then
               vbComCtlTlb.SetFocus UserControl.hWnd
               lReturn = MA_NOACTIVATE
               bHandled = True
            End If
        End If
        
    Case WM_MOUSEWHEEL
        'kind of a hack, but it's too easy!  WM_MOUSEWHEEL Always starts at the focus window
        'and works its way up, so since our dtp has no children, it must be in focus when
        'this message is received.  DTP does not provide a way that I know of to tell whether
        'the user has selected the month, day, year, hour, minute, second etc to change that,
        'but by sending up/down the control figures it out on its own.
        If mhWnd Then
            miWheelDelta = miWheelDelta - hiword(wParam)
            If Abs(miWheelDelta) >= 120& Then
                If mhWndCal Then
                    pUpdateMonthCal miWheelDelta
                Else
                    If Sgn(miWheelDelta) = -1& Then
                        SendMessage mhWnd, WM_KEYDOWN, VK_UP, &H1480001
                        SendMessage mhWnd, WM_KEYUP, VK_UP, &HC1480001
                    Else
                        SendMessage mhWnd, WM_KEYDOWN, VK_DOWN, &H1500001
                        SendMessage mhWnd, WM_KEYUP, VK_DOWN, &HC1500001
                    End If
                End If
                miWheelDelta = ZeroL
            End If
            lReturn = ZeroL
            bHandled = True
        End If
    End Select
    mbInCallback = lbOldInCallback

End Sub

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Tell the font property page which font properties we support.
'---------------------------------------------------------------------------------------
    o.ShowProps PROP_Font, PROP_CalFont
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Provide the font property page with our ambient font.
'---------------------------------------------------------------------------------------
    Set o = Ambient.Font
End Sub

Private Sub moCalFont_Changed()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the dropdown month calendar font.
'---------------------------------------------------------------------------------------
    pSetCalFont
    If Not Ambient.UserMode Then pPropChanged PROP_Font
End Sub

Private Sub moFont_Changed()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the dtp font.
'---------------------------------------------------------------------------------------
    pSetFont
    If Not Ambient.UserMode Then pPropChanged PROP_Font
End Sub

Private Sub pOnNotify(ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Handle notification messages from the dtp.
'---------------------------------------------------------------------------------------
    Const NMHDR_code = 8&
    Select Case MemOffset32(lParam, NMHDR_code)
    Case DTN_DATETIMECHANGE
        Dim DTC As NMDATETIMECHANGE
        CopyMemory DTC, ByVal lParam, LenB(DTC)
        Dim ldDate As Date
        If DTC.dwFlags <> GDT_NONE Then SysTimeToDate ldDate, DTC.st
        If ldDate <> mdValue Then
            mdValue = ldDate
            RaiseEvent DateTimeChange
        End If
    Case DTN_CLOSEUP
        RaiseEvent CloseUp
        mhWndCal = ZeroL
    Case DTN_DROPDOWN
        If mhWnd Then
            mhWndCal = SendMessage(mhWnd, DTM_GETMONTHCAL, ZeroL, ZeroL)
            
            If mhWndCal Then
                pSetCalStyle
                EnableWindowTheme mhWndCal, mbThemeable
                RaiseEvent DropDown(mhWndCal)
            End If
            
        End If
'not supported yet:
'        Case DTN_FORMAT
'        Case DTN_FORMATQUERY
'        Case DTN_WMKEYDOWN
    End Select
End Sub

Private Sub pSetCalStyle()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the window style of the drop down monthcal based on the property settings.
'---------------------------------------------------------------------------------------
    If mhWndCal Then
        SetWindowStyle mhWndCal, (miBooleanProps And bpValidMCStyles) \ bpMCDivisor, bpValidMCStyles
        
        Dim ltRectSize As RECT
        SendMessageAny mhWndCal, MCM_GETMINREQRECT, 0, ltRectSize
        
        Dim ltRectPos As RECT
        GetWindowRect mhWndCal, ltRectPos
        
        MoveWindow mhWndCal, ltRectPos.Left, ltRectPos.Top, _
                             ltRectSize.Right - ltRectSize.Left + TwoL, _
                             ltRectSize.bottom - ltRectSize.Top + TwoL, _
                             OneL
    End If
End Sub

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Notify the dtp of a new font handle.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lhFont As Long
        lhFont = moFont.GetHandle
        If lhFont Then
            SendMessage mhWnd, WM_SETFONT, lhFont, NegOneL
            If mhFont Then moFont.ReleaseHandle mhFont
            mhFont = lhFont
        End If
    End If
End Sub

Private Sub pSetTheme()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the window theme for the dtp and updown control if it exists.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        EnableWindowTheme mhWnd, mbThemeable
        If (miBooleanProps And bpUpDown) Then
            Dim lhWnd As Long
            lhWnd = FindWindowExW(mhWnd, ZeroL, vbNullString, vbNullString)
            If lhWnd Then EnableWindowTheme lhWnd, mbThemeable
        End If
    End If
End Sub

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Create the dtp and install the needed subclasses.
'---------------------------------------------------------------------------------------
    If mbInCallback Then Debug.Assert False: Exit Sub
    
    pDestroy
    
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_DATETIMEPICKER & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, pGetDTPStyle(), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        pSetTheme
        
        Value = mdValue
        pSetMaxMinDate
        
        pSetFont
        pSetCalFont
        
        pSetColor MCSC_BACKGROUND
        pSetColor MCSC_TEXT
        pSetColor MCSC_TITLEBK
        pSetColor MCSC_TITLETEXT
        pSetColor MCSC_MONTHBK
        pSetColor MCSC_TRAILINGTEXT
        
        EnableWindow mhWnd, -CLng(UserControl.Enabled)
        FormatString = msFormat
        
        If Ambient.UserMode Then
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_MOUSEACTIVATE), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_MOUSEACTIVATE, WM_SETFOCUS, WM_MOUSEWHEEL), WM_KILLFOCUS
            VTableSubclass_IPAO_Install Me
        End If
        
    End If
        
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Destroy the dtp and subclasses.
'---------------------------------------------------------------------------------------
    
    If mhWnd Then
    
        VTableSubclass_IPAO_Remove
        
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        
        DestroyWindow mhWnd
        mhWnd = ZeroL
        mhWndCal = ZeroL
        
    End If
    
End Sub

Private Sub pSetCalFont()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Tell the dropdown monthcal which font handle to use.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lhFont As Long
        lhFont = moCalFont.GetHandle
        If lhFont Then
            SendMessage mhWnd, DTM_SETMCFONT, lhFont, -1&
            If mhCalFont Then moCalFont.ReleaseHandle mhCalFont
            mhCalFont = lhFont
        End If
    End If
End Sub
Private Sub pSetMaxMinDate()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the date range of the dtp control.
'---------------------------------------------------------------------------------------
    Dim SysTimes(1) As SYSTEMTIME
    If mhWnd Then
        Dim liFlags As Long
        If mdMinDate <> NoDate Then
            DateToSysTime mdMinDate, SysTimes(0)
            liFlags = liFlags Or GDTR_MAX
        End If
        If mdMaxDate <> NoDate Then
            DateToSysTime mdMaxDate, SysTimes(1)
            liFlags = liFlags Or GDTR_MIN
        End If
        SendMessage mhWnd, DTM_SETRANGE, liFlags, VarPtr(SysTimes(0))
    End If
End Sub

Private Function pGetDTPStyle() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Get the window style of the dtp from property values.
'---------------------------------------------------------------------------------------
    pGetDTPStyle = (miBooleanProps And bpValidDTPStyles) Or WS_CHILD Or WS_VISIBLE
End Function

Private Sub pSetColor(ByVal iColor As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Tell the dtp which colors we would like to see on the dropdown monthcal.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, DTM_SETMCCOLOR, iColor, TranslateColor(miColors(iColor))
    End If
End Sub

Private Sub pPropChanged(ByRef sProp As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Call PropertyChanged unless we are updating property values.
'---------------------------------------------------------------------------------------
    If Not mbNoPropChange Then PropertyChanged sProp
End Sub

Private Sub pUpdateMonthCal(ByVal iDelta As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Update the date in the monthcalendar based on a mousewheel notification.
'---------------------------------------------------------------------------------------
    Dim ltSysTime As SYSTEMTIME
    Value = DateAdd("m", Sgn(iDelta), mdValue)
    DateToSysTime mdValue, ltSysTime
    SendMessage mhWndCal, MCM_SETCURSEL, ZeroL, VarPtr(ltSysTime)
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a proxy object to receive notifications from the font property page.
'---------------------------------------------------------------------------------------
    Set fSupportFontPropPage = moFontPage
End Property


Public Property Get ColorText() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color used for regular day text in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    ColorText = miColors(MCSC_TEXT)
End Property
Public Property Let ColorText(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color used for regular day text in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TEXT) = iNew
    If Not Ambient.UserMode Then pPropChanged PROP_ColorText
    pSetColor MCSC_TEXT
End Property

Public Property Get ColorTitleBackground() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color used for the title background in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    ColorTitleBackground = miColors(MCSC_TITLEBK)
End Property
Public Property Let ColorTitleBackground(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color used for the title background in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TITLEBK) = iNew
    If Not Ambient.UserMode Then pPropChanged PROP_ColorTitleBackground
    pSetColor MCSC_TITLEBK
End Property

Public Property Get ColorTitleText() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color used for the title text in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    ColorTitleText = miColors(MCSC_TITLETEXT)
End Property
Public Property Let ColorTitleText(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color used for the title text in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TITLETEXT) = iNew
    If Not Ambient.UserMode Then pPropChanged PROP_ColorTitleText
    pSetColor MCSC_TITLETEXT
End Property

Public Property Get ColorBackground() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color used for the background in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    ColorBackground = miColors(MCSC_MONTHBK)
End Property
Public Property Let ColorBackground(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color used for the background in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_MONTHBK) = iNew
    If Not Ambient.UserMode Then pPropChanged PROP_ColorBackground
    pSetColor MCSC_MONTHBK
End Property

Public Property Get ColorTrailingText() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color used for the days of next/previous months in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    ColorTrailingText = miColors(MCSC_TRAILINGTEXT)
End Property
Public Property Let ColorTrailingText(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color used for the days of next/previous months in the dropdown monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TRAILINGTEXT) = iNew
    If Not Ambient.UserMode Then pPropChanged PROP_ColorTrailingText
    pSetColor MCSC_TRAILINGTEXT
End Property

Public Property Let Enabled(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set enabled status of the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then EnableWindow mhWnd, -CLng(bNew)
    UserControl.Enabled = bNew
    If Not Ambient.UserMode Then pPropChanged PROP_Enabled
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the enabled status of the control.
'---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property

Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the font used by the control.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    pSetFont
    If Not Ambient.UserMode Then pPropChanged PROP_Font
End Property

Public Property Get Font() As cFont
Attribute Font.VB_ProcData.VB_Invoke_Property = "ppFont"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the font used by the control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property

Public Sub Refresh()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Redraw the dtp control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        RedrawWindow mhWnd, ByVal ZeroL, ZeroL, RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_INVALIDATE Or RDW_ERASE Or RDW_FRAME
    End If
End Sub

Public Property Let ShowWeekNumbers(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether week numbers appear in the drop down monthcal.
'---------------------------------------------------------------------------------------
    If bNew _
        Then miBooleanProps = miBooleanProps Or bpWeekNumbers _
        Else miBooleanProps = miBooleanProps And Not bpWeekNumbers
    
    pSetCalStyle
    If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
    
End Property
Public Property Get ShowWeekNumbers() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether week numbers appear in the drop down monthcal.
'---------------------------------------------------------------------------------------
    ShowWeekNumbers = CBool(miBooleanProps And bpWeekNumbers)
End Property

Public Property Let ShowToday(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether the date is shown at the bottom of the drop down monthcal.
'---------------------------------------------------------------------------------------
    If bNew _
        Then miBooleanProps = miBooleanProps And Not bpNoToday _
        Else miBooleanProps = miBooleanProps Or bpNoToday
    
    pSetCalStyle
    If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
    
End Property
Public Property Get ShowToday() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether the date is shown at the bottom of the drop down monthcal.
'---------------------------------------------------------------------------------------
    ShowToday = Not CBool(miBooleanProps And bpNoToday)
End Property

Public Property Let ShowTodayCircle(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether the current date is circled on the drop down monthcal.
'---------------------------------------------------------------------------------------
    If bNew _
        Then miBooleanProps = miBooleanProps And Not bpNoTodayCircle _
        Else miBooleanProps = miBooleanProps Or bpNoTodayCircle
    
    pSetCalStyle
    If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
    
End Property
Public Property Get ShowTodayCircle() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether the current date is circled on the drop down monthcal.
'---------------------------------------------------------------------------------------
    ShowTodayCircle = Not CBool(miBooleanProps And bpNoTodayCircle)
End Property

Public Property Set CalFont(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the font used by the drop down monthcal.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moCalFont = Font_CreateDefault(Ambient.Font) _
        Else Set moCalFont = oNew

    pSetCalFont
    If Not Ambient.UserMode Then pPropChanged PROP_CalFont
End Property

Public Property Get CalFont() As cFont
Attribute CalFont.VB_ProcData.VB_Invoke_Property = "ppFont"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the font used by the drop down monthcal.
'---------------------------------------------------------------------------------------
    Set CalFont = moCalFont
End Property


Public Property Let Format(ByVal iNew As eDateTimePickerFormats)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Tell the dtp to use a standard format.
'---------------------------------------------------------------------------------------
    If iNew = dtpLongDate Or iNew = dtpShortDate Or iNew = dtpTime Then
        miBooleanProps = (miBooleanProps And Not bpValidFormats) Or iNew
        msFormat = vbNullString
        pCreate
        If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
    End If
End Property

Public Property Get Format() As eDateTimePickerFormats
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a standard format in use by the control.
'---------------------------------------------------------------------------------------
    Format = miBooleanProps And bpValidFormats
End Property

Public Property Get FormatString() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a custom format in use by the control.
'---------------------------------------------------------------------------------------
    FormatString = msFormat
End Property

Public Property Let FormatString(ByVal NewS As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set a custom format in use by the control.
'---------------------------------------------------------------------------------------
    msFormat = NewS
    If mhWnd Then
        If LenB(NewS) = ZeroL Then
            SendMessage mhWnd, DTM_SETFORMATA, ZeroL, ZeroL
        Else
            Dim ls As String
            ls = StrConv(NewS, vbFromUnicode)
            SendMessage mhWnd, DTM_SETFORMATA, ZeroL, StrPtr(ls)
        End If
    End If
    If Not Ambient.UserMode Then pPropChanged PROP_FormatString
End Property

Public Property Let MaxDate(ByVal NewV As Date)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the maximum date allowable in the dtp.
'---------------------------------------------------------------------------------------
    mdMaxDate = NewV
    pSetMaxMinDate
    If Not Ambient.UserMode Then pPropChanged PROP_MaxDate
End Property
Public Property Get MaxDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the maximum date allowable in the dtp.
'---------------------------------------------------------------------------------------
    MaxDate = mdMaxDate
End Property

Public Property Let MinDate(ByVal NewV As Date)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minimum date allowable in the dtp.
'---------------------------------------------------------------------------------------
    mdMinDate = NewV
    pSetMaxMinDate
    If Not Ambient.UserMode Then pPropChanged PROP_MinDate
End Property
Public Property Get MinDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minimum date allowable in the dtp.
'---------------------------------------------------------------------------------------
    MinDate = mdMinDate
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndDTP() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the hwnd of the dtp control.
'---------------------------------------------------------------------------------------
    If mhWnd Then hWndDTP = mhWnd
End Property

Public Property Let RightAlignedCal(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether the drop down monthcal appears at the right edge of the dtp.
'---------------------------------------------------------------------------------------
    If bNew _
        Then miBooleanProps = miBooleanProps Or bpRightAligned _
        Else miBooleanProps = miBooleanProps And Not bpRightAligned
    
    If mhWnd Then SetWindowLong mhWnd, GWL_STYLE, pGetDTPStyle
    If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
End Property

Public Property Get RightAlignedCal() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether the drop down monthcal appears at the right edge of the dtp.
'---------------------------------------------------------------------------------------
    RightAlignedCal = CBool(miBooleanProps And bpRightAligned)
End Property

Public Property Let Value(ByVal dNew As Date)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the date value displayed by the dtp.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim ltSysTime As SYSTEMTIME
        Dim liValid As Long
        
        liValid = SendMessage(mhWnd, DTM_GETSYSTEMTIME, ZeroL, VarPtr(ltSysTime))
        DateToSysTime dNew, ltSysTime
        
        If SendMessage(mhWnd, DTM_SETSYSTEMTIME, liValid, VarPtr(ltSysTime)) Then mdValue = dNew
        If Not Ambient.UserMode Then pPropChanged PROP_Value
        
    End If
    
End Property
Public Property Get Value() As Date
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the date value displayed by the dtp.
'---------------------------------------------------------------------------------------
    Value = mdValue
End Property

Public Property Let ShowCheckBox(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether a checkbox is displayed on the left edge of the dtp.
'---------------------------------------------------------------------------------------
    If bNew _
        Then miBooleanProps = miBooleanProps Or bpShowCheckBox _
        Else miBooleanProps = miBooleanProps And Not bpShowCheckBox
    pCreate
    If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
End Property

Public Property Get ShowCheckBox() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether a checkbox is displayed on the left edge of the dtp.
'---------------------------------------------------------------------------------------
    ShowCheckBox = CBool(miBooleanProps And bpShowCheckBox)
End Property

Public Property Let UpDown(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether the dtp uses an updown control or a dropdown button.
'---------------------------------------------------------------------------------------
    If bNew _
        Then miBooleanProps = miBooleanProps Or bpUpDown _
        Else miBooleanProps = miBooleanProps And Not bpUpDown
    pCreate
    If Not Ambient.UserMode Then pPropChanged PROP_BooleanProps
End Property
Public Property Get UpDown() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether the dtp uses an updown control or a dropdown button.
'---------------------------------------------------------------------------------------
    UpDown = CBool(miBooleanProps And bpUpDown)
End Property

Public Property Let Checked(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether the checkbox (if any) is checked.
'---------------------------------------------------------------------------------------
    If ShowCheckBox And CBool(mhWnd) Then
        Dim ltSysTime As SYSTEMTIME
        If bNew Then
            DateToSysTime mdValue, ltSysTime
            SendMessage mhWnd, DTM_SETSYSTEMTIME, GDT_VALID, VarPtr(ltSysTime)
        Else
            SendMessage mhWnd, DTM_SETSYSTEMTIME, GDT_NONE, ZeroL
        End If
    End If
End Property
Public Property Get Checked() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return whether the checkbox (if any) is checked.
'---------------------------------------------------------------------------------------
    If mhWnd Then
    
        Dim ltSysTime As SYSTEMTIME
        Checked = SendMessage(mhWnd, DTM_GETSYSTEMTIME, ZeroL, VarPtr(ltSysTime)) = GDT_VALID

    End If
    
End Property

Public Property Get Themeable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a value indicating whether a default theme should be used if available.
'---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property
Public Property Let Themeable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set whether a default theme should be used if available.
'---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        mbThemeable = bNew
        pSetTheme
        If Not Ambient.UserMode Then pPropChanged PROP_Themeable
    End If
End Property
