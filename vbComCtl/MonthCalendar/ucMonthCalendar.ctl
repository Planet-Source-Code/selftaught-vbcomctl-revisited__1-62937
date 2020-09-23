VERSION 5.00
Begin VB.UserControl ucMonthCalendar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   PropertyPages   =   "ucMonthCalendar.ctx":0000
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   ToolboxBitmap   =   "ucMonthCalendar.ctx":001D
End
Attribute VB_Name = "ucMonthCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucMonthCalendar.ctl        12/15/04
'
'           PURPOSE:
'               Implement the Win32 SysMonthView32.
'
'           LINEAGE:
'               http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=9549&lngWId=1
'               MonthCal.ctl
'
'==================================================================================================

Option Explicit

Private Const DEF_ShowToday                 As Boolean = True
Private Const DEF_ShowTodayCircle           As Boolean = True
Private Const DEF_ShowWeekNumbers           As Boolean = False
Private Const DEF_MultiSelect               As Boolean = False
Private Const DEF_Enabled                   As Boolean = True
Private Const DEF_Max                       As Date = NoDate
Private Const DEF_Min                       As Date = NoDate
Private Const DEF_MaxSel                    As Long = 14
Private Const DEF_BorderStyle               As Long = vbccBorderSunken
Private Const DEF_ColorTrailingBack         As Long = vbWindowBackground
Private Const DEF_ColorBack                 As Long = vbWindowBackground
Private Const DEF_ColorTitleBack            As Long = vbActiveTitleBar
Private Const DEF_ColorText                 As Long = vbButtonText
Private Const DEF_ColorTitleText            As Long = vbTitleBarText
Private Const DEF_ColorTrailingText         As Long = vbGrayText
Private Const DEF_Themeable                 As Boolean = True

Private Const PROP_ShowToday                As String = "ShowToday"
Private Const PROP_ShowTodayCircle          As String = "ShowCircle"
Private Const PROP_ShowWeekNumbers          As String = "ShowWeekNums"
Private Const PROP_MultiSelect              As String = "MultiSelect"
Private Const PROP_Enabled                  As String = "Enabled"
Private Const PROP_Max                      As String = "MaxDate"
Private Const PROP_Min                      As String = "MinDate"
Private Const PROP_MaxSel                   As String = "MaxSel"
Private Const PROP_BorderStyle              As String = "BorderStyle"
Private Const PROP_ColorTrailingBack        As String = "TrailingBack"
Private Const PROP_ColorBack                As String = "Back"
Private Const PROP_ColorTitleBack           As String = "TitleBack"
Private Const PROP_ColorText                As String = "Text"
Private Const PROP_ColorTitleText           As String = "TitleText"
Private Const PROP_ColorTrailingText        As String = "TrailingText"
Private Const PROP_Font                     As String = "Font"
Private Const PROP_Themeable                As String = "Themeable"

'Subclassing:
Implements iSubclass
Implements iOleInPlaceActiveObjectVB

'Private variables:

Private WithEvents moFont   As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private mhWnd               As Long

Private mbShowToday         As Boolean
Private mbShowTodayCircle   As Boolean
Private mbShowWeekNumbers   As Boolean
Private mbMultiSelect       As Boolean
Private mbThemeable         As Boolean

Private mhFont              As Long
Private miWheelDelta        As Long
Private miMaxSel            As Long
Private miBorderStyle       As evbComCtlBorderStyle

Private miColors(MCSC_BACKGROUND To MCSC_TRAILINGTEXT) As OLE_COLOR

Private mdMax               As Date
Private mdMin               As Date

Private mbInFocus           As Boolean
Private mbInCallback        As Boolean
Private mtSysTimes(0 To 1)  As SYSTEMTIME

Private miDayState()        As Long

Event Change()
Event GetBoldDays(ByVal iMonth As Long, ByVal iYear As Long, ByRef iMask As Long)

Public Property Let Enabled(ByVal bNew As Boolean)
Attribute Enabled.VB_UserMemId = -514
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the enabled status of the control
'---------------------------------------------------------------------------------------
    UserControl.Enabled = bNew
    If Not Ambient.UserMode Then PropertyChanged PROP_Enabled
    
    'SetWindowStyle mhWnd,  WS_DISABLED * (bNew + OneL), WS_DISABLED
    pRecreate
End Property
Public Property Get Enabled() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the enabled status of the control
'---------------------------------------------------------------------------------------
    Enabled = UserControl.Enabled
End Property

Public Property Let MultiSelect(ByVal bVal As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set whether day ranges can be selected.
'             This is painted incorrectly under comctl version 6
'---------------------------------------------------------------------------------------
    mbMultiSelect = bVal
    If Not Ambient.UserMode Then PropertyChanged PROP_MultiSelect
    pRecreate
End Property
Public Property Get MultiSelect() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return whether day ranges can be selected.
'             This is painted incorrectly under comctl version 6
'---------------------------------------------------------------------------------------
    MultiSelect = mbMultiSelect
End Property

Public Property Get BorderStyle() As evbComCtlBorderStyle
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the border style of the control.
'---------------------------------------------------------------------------------------
    BorderStyle = miBorderStyle
End Property
Public Property Let BorderStyle(ByVal iNew As evbComCtlBorderStyle)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the border style of the control.
'---------------------------------------------------------------------------------------
    miBorderStyle = (iNew And 3&)
    If Not Ambient.UserMode Then PropertyChanged PROP_BorderStyle
    pRecreate
End Property

Public Property Get MaxSelCount() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the maximum number of days which may be selected if
'             the MultiSelect property is True.
'---------------------------------------------------------------------------------------
    MaxSelCount = miMaxSel
End Property
Public Property Let MaxSelCount(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Sets the maximum number of days which may be selected if
'             the MultiSelect property is True.
'---------------------------------------------------------------------------------------
    miMaxSel = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_MaxSel
    If mhWnd Then
        SendMessage mhWnd, MCM_SETMAXSELCOUNT, iNew, ZeroL
    End If
End Property

Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the font used by this control.
'---------------------------------------------------------------------------------------
    
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    pSetFont
    If Not Ambient.UserMode Then PropertyChanged PROP_Font
End Property
Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Get the font used by this control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property

Public Property Get ColorTrailingBack() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Get the color for the background outside of the current month.
'---------------------------------------------------------------------------------------
    ColorTrailingBack = miColors(MCSC_BACKGROUND)
End Property
Public Property Let ColorTrailingBack(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the color for the background outside of the current month.
'---------------------------------------------------------------------------------------
    miColors(MCSC_BACKGROUND) = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ColorTrailingBack
    pSetColors
End Property

Public Property Get ColorBackground() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the color used for the background in the monthcal.
'---------------------------------------------------------------------------------------
    ColorBackground = miColors(MCSC_MONTHBK)
End Property
Public Property Let ColorBackground(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the color used for the background in the monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_MONTHBK) = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ColorBack
    pSetColors
End Property

Public Property Get ColorText() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the color used for regular day text in the monthcal.
'---------------------------------------------------------------------------------------
    ColorText = miColors(MCSC_TEXT)
End Property
Public Property Let ColorText(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the color used for regular day text in the monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TEXT) = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ColorText
    pSetColors
End Property

Public Property Get ColorTitleBackground() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the color used for the title background in the monthcal.
'---------------------------------------------------------------------------------------
    ColorTitleBackground = miColors(MCSC_TITLEBK)
End Property
Public Property Let ColorTitleBackground(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the color used for the title background in the monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TITLEBK) = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ColorTitleBack
    pSetColors
End Property

Public Property Get ColorTitleText() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the color used for the title text in the monthcal.
'---------------------------------------------------------------------------------------
    ColorTitleText = miColors(MCSC_TITLETEXT)
End Property
Public Property Let ColorTitleText(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the color used for the title text in the monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TITLETEXT) = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ColorTitleText
    pSetColors
End Property

Public Property Get ColorTrailingText() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the color used for the days of next/previous months in the monthcal.
'---------------------------------------------------------------------------------------
    ColorTrailingText = miColors(MCSC_TRAILINGTEXT)
End Property
Public Property Let ColorTrailingText(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the color used for the days of next/previous months in the monthcal.
'---------------------------------------------------------------------------------------
    miColors(MCSC_TRAILINGTEXT) = iNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ColorTrailingText
    pSetColors
End Property

Public Property Let MinDate(ByVal dNew As Date)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign the minimum date to be allowed in the control.
'---------------------------------------------------------------------------------------
    mdMin = dNew
    If Not Ambient.UserMode Then PropertyChanged PROP_Min
    pSetMinMax
End Property
Public Property Get MinDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the minimum date to be allowed in the control.
'---------------------------------------------------------------------------------------
    MinDate = mdMin
End Property

Public Property Let MaxDate(ByVal dNew As Date)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign the maximum date to be allowed in the control.
'---------------------------------------------------------------------------------------
    mdMax = dNew
    If Not Ambient.UserMode Then PropertyChanged PROP_Max
    pSetMinMax
End Property
Public Property Get MaxDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the maximum date to be allowed in the control.
'---------------------------------------------------------------------------------------
    MaxDate = mdMax
End Property

Public Sub GetBoldDays()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Force the control to update the daystate to show the correct days in bold.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liCount As Long
        liCount = SendMessage(mhWnd, MCM_GETMONTHRANGE, GMR_DAYSTATE, VarPtr(mtSysTimes(0)))
        SendMessage mhWnd, MCM_SETDAYSTATE, liCount, pGetBoldDays(mtSysTimes(0).wMonth, mtSysTimes(0).wYear, liCount)
    End If
End Sub

Public Property Let SelDate(ByVal NewV As Date)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign the currently selected date.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        DateToSysTime NewV, mtSysTimes(0)
        SendMessage mhWnd, MCM_SETCURSEL, ZeroL, VarPtr(mtSysTimes(0))
        RaiseEvent Change
    End If
End Property

Public Property Get SelDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the currently selected date.
'---------------------------------------------------------------------------------------
    Dim SysTime As SYSTEMTIME
    
    If mhWnd Then
        If Me.MultiSelect Then
            SelDate = SelMinDate
        Else
            SendMessage mhWnd, MCM_GETCURSEL, ZeroL, VarPtr(SysTime)
            With SysTime
                SelDate = DateSerial(.wYear, .wMonth, .wDay)
            End With
        End If
    End If
End Property

Public Property Get SelMinDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the first day selected if there is a selection range, or the
'             currently selected date otherwise.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If Not Me.MultiSelect Then
            SelMinDate = SelDate
        Else
            SendMessage mhWnd, MCM_GETSELRANGE, ZeroL, VarPtr(mtSysTimes(0))
            SysTimeToDate SelMinDate, mtSysTimes(0)
        End If
    End If
End Property
Public Property Let SelMinDate(ByVal dDate As Date)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the first day selected, keeping the last day the same if possible while
'             enforcing the maximum sellength.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        pSetSelection dDate
    End If
End Property

Public Property Get SelMaxDate() As Date
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the last day selected if there is a selection range, or the
'             currently selected date otherwise.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If Not Me.MultiSelect Then
            SelMaxDate = SelDate
        Else
            SendMessage mhWnd, MCM_GETSELRANGE, ZeroL, VarPtr(mtSysTimes(0))
            SysTimeToDate SelMaxDate, mtSysTimes(1)
        End If
    End If
End Property
Public Property Let SelMaxDate(ByVal dDate As Date)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the last day selected, keeping the first day the same if possible while
'             enforcing the maximum sellength.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        pSetSelection , dDate
    End If
End Property

Private Sub pSetSelection(Optional ByVal d1 As Date = NoDate, Optional ByVal d2 As Date = NoDate)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set the new min/max dates while enforcing the maximum selection length.
'---------------------------------------------------------------------------------------
    
    Dim lb1 As Boolean
    If d1 = NoDate Then
        d1 = SelMinDate
        lb1 = True
    End If
    If d2 = NoDate Then d2 = SelMaxDate

    Select Case DateDiff("d", d1, d2)
    Case Is >= miMaxSel
        d2 = DateAdd("d", miMaxSel - OneL, d1)
    Case Is < 0
        If lb1 Then
            d1 = d2
        Else
            d2 = d1
        End If
    End Select
    DateToSysTime d1, mtSysTimes(0)
    DateToSysTime d2, mtSysTimes(1)
    If mhWnd Then
        SendMessage mhWnd, MCM_SETSELRANGE, ZeroL, VarPtr(mtSysTimes(0))
        RaiseEvent Change
    End If
End Sub

Public Property Let ShowToday(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign whether the date is shown at the bottom of the calendar.
'---------------------------------------------------------------------------------------
    
    mbShowToday = bNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ShowToday
    If mhWnd Then
        If bNew Then
            SetWindowStyle mhWnd, ZeroL, MCS_NOTODAY
        Else
            SetWindowStyle mhWnd, MCS_NOTODAY, ZeroL
        End If
    End If
    'pAutoSize
End Property
Public Property Get ShowToday() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return whether the date is shown at the bottom of the calendar.
'---------------------------------------------------------------------------------------
    ShowToday = mbShowToday
End Property

Public Property Let ShowTodayCircle(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign whether today is circled on the calendar.
'---------------------------------------------------------------------------------------
    mbShowTodayCircle = bNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ShowTodayCircle
    If mhWnd Then
        If bNew Then
            SetWindowStyle mhWnd, ZeroL, MCS_NOTODAYCIRCLE
        Else
            SetWindowStyle mhWnd, MCS_NOTODAYCIRCLE, ZeroL
        End If
    End If
    'pAutoSize
End Property
Public Property Get ShowTodayCircle() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set whether today is circled on the calendar.
'---------------------------------------------------------------------------------------
    ShowTodayCircle = mbShowTodayCircle
End Property

Public Property Let ShowWeekNumbers(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign whether Week numbers are shown on the side of the calendar.
'---------------------------------------------------------------------------------------
    mbShowWeekNumbers = bNew
    If Not Ambient.UserMode Then PropertyChanged PROP_ShowWeekNumbers
    If mhWnd Then
        If bNew Then
            SetWindowStyle mhWnd, MCS_WEEKNUMBERS, ZeroL
        Else
            SetWindowStyle mhWnd, ZeroL, MCS_WEEKNUMBERS
        End If
    End If
    'pAutoSize
End Property

Public Property Get ShowWeekNumbers() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return whether Week numbers are shown on the side of the calendar.
'---------------------------------------------------------------------------------------
    ShowWeekNumbers = mbShowWeekNumbers
End Property

Public Property Get MinReqWidth() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the minimum width for a month in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim ltR As RECT
        SendMessage mhWnd, MCM_GETMINREQRECT, ZeroL, VarPtr(ltR)
        MinReqWidth = ScaleX(ltR.Right + pBorderWidth, vbPixels, vbContainerSize)
    End If
End Property
Public Property Get MinReqHeight() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the minimum height for a month in the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim ltR As RECT
        SendMessage mhWnd, MCM_GETMINREQRECT, ZeroL, VarPtr(ltR)
        MinReqHeight = ScaleY(ltR.bottom + pBorderWidth, vbPixels, vbContainerSize)
    End If
End Property

Private Property Get pBorderWidth() As Long
    Select Case miBorderStyle
    Case vbccBorderSingle
        pBorderWidth = 3&
    Case vbccBorderThin
        pBorderWidth = 2&
    Case vbccBorderSunken
        pBorderWidth = 5&
    Case Else
        pBorderWidth = 1&
    End Select
End Property

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign our font handle to the monthcal.
'---------------------------------------------------------------------------------------
    Dim NewhFont As Long
    If mhWnd Then
        NewhFont = moFont.GetHandle
        
        SendMessage mhWnd, WM_SETFONT, NewhFont, NegOneL

        If mhFont Then moFont.ReleaseHandle mhFont
        mhFont = NewhFont
        'pAutoSize
    End If
End Sub

Private Sub pSetColors()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Assign the selected colors to the monthcal.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim L As Long
        For L = 0 To 5
            SendMessage mhWnd, MCM_SETCOLOR, L, TranslateColor(miColors(L))
        Next
        UserControl.BackColor = miColors(MCSC_BACKGROUND)
    End If
End Sub

Public Property Get hWndCalendar() As Long
Attribute hWndCalendar.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the hwnd of the monthcal.
'---------------------------------------------------------------------------------------
    hWndCalendar = mhWnd
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Create the monthcal window with the selected style and install the
'             needed subclasses.
'---------------------------------------------------------------------------------------
    Dim liStyle As Long
    If mbInCallback Then Exit Sub
    
    'Destroy old Calendar
    pDestroy
    
    'Get window style
    If mbShowWeekNumbers Then liStyle = liStyle Or MCS_WEEKNUMBERS
    If Not mbShowToday Then liStyle = liStyle Or MCS_NOTODAY
    If Not mbShowTodayCircle Then liStyle = liStyle Or MCS_NOTODAYCIRCLE
    If mbMultiSelect Then liStyle = liStyle Or MCS_MULTISELECT
    If UserControl.Enabled = False Then liStyle = liStyle Or WS_DISABLED
    liStyle = liStyle Or WS_CHILD Or (MCS_DAYSTATE * Abs(Ambient.UserMode)) Or WS_VISIBLE Or WS_CLIPSIBLINGS Or (Abs(CBool(miBorderStyle = vbccBorderSingle)) * WS_BORDER)
    'Create Win:
    
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_MONTHCAL & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx((WS_EX_CLIENTEDGE * Abs(CBool(miBorderStyle = vbccBorderSunken))) Or (WS_EX_STATICEDGE * Abs(CBool(miBorderStyle = vbccBorderThin))), StrPtr(lsAnsi), ZeroL, liStyle, ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        If Ambient.UserMode Then
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_MOUSEACTIVATE), WM_SETFOCUS
            Subclass_Install Me, mhWnd, Array(WM_MOUSEACTIVATE, WM_SETFOCUS, WM_MOUSEWHEEL, WM_COMMAND), Array(WM_KILLFOCUS, WM_PAINT)
            VTableSubclass_IPAO_Install Me
        End If
        
        EnableWindowTheme mhWnd, mbThemeable
    
        'If CheckCCVersion(6&) Then
            
            '$#@!#$%@!!!!
            
            'SetWindowTheme moWnd.hwnd, StrPtr(" "), StrPtr(" ")
            'SendMessage mhWnd, CCM_SETVERSION, 5&, ZeroL
            'SendMessage mhWnd, CCM_SETWINDOWTHEME, ZeroL, StrPtr(" ")
            
        'End If
        
        pSetMinMax
        pSetColors
        pSetFont
        EnableWindow mhWnd, -CLng(UserControl.Enabled)
        
        If miMaxSel <> ZeroL Then SendMessage mhWnd, MCM_SETMAXSELCOUNT, miMaxSel, ZeroL
        'pAutoSize
    End If
    
End Sub

Private Sub pRecreate()
    Dim ldDateMin As Date: ldDateMin = SelMinDate
    Dim ldDateMax As Date: ldDateMax = SelMaxDate
    
    pCreate
    
    DateToSysTime ldDateMin, mtSysTimes(0)
    DateToSysTime ldDateMax, mtSysTimes(1)
    If mhWnd Then
        SendMessage mhWnd, MCM_SETCURSEL, ZeroL, VarPtr(mtSysTimes(0))
        SendMessage mhWnd, MCM_SETSELRANGE, ZeroL, VarPtr(mtSysTimes(0))
    End If
    GetBoldDays
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Destroy the monthcal and subclasses.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    If mhWnd Then
        VTableSubclass_IPAO_Remove
        Subclass_Remove Me, mhWnd
        Subclass_Remove Me, UserControl.hWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
    On Error GoTo 0
End Sub


Private Sub pSetMinMax()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Inform the monthcal of the minimum and maximum dates that have been chosen.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liFlags As Long
        If mdMin <> NoDate Then
            DateToSysTime mdMin, mtSysTimes(0)
            liFlags = liFlags Or GDTR_MAX
        End If
        If mdMax <> NoDate Then
            DateToSysTime mdMax, mtSysTimes(1)
            liFlags = liFlags Or GDTR_MIN
        End If
        SendMessage mhWnd, MCM_SETRANGE, liFlags, VarPtr(mtSysTimes(0))
    End If
End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, ByVal iShift As evbComCtlKeyboardState, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Forward arrow keys, pgup, pgdn, return, home, end to the monthcal.
'---------------------------------------------------------------------------------------
   If uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Then
      Select Case wParam And &HFFFF&
      Case vbKeyPageUp To vbKeyDown, vbKeyReturn
            If mhWnd Then
                SendMessage mhWnd, uMsg, wParam, lParam
                bHandled = True
            End If
      End Select
   End If
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
    Case WM_SETFOCUS
        vbComCtlTlb.SetFocus mhWnd
    Case WM_KILLFOCUS
        SendMessage mhWnd, WM_CANCELMODE, ZeroL, ZeroL
        mbInFocus = False
        pPaintFocusRect
        UpdateWindow mhWnd
        DeActivateIPAO Me
    Case WM_PAINT
        pPaintFocusRect
    End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Handled focus and notifications from the monthcal.
'---------------------------------------------------------------------------------------
    Dim lbOldInCallback As Boolean
    lbOldInCallback = mbInCallback
    mbInCallback = True
    Const NMHDR_code As Long = 8&
    Const NMSELCHANGE_stStart = 12&
    On Error Resume Next
    'Process messages:
    Select Case uMsg
    Case WM_NOTIFY 'Notify Msg:
        bHandled = True
        Select Case MemOffset32(lParam, NMHDR_code)
        Case MCN_SELCHANGE
            CopyMemory mtSysTimes(0), ByVal UnsignedAdd(lParam, NMSELCHANGE_stStart), Len(mtSysTimes(0)) * TwoL
            pCheckDates
            pChange
        Case MCN_GETDAYSTATE
            Const NMDAYSTATE_stStart As Long = 12
            Const NMDAYSTATE_cDayState As Long = 28
            Const NMDAYSTATE_prgDayState As Long = 32&
            MemOffset32(lParam, NMDAYSTATE_prgDayState) = pGetBoldDays(hiword(MemOffset32(lParam, NMDAYSTATE_stStart)), loword(MemOffset32(lParam, NMDAYSTATE_stStart)), MemOffset32(lParam, NMDAYSTATE_cDayState))
        End Select
    Case WM_COMMAND
        If hiword(wParam) = EN_SETFOCUS Then
            EnableWindowTheme lParam, mbThemeable
            EnableWindowTheme FindWindowExW(mhWnd, ZeroL, "msctls_updown32", vbNullString), mbThemeable
        End If
    Case WM_SETFOCUS
        ActivateIPAO Me
        mbInFocus = True
        pPaintFocusRect
        UpdateWindow mhWnd
    Case WM_MOUSEACTIVATE
        If GetFocus() <> mhWnd Then
            bHandled = True
            vbComCtlTlb.SetFocus UserControl.hWnd
            lReturn = MA_NOACTIVATE
        End If
    Case WM_MOUSEWHEEL
        miWheelDelta = miWheelDelta - hiword(wParam)
        If Abs(miWheelDelta) >= 120& Then
            Me.SelDate = DateAdd("m", Sgn(miWheelDelta), Me.SelDate)
            miWheelDelta = ZeroL
        End If
        lReturn = ZeroL
        bHandled = True
    End Select
    mbInCallback = lbOldInCallback
End Sub

Private Sub pCheckDates()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Enforce the maximum selected number of days.
'---------------------------------------------------------------------------------------
    Dim ld1 As Date
    Dim ld2 As Date
    
    Dim liMaxSel As Long
    
    If miMaxSel = ZeroL Then liMaxSel = 7& Else liMaxSel = miMaxSel
    
    SysTimeToDate ld1, mtSysTimes(0)
    SysTimeToDate ld2, mtSysTimes(1)
    
    If DateDiff("d", ld1, ld2) > liMaxSel Then
        ld2 = DateAdd("d", liMaxSel - OneL, ld1)
        DateToSysTime ld2, mtSysTimes(1)
        If mhWnd Then
            SendMessage mhWnd, MCM_SETSELRANGE, ZeroL, VarPtr(mtSysTimes(0))
        End If
    End If
End Sub

Private Sub pChange()
    Static ldLastDate As Date
    Dim ldDate As Date: ldDate = SelDate
    If ldLastDate <> ldDate Then RaiseEvent Change
    ldLastDate = ldDate
End Sub

Private Function pGetBoldDays(ByVal iMonth As Long, ByVal iYear As Long, ByVal iCount As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Raise an event to get the 32-bit mask of bold days for each month.
'---------------------------------------------------------------------------------------
    ReDim miDayState(0 To iCount - OneL)
    
    For iCount = 0 To iCount - OneL
        RaiseEvent GetBoldDays(iMonth, iYear, miDayState(iCount))
        iMonth = iMonth + OneL
        If iMonth = 13& Then
            iMonth = OneL
            iYear = iYear + OneL
        End If
    Next
    pGetBoldDays = VarPtr(miDayState(0))
End Function

Private Sub pPaintFocusRect()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Paint over any existing focusrect and paint a focus rect if the control
'             has the focus.
'---------------------------------------------------------------------------------------
    Dim lhDc As Long
    Dim lR As RECT
    Dim tP As POINT
    Dim tPJunk As POINT
    Dim lhPen As Long
    Dim lhPenOld As Long
        
    If mhWnd Then
        lhDc = GetDC(mhWnd)
        If lhDc Then
            lhPen = GdiMgr_CreatePen(PS_SOLID, 1, TranslateColor(vbButtonFace))
            If lhPen Then
                lhPenOld = SelectObject(lhDc, lhPen)
                If lhPenOld Then
                    GetClientRect mhWnd, lR
                    MoveToEx lhDc, ZeroL, ZeroL, tPJunk
                    tP.x = lR.Right
                    LineTo lhDc, tP.x - OneL, tP.y
                    tP.y = lR.bottom
                    LineTo lhDc, tP.x - OneL, tP.y - OneL
                    tP.x = ZeroL
                    LineTo lhDc, tP.x, tP.y - OneL
                    tP.y = ZeroL
                    LineTo lhDc, tP.x, tP.y
                    
                    If mbInFocus Then
                        DrawFocusRect lhDc, lR
                    End If
                    
                    lhPenOld = SelectObject(lhDc, lhPenOld)
                End If
                GdiMgr_DeletePen lhPen
            End If
            ReleaseDC mhWnd, lhDc
        End If
    End If
    
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub moFont_Changed()
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    If Not Ambient.UserMode Then PropertyChanged PROP_Font
End Sub

Public Sub Refresh()
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Redraw the control.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        RedrawWindow mhWnd, ByVal ZeroL, ZeroL, RDW_UPDATENOW Or RDW_ALLCHILDREN Or RDW_INVALIDATE Or RDW_ERASE Or RDW_FRAME
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp("Font", PropertyName) = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    LoadShellMod
    InitCC ICC_DATE_CLASSES
    Set moFontPage = New pcSupportFontPropPage
End Sub

Private Sub UserControl_InitProperties()
    miColors(MCSC_BACKGROUND) = DEF_ColorTrailingBack
    miColors(MCSC_MONTHBK) = DEF_ColorBack
    miColors(MCSC_TITLEBK) = DEF_ColorTitleBack
    miColors(MCSC_TEXT) = DEF_ColorText
    miColors(MCSC_TITLETEXT) = DEF_ColorTitleText
    miColors(MCSC_TRAILINGTEXT) = DEF_ColorTrailingText
    mbShowWeekNumbers = DEF_ShowWeekNumbers
    mbShowToday = DEF_ShowToday
    mbShowTodayCircle = DEF_ShowTodayCircle
    mbMultiSelect = DEF_MultiSelect
    mdMin = DEF_Min
    mdMax = DEF_Max
    miMaxSel = DEF_MaxSel
    miBorderStyle = DEF_BorderStyle
    Set moFont = Font_CreateDefault(Ambient.Font)
    UserControl.Enabled = DEF_Enabled
    mbThemeable = DEF_Themeable
    pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    miColors(MCSC_BACKGROUND) = PropBag.ReadProperty(PROP_ColorTrailingBack, DEF_ColorTrailingBack)
    miColors(MCSC_MONTHBK) = PropBag.ReadProperty(PROP_ColorBack, DEF_ColorBack)
    miColors(MCSC_TITLEBK) = PropBag.ReadProperty(PROP_ColorTitleBack, DEF_ColorTitleBack)
    miColors(MCSC_TEXT) = PropBag.ReadProperty(PROP_ColorText, DEF_ColorText)
    miColors(MCSC_TITLETEXT) = PropBag.ReadProperty(PROP_ColorTitleText, DEF_ColorTitleText)
    miColors(MCSC_TRAILINGTEXT) = PropBag.ReadProperty(PROP_ColorTrailingText, DEF_ColorTrailingText)
    mbShowWeekNumbers = PropBag.ReadProperty(PROP_ShowWeekNumbers, DEF_ShowWeekNumbers)
    mbShowToday = PropBag.ReadProperty(PROP_ShowToday, DEF_ShowToday)
    mbShowTodayCircle = PropBag.ReadProperty(PROP_ShowTodayCircle, DEF_ShowTodayCircle)
    mbMultiSelect = PropBag.ReadProperty(PROP_MultiSelect, DEF_MultiSelect)
    mdMin = PropBag.ReadProperty(PROP_Min, DEF_Min)
    mdMax = PropBag.ReadProperty(PROP_Max, DEF_Max)
    miMaxSel = PropBag.ReadProperty(PROP_MaxSel, DEF_MaxSel)
    miBorderStyle = PropBag.ReadProperty(PROP_BorderStyle, DEF_BorderStyle)
    UserControl.Enabled = PropBag.ReadProperty(PROP_Enabled, DEF_Enabled)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    pCreate
End Sub

Private Sub UserControl_Resize()
    If mhWnd Then MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
    'pAutoSize
End Sub
'
'Private Sub 'pAutoSize()
'    Static bExit As Boolean
'    If bExit Then Exit Sub
'
'    if mhWnd then
'        Dim lR as RECT
'
'        SendMessage mhWnd, MCM_GETMINREQRECT, ZeroL, VarPtr(lR)
'
'        'this is far from perfect...
'        lR.Right = ((lR.Right + 6&) * miColumns)
'        lR.Bottom = ((lR.Bottom + TwoL) * miRows) - (12& * (miRows - OneL))
'
'        bExit = True
'        UserControl.Size ScaleX(lR.Right + Choose(miBorderStyle + OneL, ZeroL, -miColumns - 6&, ZeroL, ZeroL), vbPixels, vbTwips), ScaleY(lR.Bottom + (Choose(miBorderStyle + OneL, ZeroL, (-miRows + OneL) * TwoL, TwoL, TwoL)), vbPixels, vbTwips)
'        bExit = False
'
'        if mhWnd then
'            moWnd.Move ZeroL, ZeroL, ScaleWidth, ScaleHeight
'        End If
'
'    End If
'End Sub


Private Sub UserControl_Show()
    If mhWnd Then
        Static bShown As Boolean
        If Not bShown Then
            bShown = True
            GetBoldDays
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    pDestroy
    If mhFont Then moFont.ReleaseHandle mhFont
    ReleaseShellMod
    Set moFontPage = Nothing
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty PROP_ColorTrailingBack, miColors(MCSC_BACKGROUND), DEF_ColorTrailingBack
    PropBag.WriteProperty PROP_ColorBack, miColors(MCSC_MONTHBK), DEF_ColorBack
    PropBag.WriteProperty PROP_ColorTitleBack, miColors(MCSC_TITLEBK), DEF_ColorTitleBack
    PropBag.WriteProperty PROP_ColorText, miColors(MCSC_TEXT), DEF_ColorText
    PropBag.WriteProperty PROP_ColorTitleText, miColors(MCSC_TITLETEXT), DEF_ColorTitleText
    PropBag.WriteProperty PROP_ColorTrailingText, miColors(MCSC_TRAILINGTEXT), DEF_ColorTrailingText
    PropBag.WriteProperty PROP_ShowWeekNumbers, mbShowWeekNumbers, DEF_ShowWeekNumbers
    PropBag.WriteProperty PROP_ShowToday, mbShowToday, DEF_ShowToday
    PropBag.WriteProperty PROP_ShowTodayCircle, mbShowTodayCircle, DEF_ShowTodayCircle
    PropBag.WriteProperty PROP_MultiSelect, mbMultiSelect, DEF_MultiSelect
    PropBag.WriteProperty PROP_Min, mdMin, DEF_Min
    PropBag.WriteProperty PROP_Max, mdMax, DEF_Max
    PropBag.WriteProperty PROP_MaxSel, miMaxSel, DEF_MaxSel
    PropBag.WriteProperty PROP_BorderStyle, miBorderStyle, DEF_BorderStyle
    PropBag.WriteProperty PROP_Enabled, UserControl.Enabled, DEF_Enabled
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    Font_Write moFont, PropBag, PROP_Font
    
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
        mbThemeable = bNew
        If mhWnd Then EnableWindowTheme mhWnd, mbThemeable
        PropertyChanged PROP_Themeable
        'moWnd.SetPos , , , , , SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
        'moWnd.Redraw RDW_INVALIDATE
    End If
End Property

