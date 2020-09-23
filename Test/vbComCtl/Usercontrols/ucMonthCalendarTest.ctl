VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#109.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucMonthCalendarTest 
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   6330
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   3195
      Left            =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5636
      Begin VB.CheckBox chk 
         Caption         =   "Wednesdays in Bold"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Weeknumbers"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Today Circle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Today"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Colors ..."
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucMonthCalendarTest.ctx":0000
         Left            =   120
         List            =   "ucMonthCalendarTest.ctx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin vbComCtl.ucMonthCalendar cal 
         Height          =   1755
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   3096
         BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
         EndProperty
      End
      Begin vbComCtlTest.ucEvents ucEvents1 
         Height          =   2055
         Left            =   3780
         TabIndex        =   8
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
      End
      Begin VB.Label Label1 
         Caption         =   "Border:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   615
      End
   End
End
Attribute VB_Name = "ucMonthCalendarTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucMonthCalendarTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the month calendar control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eChk   'checkboxes
    chkEnabled
    chkShowToday
    chkShowTodayCircle
    chkShowWeeknumbers
    chkWednesdayBold
End Enum

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the font or back color properties to match the container.
'---------------------------------------------------------------------------------------
    Select Case ZeroL
    Case StrComp(PropertyName, "Font")
        Set UserControl.Font = Ambient.Font
        With cal
            .Move .Left, .Top, .MinReqWidth, .MinReqHeight
            ucEvents1.Left = .Left + .Width + 150
        End With
        ucScrollBox1.AutoSize
    Case StrComp(PropertyName, "BackColor")
        UserControl.BackColor = Ambient.BackColor
        vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
    End Select
End Sub

Private Sub Usercontrol_EnterFocus()
    vbComCtl.EnterFocus Controls
End Sub

Private Sub UserControl_Initialize()
    vbComCtl.ShowAllUIStates hWnd
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    Set UserControl.Font = Ambient.Font
    With cal
        .Move .Left, .Top, .MinReqWidth, .MinReqHeight
        ucEvents1.Left = .Left + .Width + 150
    End With
    ucScrollBox1.AutoSize
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    ucScrollBox1.Move 0, 0, Width, Height
End Sub

Private Sub cal_Change()
    ucEvents1.LogItem "Change " & cal.SelDate
End Sub

Private Sub cal_GetBoldDays(ByVal iMonth As Long, ByVal iYear As Long, iMask As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a bitmask indicating which days of the month to display in bold.
'---------------------------------------------------------------------------------------
    If chk(chkWednesdayBold).Value Then
        Dim dDate As Date: dDate = DateSerial(iYear, iMonth, OneL)
        Dim liBit As Long: liBit = OneL
        
        Do While Month(dDate) = iMonth
            If Weekday(dDate) = vbWednesday Then iMask = iMask Or liBit
            dDate = DateAdd("d", OneL, dDate)
            If liBit < &H40000000 Then liBit = liBit + liBit
        Loop
    End If
End Sub

Private Sub cmb_Click()
    cal.BorderStyle = cmb.ListIndex
    With cal
        .Move .Left, .Top, .MinReqWidth, .MinReqHeight
        ucEvents1.Left = .Left + .Width + 150
    End With
End Sub

Private Sub chk_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the appropriate property of the monthcalendar.
'---------------------------------------------------------------------------------------
    Dim lbVal As Boolean
    lbVal = chk(Index).Value
    Select Case Index
    Case chkEnabled:            cal.Enabled = lbVal
    Case chkShowToday:          cal.ShowToday = lbVal
    Case chkShowTodayCircle:    cal.ShowTodayCircle = lbVal
    Case chkShowWeeknumbers:    cal.ShowWeekNumbers = lbVal
    Case chkWednesdayBold:      cal.GetBoldDays
    End Select
    With cal
        .Move .Left, .Top, .MinReqWidth, .MinReqHeight
        ucEvents1.Left = .Left + .Width + 150
    End With
    ucScrollBox1.AutoSize
End Sub

Private Sub cmd_Click()
    fEditColors.EditColors Parent, cal, "ColorTrailingText", "ColorTitleText", "ColorTitleBackground", "ColorText", "ColorBackground"
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property
