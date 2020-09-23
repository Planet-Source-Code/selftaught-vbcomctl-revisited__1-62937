VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucDateTimePickerTest 
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   6510
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   5530
      Begin VB.CheckBox chk 
         Caption         =   "Enabled"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "Right Align Cal"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucDateTimePickerTest.ctx":0000
         Left            =   1980
         List            =   "ucDateTimePickerTest.ctx":000D
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Checkbox"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   3
         Top             =   660
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Today"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Today Circle"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   5
         Top             =   1260
         Width           =   1695
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Weeknumbers"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Caption         =   "Updown"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   7
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Colors ..."
         Height          =   375
         Left            =   4260
         TabIndex        =   8
         Top             =   1980
         Width           =   1995
      End
      Begin vbComCtl.ucDateTimePicker dtp 
         Height          =   555
         Left            =   0
         TabIndex        =   10
         Top             =   2460
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   979
         Value           =   38403.5636574074
         BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
         EndProperty
         BeginProperty CalFont {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
         EndProperty
      End
      Begin vbComCtlTest.ucEvents log 
         Height          =   1635
         Left            =   4260
         TabIndex        =   9
         Top             =   240
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   2884
      End
      Begin VB.Label Label1 
         Caption         =   "Format:"
         Height          =   255
         Index           =   0
         Left            =   1980
         TabIndex        =   12
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         Height          =   1755
         Index           =   1
         Left            =   1980
         TabIndex        =   11
         Top             =   600
         Width           =   1755
      End
   End
End
Attribute VB_Name = "ucDateTimePickerTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucDateTimePickerTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the datetimepicker control
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eChk
    chkEnabled
    chkRightAlign
    chkShowCheckbox
    chkShowToday
    chkShowTodayCircle
    chkShowWeeknumbers
    chkUpdown
End Enum

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the font or back color properties to match the container.
'---------------------------------------------------------------------------------------
    Select Case ZeroL
    Case StrComp(PropertyName, "Font")
        Set UserControl.Font = Ambient.Font
        dtp.Height = ScaleX(dtp.Font.TextHeight("A") + 8, vbPixels, vbTwips)
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
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the constituent controls.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowAllUIStates hWnd
    With chk
        .Item(chkEnabled).Value = -dtp.Enabled
        .Item(chkRightAlign) = -dtp.RightAlignedCal
        .Item(chkShowCheckbox) = -dtp.ShowCheckBox
        .Item(chkShowToday) = -dtp.ShowToday
        .Item(chkShowTodayCircle) = -dtp.ShowTodayCircle
        .Item(chkShowWeeknumbers) = -dtp.ShowWeekNumbers
        .Item(chkUpdown) = -dtp.UpDown
    End With
    
    Label1(1).Caption = "d dd ddd dddd" & vbNewLine & _
                        "h hh" & vbNewLine & _
                        "m mm" & vbNewLine & _
                        "s ss" & vbNewLine & _
                        "t tt" & vbNewLine & _
                        "y yy yyy yyyy" & vbNewLine & _
                        "H HH" & vbNewLine & _
                        "M MM MM MMMM"
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    Set UserControl.Font = Ambient.Font
    dtp.Height = ScaleY(dtp.Font.TextHeight("A") + 8, vbPixels, vbTwips)
    ucScrollBox1.AutoSize
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Correct the width of the control.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    ucScrollBox1.Move 0, 0, Width, Height
    On Error GoTo 0
End Sub

Private Sub chk_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Change the corresponding property of the control.
'---------------------------------------------------------------------------------------
    Dim lbVal As Boolean
    lbVal = chk(Index).Value
    Select Case Index
    Case chkEnabled:            dtp.Enabled = lbVal
    Case chkRightAlign:         dtp.RightAlignedCal = lbVal
    Case chkShowCheckbox:       dtp.ShowCheckBox = lbVal
    Case chkShowToday:          dtp.ShowToday = lbVal
    Case chkShowTodayCircle:    dtp.ShowTodayCircle = lbVal
    Case chkShowWeeknumbers:    dtp.ShowWeekNumbers = lbVal
    Case chkUpdown:             dtp.UpDown = lbVal
    End Select
End Sub

Private Sub cmb_Change()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the new format string.
'---------------------------------------------------------------------------------------
    dtp.FormatString = cmb.Text
End Sub

Private Sub cmb_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the new predefined format.
'---------------------------------------------------------------------------------------
    dtp.Format = cmb.ItemData(cmb.ListIndex)
End Sub

Private Sub cmd_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Allow the user to edit the control's colors.
'---------------------------------------------------------------------------------------
    fEditColors.EditColors Parent, dtp, "ColorTrailingText", "ColorTitleText", "ColorTitleBackground", "ColorText", "ColorBackground"
End Sub

Private Sub dtp_DateTimeChange()
    log.LogItem "DTChange " & dtp.Value
End Sub

Private Sub dtp_DropDown(ByVal hWndMonthCal As Long)
    log.LogItem "Dropdown " & hWndMonthCal
End Sub

Private Sub dtp_CloseUp()
    log.LogItem "CloseUp"
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property
