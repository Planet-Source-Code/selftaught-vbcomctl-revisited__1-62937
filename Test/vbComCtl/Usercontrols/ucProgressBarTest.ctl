VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucProgressBarTest 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   1395
      Left            =   0
      Top             =   480
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   2461
      Begin VB.CheckBox Check1 
         Caption         =   "Animate"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Step"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Smooth"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1800
      Top             =   1560
   End
   Begin vbComCtl.ucProgressBar pbar 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   847
      Smooth          =   -1  'True
   End
End
Attribute VB_Name = "ucProgressBarTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucProgressBarTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the progressbar control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the font or back color properties to match the container.
'---------------------------------------------------------------------------------------
    Select Case ZeroL
    Case StrComp(PropertyName, "BackColor")
        UserControl.BackColor = Ambient.BackColor
        vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
    End Select
End Sub

Private Sub UserControl_Initialize()
    vbComCtl.ShowAllUIStates hWnd
End Sub

Private Sub Usercontrol_EnterFocus()
    vbComCtl.EnterFocus Controls
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With ucScrollBox1
        .Move .Left, .Top, Width - .Left, Height - .Top
    End With
    On Error GoTo 0
End Sub


Private Sub Check1_Click()
    Timer1.Enabled = Check1.Value
End Sub

Private Sub Check2_Click()
    pbar.Smooth = Check2.Value
End Sub

Private Sub Command1_Click()
    pbar.Value = (pbar.Value + 1) Mod 101
End Sub

Private Sub Timer1_Timer()
    Command1_Click
End Sub
