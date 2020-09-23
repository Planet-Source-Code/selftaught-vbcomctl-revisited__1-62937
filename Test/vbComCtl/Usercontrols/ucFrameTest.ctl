VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucFrameTest 
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   HasDC           =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6600
   Begin vbComCtl.ucFrame fra 
      Align           =   1  'Align Top
      Height          =   3600
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   6350
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Caption         =   "Common Controls 6.0 Test"
      Begin vbComCtl.ucScrollBox ucScrollBox1 
         Height          =   3615
         Left            =   180
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6376
         Begin VB.Frame Frame1 
            Caption         =   "VB.Frame"
            Height          =   1500
            Left            =   0
            TabIndex        =   3
            Top             =   1560
            Width           =   4290
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   375
               Index           =   1
               Left            =   180
               TabIndex        =   5
               Top             =   1020
               Width           =   1335
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Option1"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   4
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Flickers on mouse in/out and partially transparent controls do not display ok!"
               Height          =   735
               Index           =   1
               Left            =   1800
               TabIndex        =   6
               Top             =   720
               Width           =   1935
            End
         End
         Begin vbComCtl.ucFrame fra 
            Height          =   1500
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   2646
            BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
            EndProperty
            Caption         =   "vbComCtl.ucFrame"
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   375
               Index           =   0
               Left            =   180
               TabIndex        =   1
               Top             =   1020
               Width           =   1335
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Option1"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   0
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "No flicker on mouse in/out and partially transparent controls display ok!"
               Height          =   735
               Index           =   0
               Left            =   1800
               TabIndex        =   2
               Top             =   720
               Width           =   1935
            End
         End
      End
   End
End
Attribute VB_Name = "ucFrameTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucFrameTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the frame control
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the font or back color properties to match the container.
'---------------------------------------------------------------------------------------
    Select Case ZeroL
    Case StrComp(PropertyName, "Font")
        Set UserControl.Font = Ambient.Font
        Set Frame1.Font = Ambient.Font
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
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    fra(1).Height = Height
End Sub

Private Sub fra_Resize(Index As Integer)
    If Index = 1 Then fra(Index).MoveToClient ucScrollBox1
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property
