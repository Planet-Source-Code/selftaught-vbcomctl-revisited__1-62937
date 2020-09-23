VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucTrackbarTest 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   3435
      Left            =   0
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   6059
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucTrackbarTest.ctx":0000
         Left            =   1800
         List            =   "ucTrackbarTest.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2340
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tooltips"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2340
         Width           =   915
      End
      Begin vbComCtl.ucTrackbar kbar 
         Height          =   615
         Left            =   240
         TabIndex        =   0
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         ToolTips        =   0   'False
         Back            =   -2147483633
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   5
         Left            =   3480
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         BProps          =   327719
         Buddy           =   "txt(5)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   4
         Left            =   3480
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         BProps          =   327719
         Buddy           =   "txt(4)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   3
         Left            =   3480
         Top             =   960
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         BProps          =   327719
         Buddy           =   "txt(3)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   2
         Left            =   1080
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         BProps          =   327719
         Buddy           =   "txt(2)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   1
         Left            =   1080
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         BProps          =   327719
         Buddy           =   "txt(1)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   0
         Left            =   1080
         Top             =   960
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         BProps          =   327719
         Buddy           =   "txt(0)"
         BuddyProp       =   "Text"
      End
      Begin VB.Label Label1 
         Caption         =   "Tic Style:"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   15
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Min:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Max"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Pos"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Line Size:"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Page Size"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tic Freq:"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
   End
End
Attribute VB_Name = "ucTrackbarTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucTrackbarTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the trackbar control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eTxt
    txtMin
    txtMax
    txtPos
    txtLine
    txtPage
    txtTic
End Enum

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

Private Sub Usercontrol_EnterFocus()
    vbComCtl.EnterFocus Controls
End Sub

Private Sub UserControl_Initialize()
    vbComCtl.ShowAllUIStates hWnd
    With ud
        .Item(txtMin).Value = kbar.Min
        .Item(txtMax).Value = kbar.Max
        .Item(txtPos).Value = kbar.Pos
        .Item(txtLine).Value = kbar.LineSize
        .Item(txtPage).Value = kbar.PageSize
        .Item(txtTic).Value = kbar.TicFreq
    End With
    cmb.ListIndex = trkBottomOrRight
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
    ucScrollBox1.Move 0, 0, Width, Height
End Sub

Private Sub cmb_Click()
    kbar.TicStyle = cmb.ListIndex
End Sub

Private Sub kbar_Change()
    ud(txtPos).Value = kbar.Pos
End Sub

Private Sub ud_Change(Index As Integer, ByVal iValue As Long)
    Dim liVal As Long
    liVal = ud(Index).Value
    Select Case Index
    Case txtMin: kbar.Min = liVal
    Case txtMax: kbar.Max = liVal
    Case txtPos: kbar.Pos = liVal
    Case txtLine: kbar.LineSize = liVal
    Case txtPage: kbar.PageSize = liVal
    Case txtTic: kbar.TicFreq = liVal
    End Select
End Sub

Private Sub Check1_Click()
    kbar.ToolTips = Check1.Value
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property
