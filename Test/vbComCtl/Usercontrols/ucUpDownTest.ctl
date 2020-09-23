VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#109.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucUpDownTest 
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   HasDC           =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5355
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   6165
      Begin vbComCtl.ucUpDown ud 
         Height          =   300
         Index           =   4
         Left            =   4965
         Top             =   1500
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         Min             =   -2147483648
         Max             =   2147483647
         BProps          =   327846
         Buddy           =   "txt(4)"
         BuddyProp       =   "Text"
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   540
         TabIndex        =   5
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   540
         TabIndex        =   6
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   540
         TabIndex        =   7
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   3780
         TabIndex        =   8
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   3780
         TabIndex        =   9
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   3780
         TabIndex        =   10
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "Hexadecimal"
         Height          =   255
         Index           =   0
         Left            =   2340
         TabIndex        =   2
         Top             =   660
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "Show Thousands Sep"
         Height          =   255
         Index           =   1
         Left            =   2340
         TabIndex        =   4
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CheckBox chk 
         Caption         =   "Wrap"
         Height          =   255
         Index           =   2
         Left            =   4020
         TabIndex        =   3
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucUpDownTest.ctx":0000
         Left            =   2340
         List            =   "ucUpDownTest.ctx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   2895
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   300
         Index           =   0
         Left            =   1725
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         BProps          =   327846
         Buddy           =   "txt(0)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   6
         Left            =   4980
         Top             =   2220
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         Min             =   -2147483648
         Max             =   2147483647
         BProps          =   327846
         Buddy           =   "txt(6)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   5
         Left            =   4980
         Top             =   1860
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         Min             =   -2147483648
         Max             =   2147483647
         BProps          =   327846
         Buddy           =   "txt(5)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   3
         Left            =   1740
         Top             =   2220
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         Min             =   -2147483648
         Max             =   2147483647
         BProps          =   327846
         Buddy           =   "txt(3)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   2
         Left            =   1740
         Top             =   1860
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         Min             =   -2147483648
         Max             =   2147483647
         BProps          =   327846
         Buddy           =   "txt(2)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   1
         Left            =   1740
         Top             =   1500
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         Min             =   -2147483648
         Max             =   2147483647
         BProps          =   327846
         Buddy           =   "txt(1)"
         BuddyProp       =   "Text"
      End
      Begin VB.Label Label1 
         Caption         =   "Min:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Max:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   1860
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Pos:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Small Change:"
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   14
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Large Change:"
         Height          =   255
         Index           =   4
         Left            =   2700
         TabIndex        =   13
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Large Change Delay:"
         Height          =   255
         Index           =   5
         Left            =   2220
         TabIndex        =   12
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   $"ucUpDownTest.ctx":0064
         Height          =   615
         Index           =   6
         Left            =   300
         TabIndex        =   11
         Top             =   2700
         Width           =   4455
      End
   End
End
Attribute VB_Name = "ucUpDownTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucUpDownTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide funtionality for testing the updown control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eTxt
    txtMain
    txtMin
    txtMax
    txtPos
    txtSmall
    txtLarge
    txtDelay
End Enum

Private Enum eChk
    chkHex
    chkThou
    chkWrap
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
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the constituent controls.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowAllUIStates hWnd
    With ud
        .Item(txtMin).Value = .Item(txtMain).Min
        .Item(txtMax).Value = .Item(txtMain).Max
        .Item(txtPos).Value = .Item(txtMain).Value
        .Item(txtSmall).Value = .Item(txtMain).SmallChange
        .Item(txtLarge).Value = .Item(txtMain).LargeChange
        .Item(txtDelay).Value = .Item(txtMain).LargeChangeDelay
    End With
    cmb.ListIndex = ud(txtMain).BuddyAlignment
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    ucScrollBox1.Height = Height
End Sub

Private Sub chk_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the appropriate property of the updown control.
'---------------------------------------------------------------------------------------
    Dim lbVal As Boolean:   lbVal = chk(Index).Value
    Select Case Index
    Case chkHex:    ud(txtMain).Hexadecimal = lbVal
    Case chkThou:   ud(txtMain).ShowThousandsSeparator = lbVal
    Case chkWrap:   ud(txtMain).Wrap = lbVal
    End Select
End Sub

Private Sub cmb_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the appropriate property of the updown control.
'---------------------------------------------------------------------------------------
    Dim liAlign As evbComCtlAlignment
    liAlign = cmb.ListIndex
    ud(txtMain).BuddyAlignment = liAlign
    If liAlign <> vbccAlignNone Then ud(txtMain).Orientation = udHorizontal * -CBool(liAlign < vbccAlignLeft)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Allow the updown controls to change on keypress events.
'---------------------------------------------------------------------------------------
    ud(Index).BuddyKeyDown KeyCode, Shift
End Sub

Private Sub txt_LostFocus(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Cause changes made by the user to take effect.
'---------------------------------------------------------------------------------------
    ud(Index).SyncValueFromBuddy
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Cause changes made by the user to take effect.
'---------------------------------------------------------------------------------------
    ud(Index).SyncValueFromBuddy
End Sub

Private Sub ud_Change(Index As Integer, ByVal iValue As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the appropriate property of the updown control.
'---------------------------------------------------------------------------------------
    Select Case Index
    Case txtMain:   ud(txtPos).Value = iValue
    Case txtMin:    ud(txtMain).Min = iValue
    Case txtMax:    ud(txtMain).Max = iValue
    Case txtPos:    ud(txtMain).Value = iValue
    Case txtSmall:  ud(txtMain).SmallChange = iValue
    Case txtLarge:  ud(txtMain).LargeChange = iValue
    Case txtDelay:  ud(txtMain).LargeChangeDelay = iValue
    End Select
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property
