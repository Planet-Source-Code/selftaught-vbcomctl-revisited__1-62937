VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucHotKeyTest 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   5760
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   1875
         Left            =   2820
         ScaleHeight     =   1875
         ScaleWidth      =   2595
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   2595
         Begin VB.ListBox lst 
            Height          =   1860
            ItemData        =   "ucHotKeyTest.ctx":0000
            Left            =   0
            List            =   "ucHotKeyTest.ctx":001C
            Style           =   1  'Checkbox
            TabIndex        =   1
            Top             =   0
            Width           =   2595
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Set Application Hotkey"
         Height          =   495
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Height          =   285
         Left            =   60
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucHotKeyTest.ctx":0088
         Left            =   1020
         List            =   "ucHotKeyTest.ctx":00A4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin vbComCtl.ucHotKey hkey 
         Height          =   300
         Left            =   60
         TabIndex        =   4
         Top             =   2220
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   529
         BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Hot Key:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Modifier:"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Invalid Modifiers:"
         Height          =   255
         Index           =   2
         Left            =   2820
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ucHotKeyTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucHotKeyTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the hotkey control.
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
        hkey.Height = ScaleY(hkey.Font.TextHeight("A") + 6, vbPixels, vbTwips)
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
    hkey.Height = ScaleY(hkey.Font.TextHeight("A") + 6, vbPixels, vbTwips)
    ucScrollBox1.AutoSize
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    ucScrollBox1.Move 0, 0, Width, Height
End Sub

Private Sub cmb_Click()
   If cmb.ListIndex > -1 Then hkey.HotKeyModifier = cmb.ItemData(cmb.ListIndex)
End Sub

Private Sub cmd_Click()
    Dim liResult As eHotKeySetAppHotKeyResult
    liResult = hkey.SetApplicationHotKey()
    Select Case liResult
    Case hotInvalidHotKey:          MsgBox "Error: Invalid HotKey"
    Case hotInvalidWindow:          MsgBox "Error: Invalid Window"
    Case hotSuccess:                MsgBox "Hotkey set!"
    Case hotSuccessWithDuplicate:   MsgBox "Hotkey set, but somebody else beat you to it."
    Case Else:                      MsgBox "Error: " & liResult
    End Select
End Sub

Private Sub hkey_Change()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the combobox to show the current hotkey selection.
'---------------------------------------------------------------------------------------
    txt.Text = ChrW$(hkey.HotKey)
    Dim i As Long
    For i = cmb.ListCount - 1 To 0 Step -1
        If cmb.ItemData(i) = hkey.HotKeyModifier Then Exit For
    Next
    cmb.ListIndex = i
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    hkey.InvalidHotKeyOperation(lst.ItemData(Item), hotNone) = lst.Selected(Item)
End Sub

Private Sub txt_Change()
    If Len(txt.Text) Then hkey.HotKey = Asc(UCase$(txt.Text))
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Public Property Get IsHotKeyActive() As Boolean
    IsHotKeyActive = (ActiveControl Is hkey)
End Property
