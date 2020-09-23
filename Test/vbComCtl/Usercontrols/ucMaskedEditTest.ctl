VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucMaskedEditTest 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   5445
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   2778
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "ucMaskedEditTest.ctx":0000
         Left            =   120
         List            =   "ucMaskedEditTest.ctx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Valid"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1260
         Width           =   975
      End
      Begin vbComCtl.ucMaskedEdit ucMaskedEdit1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
         EndProperty
         Mask            =   ""
         Themeable       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Format:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "ucMaskedEditTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucMaskedEditTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the maskededit control.
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
        ucMaskedEdit1.Height = ScaleY(ucMaskedEdit1.Font.TextHeight("A") + 6, vbPixels, vbTwips)
        Check1.Top = ucMaskedEdit1.Top + ucMaskedEdit1.Height + 400
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
    cmb.ListIndex = 0
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    Set UserControl.Font = Ambient.Font
    ucMaskedEdit1.Height = ScaleY(ucMaskedEdit1.Font.TextHeight("A") + 6, vbPixels, vbTwips)
    Check1.Top = ucMaskedEdit1.Top + ucMaskedEdit1.Height + 420
    ucScrollBox1.AutoSize
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Adjust the position of the constituent controls.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    'make the combobox and masked edit as thin as possible so that
    'they do not cause the scrollbox to dislay a horizontal scrollbar.
    'also hide them to reduce flicker.
    cmb.Visible = False
    cmb.Width = 0
    ucMaskedEdit1.Visible = False
    ucMaskedEdit1.Width = 0
    ucScrollBox1.Move 0, 0, Width, Height
    ucScrollBox1.AutoSize
    With cmb
        .Visible = True
        .Move .Left, .Top, Width - .Left - .Left - ucScrollBox1.ScrollBarWidth
    End With
    With ucMaskedEdit1
        .Visible = True
        .Move .Left, .Top, Width - .Left - .Left - ucScrollBox1.ScrollBarWidth, .Height
    End With
    Refresh
    On Error GoTo 0
End Sub

Private Sub cmb_Change()
    ucMaskedEdit1.Mask = cmb.Text
End Sub

Private Sub cmb_Click()
    ucMaskedEdit1.Mask = cmb.Text
End Sub

Private Sub ucMaskedEdit1_Changed()
    Check1.Value = -ucMaskedEdit1.IsValid
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property
