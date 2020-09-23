VERSION 5.00
Begin VB.UserControl ucEvents 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   2310
   Begin VB.CommandButton cmd 
      Caption         =   "Clear"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ListBox lst 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "ucEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucEvents.ctl           3/31/05
'
'            PURPOSE:
'               Display a log of events.
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

Private Sub Usercontrol_EnterFocus()
    vbComCtl.EnterFocus Controls
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
    If Ambient.UserMode Then vbComCtl.ShowAllUIStates hWnd
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then vbComCtl.ShowAllUIStates hWnd
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    lst.Width = UserControl.Width
    cmd.Width = UserControl.Width
    cmd.Top = Height - cmd.Height
    lst.Height = Height - cmd.Height
    On Error GoTo 0
End Sub

Private Sub cmd_Click()
    lst.Clear
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Public Sub LogItem(ByRef sItem As String)
    If lst.ListCount >= 100 Then lst.RemoveItem 0
    lst.AddItem sItem
    lst.ListIndex = lst.NewIndex
End Sub

