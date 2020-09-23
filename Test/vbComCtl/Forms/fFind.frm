VERSION 5.00
Begin VB.Form fFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chk 
      Caption         =   "Match Case"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   660
      Width           =   1215
   End
   Begin VB.CheckBox chk 
      Caption         =   "Whole Word"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3180
      TabIndex        =   2
      Top             =   660
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Find"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3180
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2595
   End
   Begin VB.Label lbl 
      BackColor       =   &H8000000D&
      Caption         =   " Passed the end of the document"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   4710
   End
End
Attribute VB_Name = "fFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
'fFind.frm      3/31/05
'
'           PURPOSE:
'               Find text in a ucRichEdit control.
'
'==================================================================================================

Option Explicit

Private Enum eCmd
    cmdFind
    cmdCancel
End Enum

Private Enum eChk
    chkWholeWord
    chkMatchCase
End Enum

Event FindText(ByRef sText As String, ByVal bWholeWord As Boolean, ByVal bMatchCase As Boolean, ByRef bPassedEnd As Boolean)

Private Sub Form_Load()
    vbComCtl.ShowAllUIStates hwnd
End Sub

Private Sub cmd_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Perform the default or cancel dialog actions.
'---------------------------------------------------------------------------------------
    If Index = cmdFind Then
        Dim lbPassedEnd As Boolean
        RaiseEvent FindText(txt.Text, chk(chkWholeWord).Value, chk(chkMatchCase).Value, lbPassedEnd)
        lbl.Visible = lbPassedEnd
    ElseIf Index = cmdCancel Then
        Unload Me
    End If
End Sub

Private Sub txt_Change()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Enable the Find button when there is text to search for.
'---------------------------------------------------------------------------------------
    cmd(cmdFind).Enabled = CBool(Len(txt.Text))
End Sub

Public Sub ShowFind(ByVal oOwner As Form, ByRef sText As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the dialog.
'---------------------------------------------------------------------------------------
    If Len(sText) Then txt.Text = sText
    Show vbModeless, oOwner
End Sub
