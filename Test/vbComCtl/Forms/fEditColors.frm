VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#95.0#0"; "vbComCtl.ocx"
Begin VB.Form fEditColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Colors"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin vbComCtl.ucComDlg dlg 
      Left            =   3840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "fEditColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'fEditColors.frm           3/28/05
'
'            PURPOSE:
'               Display an interface for editing any number of color properties on any object.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private moObject As Object

Private Sub Form_Load()
    vbComCtl.ShowAllUIStates hwnd
    pStoreCustomColors dlg, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pStoreCustomColors dlg, False
End Sub

Public Sub EditColors(ByVal oOwner As Form, ByVal oObject As Object, ParamArray vPropertyNames())
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : show the command buttons and display the form.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim i As Long
    Dim liColor As Long
    
    For i = cmd.LBound + OneL To cmd.UBound
        Unload lbl(i)
        Unload cmd(i)
    Next
    
    For i = cmd.LBound To cmd.LBound + UBound(vPropertyNames) - LBound(vPropertyNames)
        If i > cmd.LBound Then
            Load cmd(i)
            Load lbl(i)
            lbl(i).Visible = True
            cmd(i).Visible = True
            lbl(i).Top = cmd(i - 1).Top + cmd(i - 1).Height + 20
            cmd(i).Top = lbl(i).Top + lbl(i).Height
        End If
        
        lbl(i).Caption = vPropertyNames(i - cmd.LBound)
        liColor = CallByName(oObject, vPropertyNames(i - cmd.LBound), VbGet)
        cmd(i).BackColor = IIf(liColor = NegOneL, vbButtonFace, liColor)
    Next
    Height = cmd(i - 1).Top + cmd(i - 1).Height + cmd(i - 1).Height
    Set moObject = oObject
    Show vbModal, oOwner
    Set moObject = Nothing
    On Error GoTo 0
End Sub

Private Sub cmd_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the color dialog for the property that was clicked.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    Dim liColor As OLE_COLOR
    If dlg.ShowColor(CallByName(moObject, lbl(Index).Caption, VbGet), dlgColorAny) Then
        CallByName moObject, lbl(Index).Caption, VbLet, dlg.Color
        cmd(Index).BackColor = dlg.Color
    End If
handler:
    If Err.Number Then MsgBox "Error: " & Err.Number & vbNewLine & Err.Description
    On Error GoTo 0
End Sub

Private Sub cmdDone_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Dismiss the dialog.
'---------------------------------------------------------------------------------------
    Hide
End Sub


Private Sub pStoreCustomColors(ByVal oDlg As ucComDlg, ByVal bLoad As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Store the custom colors used by a color dialog between instances.
'---------------------------------------------------------------------------------------
    Static liColors() As Long
    Static bInit As Boolean
    
    Dim i As Long
    
    If bLoad Then
        If bInit Then
            For i = ZeroL To oDlg.ColorCustomCount
                oDlg.ColorCustom(i) = liColors(i)
            Next
        End If
    Else
        bInit = True
        ReDim liColors(0 To oDlg.ColorCustomCount)
        For i = ZeroL To oDlg.ColorCustomCount
            liColors(i) = oDlg.ColorCustom(i)
        Next
    End If
End Sub
