VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#111.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucAnimationTest 
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   8580
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   2595
      Left            =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4577
      Begin VB.CommandButton cmd 
         Caption         =   "Play"
         Height          =   495
         Index           =   2
         Left            =   6180
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Open From Resource"
         Height          =   495
         Index           =   0
         Left            =   6180
         TabIndex        =   0
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Open From File ..."
         Height          =   495
         Index           =   1
         Left            =   6180
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chk 
         Caption         =   "Center"
         Height          =   255
         Index           =   0
         Left            =   6180
         TabIndex        =   2
         Top             =   1800
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chk 
         Caption         =   "Autoplay"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   3
         Top             =   1800
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chk 
         Caption         =   "Timer"
         Height          =   255
         Index           =   2
         Left            =   6180
         TabIndex        =   4
         Top             =   2160
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chk 
         Caption         =   "Transparent"
         Height          =   255
         Index           =   3
         Left            =   7080
         TabIndex        =   5
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin vbComCtl.ucAnimation ani 
         Height          =   2055
         Left            =   60
         Top             =   60
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   3625
         Tmr             =   -1  'True
      End
      Begin VB.Shape Shape1 
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   5955
      End
   End
   Begin vbComCtl.ucComDlg dlg 
      Left            =   0
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "ucAnimationTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucAnimationTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the animation control.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eChk
    chkCenter
    chkAutoPlay
    chkTimer
    chkTransparent
End Enum

Private Enum eCmd
    cmdLoadFromResource
    cmdLoadFromFile
    cmdPlayStop
End Enum

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the font or back color properties to match the container.
'---------------------------------------------------------------------------------------
    Select Case ZeroL
    Case StrComp(PropertyName, "Font")
        Set UserControl.Font = Ambient.Font
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
    cmd(cmdLoadFromResource).Enabled = Not InIDE()
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    Set UserControl.Font = Ambient.Font
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor Controls, Ambient.BackColor
End Sub

Private Sub chk_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set a property based on the user's selection.
'---------------------------------------------------------------------------------------
    Dim lb As Boolean:      lb = CBool(chk(Index).Value)
    
    Select Case Index
    Case chkCenter:         ani.Center = lb
    Case chkAutoPlay:       ani.AutoPlay = lb
    Case chkTimer:          ani.Timer = lb
    Case chkTransparent:    ani.Transparent = lb
    End Select
    
    If Index <> chkAutoPlay Then cmd(cmdPlayStop).Enabled = False
End Sub

Private Sub cmd_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Load, play or stop the animation based on the user's selection.
'---------------------------------------------------------------------------------------
    Select Case Index
    Case cmdLoadFromFile, cmdLoadFromResource
        Dim lbLoaded As Boolean
        If Index = cmdLoadFromResource Then
            lbLoaded = ani.LoadFromResource(App.hInstance, 101)
        Else
            Dim lsFile As String
            lbLoaded = dlg.ShowFileOpen(lsFile, dlgFileMustExist Or dlgFileExplorerStyle Or dlgFileHideReadOnly, dlg.FileGetFilter("Audio Video Interlay Files (*.avi)", "*.avi", "All Files (*.*)", "*.*"), , "avi", , , "Select an file for animation")
            If lbLoaded Then
                lbLoaded = ani.LoadFromFile(lsFile)
                If Not lbLoaded Then MsgBox "Invalid File Type or File/Path access error on " & lsFile
            End If
        End If
        
        If lbLoaded Then
            cmd(cmdPlayStop).Caption = IIf(ani.AutoPlay, "Stop", "Play")
            cmd(cmdPlayStop).Enabled = True
        End If
    Case cmdPlayStop
        If cmd(Index).Caption = "Stop" Then
            If ani.StopPlaying() Then cmd(Index).Caption = "Play"
        Else
            If ani.Play() Then cmd(Index).Caption = "Stop"
        End If
    End Select
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    ucScrollBox1.Move 0, 0, Width, Height
    On Error GoTo 0
End Sub

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Public Sub StressTest()
    Dim i As Long
    
    On Error Resume Next
    
    Dim lsFileName As String
    lsFileName = Environ$("TEMP") & "\Srch9537.avi"
    
    Dim lyB() As Byte
    lyB = LoadResData(101, "AVI")
    
    Dim liFile As Long
    liFile = FreeFile
    
    Open lsFileName For Binary As #liFile
    Put #1, , lyB
    Close #liFile
    
    For i = 1 To 3
        ani.LoadFromFile lsFileName
        If Not ani.AutoPlay Then ani.Play
        ani.StopPlaying
        If KeyIsDown(VK_ESCAPE) Then Exit For
    Next
    
    ani.Timer = Not ani.Timer
    ani.Timer = Not ani.Timer
    Kill lsFileName
    
    If Not (KeyIsDown(VK_ESCAPE) Or InIDE()) Then
        
        For i = 1 To 3
            
            ani.LoadFromResource App.hInstance, 101
            If Not ani.AutoPlay Then ani.Play
            ani.StopPlaying
            If KeyIsDown(VK_ESCAPE) Then Exit For
            
        Next
        
    End If
    
    On Error GoTo 0
End Sub
