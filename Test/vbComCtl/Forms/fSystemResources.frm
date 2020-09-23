VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#112.0#0"; "vbComCtl.ocx"
Begin VB.Form fSystemResources 
   Caption         =   "Resource Usage"
   ClientHeight    =   4950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4125
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin vbComCtl.ucFrame fra 
      Height          =   3615
      Left            =   540
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6376
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Border          =   0   'False
      Begin VB.ListBox lst 
         Height          =   1260
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   13
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   12
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   11
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   10
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   9
         Left            =   1680
         TabIndex        =   8
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Peak Page File Use:"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Page File Use:"
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Peak Working Set:"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Working Set:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Page Faults:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
   Begin vbComCtl.ucListView lvw 
      Height          =   1035
      Left            =   420
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1826
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      StyleEx         =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1740
      Top             =   1380
   End
   Begin vbComCtl.ucTabStrip tabstrip 
      Height          =   3975
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   7011
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuOnTop 
         Caption         =   "On &Top"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "fSystemResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
'fSystemResources.frm      10/9/05
'
'           PURPOSE:
'               Display memory and other resource counters.
'
'==================================================================================================
Option Explicit

'This option enables individual resource counters, such as memory or window handles.  If this switch
'is enabled, the bDebug compiler switch in vbComCtl.vbp must also be enabled in order to compile.
#Const bDebug = 1

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As Any, ByVal cb As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const PROCESS_QUERY_INFORMATION As Long = &H400

Private Sub Form_Load()
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Load the listview columns and tabstrips tabs.
'---------------------------------------------------------------------------------------
    #If bDebug Then
        pUpdateList
        lvw.Columns.Item(OneL).Sort
    #End If
    
    With tabstrip
        #If bDebug Then
            .Tabs.Add "Handles"
        #End If
        .Tabs.Add "Memory"
        .SetSelectedTab OneL
    End With
    
    pUpdateMemory
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Visible = False
    End If
End Sub

Private Sub Form_Resize()
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Move the controls into position on the form.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    tabstrip.Move ZeroL, ZeroL, ScaleX(ScaleWidth, ScaleMode, vbTwips), ScaleY(ScaleHeight, ScaleMode, vbTwips)
    tabstrip.MoveToClient IIf(tabstrip.SelectedTab.Text = "Handles", lvw, fra)
    On Error GoTo 0
End Sub

Private Sub mnuClose_Click()
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Unload me.
'---------------------------------------------------------------------------------------
    Visible = False
End Sub

Private Sub mnuOnTop_Click()
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Toggle topmost state.
'---------------------------------------------------------------------------------------
    mnuOnTop.Checked = Not mnuOnTop.Checked
    OnTop hWnd, mnuOnTop.Checked
End Sub

Private Sub lvw_ColumnClick(ByVal oColumn As vbComCtl.cColumn)
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Sort the column that the user clicks.
'---------------------------------------------------------------------------------------
    oColumn.Sort
End Sub

Private Sub tabstrip_Click(ByVal oTab As vbComCtl.cTab)
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Move the control to the client area of the tabstrip.
'---------------------------------------------------------------------------------------
    fra.Visible = False
    lvw.Visible = False
    tabstrip.MoveToClient IIf(oTab.Text = "Handles", lvw, fra)
End Sub

Private Sub fra_Resize()
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Move the listbox into position on the frame.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    lst.Move ZeroL, lst.Top, fra.Width, fra.Height - lst.Top
    On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Update the resource counters.
'---------------------------------------------------------------------------------------
    #If bDebug Then
        pUpdateList
    #End If
    pUpdateMemory
End Sub

#If bDebug Then

    Private Sub pUpdateList()
    '---------------------------------------------------------------------------------------
    ' Date      : 10/9/05
    ' Purpose   : Update the listview to the current counters from vbComCtl.
    '---------------------------------------------------------------------------------------
        Dim x As Long, y As Long
        For y = ZeroL To vbComCtl.DEBUG_GridCountY - OneL
            For x = ZeroL To vbComCtl.DEBUG_GridCountX - OneL
                pUpdateCell x, y
            Next
        Next
    End Sub
    
    Private Sub pUpdateCell(ByVal x As Long, ByVal y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 10/9/05
    ' Purpose   : Update text at a column/row position in the listview.
    '---------------------------------------------------------------------------------------
        Dim lsText As String
        lsText = vbComCtl.DEBUG_Grid(x, y)
        
        If y = ZeroL Then
            If x = lvw.Columns.Count Then
                Dim lbIsNumeric As Boolean: lbIsNumeric = IsNumeric(vbComCtl.DEBUG_Grid(x, OneL))
                lvw.Columns.Add , vbComCtl.DEBUG_Grid(x, y), , IIf(lbIsNumeric, lvwSortNumeric, lvwSortString), IIf(lbIsNumeric, lvwAlignCenter, lvwAlignLeft), ScaleX(IIf(lbIsNumeric, 840, 1415), vbTwips, ScaleMode)
            End If
        Else
            Dim liIndex As Long:            liIndex = ((y * vbComCtl.DEBUG_GridCountX) + x) + OneL
            Dim loItem As cListItem:        Set loItem = lvw.FindItemData(y + 1&)
            If loItem Is Nothing Then Set loItem = lvw.ListItems.Add(, , , , y + OneL)
            If loItem.SubItem(x + OneL).Text <> lsText Then loItem.SubItem(x + OneL) = lsText
        End If
    End Sub
#End If

Private Function pUpdateMemory() As Long
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Update the memory counters.
'---------------------------------------------------------------------------------------
    Dim lpid As Long
    GetWindowThreadProcessId hWnd, lpid
    
    Dim lhProcess As Long
    lhProcess = OpenProcess(PROCESS_QUERY_INFORMATION, ZeroL, lpid)
    
    If lhProcess Then
    
        Dim lt As PROCESS_MEMORY_COUNTERS
        lt.cb = LenB(lt)
        
        GetProcessMemoryInfo lhProcess, lt, LenB(lt)
        
        With lt
            pUpdateLabel lbl(9), .PageFaultCount
            pUpdateLabel lbl(10), .WorkingSetSize, True
            pUpdateLabel lbl(11), .PeakWorkingSetSize, True
            pUpdateLabel lbl(12), .PagefileUsage, True
            pUpdateLabel lbl(13), .PeakPagefileUsage, True
            
            Static liPeakWorkingSetSize As Long
            If .PeakWorkingSetSize > liPeakWorkingSetSize Then
                liPeakWorkingSetSize = .PeakWorkingSetSize
                lst.AddItem Format$(Now, "MM/DD HH:mm:SS AMPM") & ": " & .PeakWorkingSetSize \ 1024 & " KB"
                lst.ListIndex = lst.NewIndex
            End If
            
        End With
        
        CloseHandle lhProcess
    End If

End Function

Private Sub pUpdateLabel(ByVal oLbl As Label, ByVal i As Long, Optional ByVal bKB As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 10/9/05
' Purpose   : Update a label with the given value.
'---------------------------------------------------------------------------------------
    Dim ls As String
    If bKB Then ls = CStr(i / 1024) & " KB" Else ls = CStr(i)
    If oLbl.Caption <> ls Then oLbl.Caption = ls
End Sub

