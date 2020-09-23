VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#95.0#0"; "vbComCtl.ocx"
Begin VB.Form fDev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DevMode"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin vbComCtl.ucFrame fra 
      Height          =   1155
      Index           =   2
      Left            =   60
      Top             =   3480
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2037
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Caption         =   "Paper Source"
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   5
         ItemData        =   "fDev.frx":0000
         Left            =   1380
         List            =   "fDev.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   2415
      End
      Begin VB.OptionButton opt 
         Caption         =   "Userdefined"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1155
      End
      Begin VB.OptionButton opt 
         Caption         =   "Predefined"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   300
         Index           =   1
         Left            =   2205
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Buddy           =   "txt(1)"
         BuddyProp       =   "Text"
      End
   End
   Begin vbComCtl.ucFrame fra 
      Height          =   1155
      Index           =   1
      Left            =   4080
      Top             =   3480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2037
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Caption         =   "Print Quality"
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1320
         TabIndex        =   20
         Top             =   720
         Width           =   915
      End
      Begin VB.OptionButton opt 
         Caption         =   "Predefined"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "Userdefined"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1155
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         ItemData        =   "fDev.frx":0175
         Left            =   1335
         List            =   "fDev.frx":0189
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   300
         Width           =   1155
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   300
         Index           =   4
         Left            =   2220
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Max             =   32767
         Buddy           =   "txt(4)"
         BuddyProp       =   "Text"
      End
   End
   Begin vbComCtl.ucFrame fra 
      Height          =   1335
      Index           =   0
      Left            =   2520
      Top             =   2040
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   2355
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      Caption         =   "Paper Size"
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   2250
         TabIndex        =   12
         Top             =   915
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   2250
         TabIndex        =   11
         Top             =   555
         Width           =   855
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   3
         ItemData        =   "fDev.frx":01AF
         Left            =   1320
         List            =   "fDev.frx":0395
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   180
         Width           =   2775
      End
      Begin VB.OptionButton opt 
         Caption         =   "Userdefined"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   1155
      End
      Begin VB.OptionButton opt 
         Caption         =   "Predefined"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   300
         Index           =   3
         Left            =   3075
         Top             =   915
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Max             =   32767
         Buddy           =   "txt(3)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   300
         Index           =   2
         Left            =   3075
         Top             =   555
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Max             =   32767
         Buddy           =   "txt(2)"
         BuddyProp       =   "Text"
      End
      Begin VB.Label lbl 
         Caption         =   "Height:"
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   26
         Top             =   945
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Width:"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   25
         Top             =   585
         Width           =   615
      End
   End
   Begin vbComCtl.ucUpDown ud 
      Height          =   300
      Index           =   0
      Left            =   6360
      Top             =   1620
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Buddy           =   "txt(0)"
      BuddyProp       =   "Text"
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   1620
      Width           =   1215
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   2
      ItemData        =   "fDev.frx":0D99
      Left            =   2520
      List            =   "fDev.frx":0DA3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   2415
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   0
      ItemData        =   "fDev.frx":0DC0
      Left            =   2520
      List            =   "fDev.frx":0DC2
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "cmbDevMode"
      Top             =   420
      Width           =   2415
   End
   Begin VB.CheckBox chk 
      Caption         =   "Collate"
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Top             =   1020
      Width           =   795
   End
   Begin VB.CheckBox chk 
      Caption         =   "Color"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   1020
      Width           =   675
   End
   Begin VB.ComboBox cmb 
      Height          =   315
      Index           =   4
      ItemData        =   "fDev.frx":0DC4
      Left            =   2520
      List            =   "fDev.frx":0DD1
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1620
      Width           =   2415
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox lst 
      Height          =   3210
      ItemData        =   "fDev.frx":0E09
      Left            =   60
      List            =   "fDev.frx":0E54
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Copies:"
      Height          =   255
      Index           =   6
      Left            =   5220
      TabIndex        =   24
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Orientation:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   23
      Top             =   780
      Width           =   1035
   End
   Begin VB.Label lbl 
      Caption         =   "Device Name:"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   22
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lbl 
      Caption         =   "Duplex:"
      Height          =   255
      Index           =   14
      Left            =   2520
      TabIndex        =   21
      Top             =   1380
      Width           =   1035
   End
End
Attribute VB_Name = "fDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'fDev.frm           3/28/05
'
'            PURPOSE:
'               GUI for editing properties of cDeviceMode.cls.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eChk
    chkCollate
    chkColor
End Enum

Private Enum eCmb
    cmbDevice
    cmbQuality
    cmbOrientation
    cmbPaperSize
    cmbDuplex
    cmbPaperSource
End Enum

Private Enum eUd
    udCopies
    udDefSrc
    udWidth
    udHeight
    udQuality
End Enum

Private Enum eOpt
    optSizePre
    optSizeUser
    optQualityPre
    optQualityUser
    optSourcePre
    optSourceUser
End Enum

Private Sub Form_Load()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the controls.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowAllUIStates hwnd
    
    Dim p As Printer
    For Each p In Printers
        cmb(cmbDevice).AddItem p.DeviceName
    Next
End Sub

Private Sub cmd_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Dismiss the dialog.
'---------------------------------------------------------------------------------------
    Hide
End Sub

Public Sub GetDevMode(ByVal oOwner As Form, ByVal oDevMode As cDeviceMode)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the DEVMODE data, display the form, and return any changes made by the user.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    With oDevMode
        cmb(cmbDevice).Text = .DeviceName
        pSetByItemData cmb(cmbQuality), .PrintQuality
        pSetByItemData cmb(cmbOrientation), .Orientation
        pSetByItemData cmb(cmbPaperSize), .PaperSize
        pSetByItemData cmb(cmbDuplex), .Duplex
        pSetByItemData cmb(cmbPaperSource), .DefaultSource
        
        chk(chkCollate).Value = Abs(.Collate)
        chk(chkColor).Value = Abs(.Color)
        
        ud(udCopies).Value = .Copies
        ud(udDefSrc).Value = .DefaultSource
        ud(udWidth).Value = .PaperWidth
        ud(udHeight).Value = .PaperLength
        ud(udQuality).Value = .PrintQuality
        
        If .PaperSize = dmPaperSizeUserDefined _
            Then opt(optSizePre).Value = True _
            Else opt(optSizeUser).Value = True
        
        If .PrintQuality < 0 _
            Then opt(optQualityPre).Value = True _
            Else opt(optQualityUser).Value = True
        
        If .DefaultSource < dmPaperSourceUser _
            Then opt(optSourcePre).Value = True _
            Else opt(optSourceUser).Value = True
        
        LstFlags(lst) = .Fields
        
        Show vbModal, oOwner
        
        .DeviceName = cmb(cmbDevice).Text
        If cmb(cmbOrientation).ListIndex > NegOneL _
            Then .Orientation = cmb(cmbOrientation).ItemData(cmb(cmbOrientation).ListIndex)
        If cmb(cmbDuplex).ListIndex > NegOneL _
            Then .Duplex = cmb(cmbDuplex).ItemData(cmb(cmbDuplex).ListIndex)
        
        If opt(optSizePre).Value Then
            If cmb(cmbPaperSize).ListIndex > NegOneL _
                Then .PaperSize = cmb(cmbPaperSize).ItemData(cmb(cmbPaperSize).ListIndex)
        Else
            .PaperSize = dmPaperSizeUserDefined
        End If
        
        .Collate = CBool(chk(chkCollate).Value)
        .Color = CBool(chk(chkColor).Value)
        
        .Copies = ud(udCopies).Value
        .PaperWidth = ud(udWidth).Value
        .PaperLength = ud(udHeight).Value
        
        If opt(optQualityPre).Value Then
            If cmb(cmbQuality).ListIndex > NegOneL _
                Then .PrintQuality = cmb(cmbQuality).ItemData(cmb(cmbQuality).ListIndex)
        Else
            .PrintQuality = ud(udQuality).Value
        End If
        
        If opt(optSourcePre).Value Then
            If cmb(cmbPaperSource).ListIndex > NegOneL _
                Then .DefaultSource = cmb(cmbPaperSource).ItemData(cmb(cmbPaperSource).ListIndex)
        Else
            .DefaultSource = ud(udDefSrc).Value
        End If
        
        .Fields = LstFlags(lst)
        
    End With
    On Error GoTo 0
End Sub

Private Sub pSetByItemData(ByVal oCmb As ComboBox, ByVal iItemData As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : set the listindex of the combobox to the first item found with a matching itemdata.
'---------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To oCmb.ListCount - OneL
        If oCmb.ItemData(i) = iItemData Then Exit For
    Next
    If i < oCmb.ListCount Then oCmb.ListIndex = i Else oCmb.ListIndex = NegOneL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Hide the form instead if the user tries to close it.
'---------------------------------------------------------------------------------------
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If
End Sub

Private Sub opt_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : enable or disable the corresponding controls.
'---------------------------------------------------------------------------------------
    Select Case True
    Case Index < optQualityPre
        txt(udWidth).Enabled = CBool(Index = optSizeUser)
        ud(udWidth).Enabled = CBool(Index = optSizeUser)
        txt(udHeight).Enabled = CBool(Index = optSizeUser)
        ud(udHeight).Enabled = CBool(Index = optSizeUser)
        cmb(cmbPaperSize).Enabled = (Index = optSizePre)
    Case Index < optSourcePre
        txt(udQuality).Enabled = CBool(Index = optQualityUser)
        ud(udQuality).Enabled = CBool(Index = optQualityUser)
        cmb(cmbQuality).Enabled = CBool(Index = optQualityPre)
    Case Else
        txt(udDefSrc).Enabled = CBool(Index = optSourceUser)
        ud(udDefSrc).Enabled = CBool(Index = optSourceUser)
        cmb(cmbPaperSource).Enabled = CBool(Index = optSourcePre)
    End Select
End Sub
