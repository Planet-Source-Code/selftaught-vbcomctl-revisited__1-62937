VERSION 5.00
Object = "{6B12211E-22A8-11DA-9002-C6F4D6587ECE}#107.0#0"; "vbComCtl.ocx"
Begin VB.UserControl ucComDlgTest 
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   HasDC           =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   12000
   Begin vbComCtl.ucScrollBox ucScrollBox1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4471
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   2775
         Index           =   7
         Left            =   5520
         ScaleHeight     =   2775
         ScaleWidth      =   2415
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         Begin VB.CommandButton cmdDevMode 
            Caption         =   "DevMode ..."
            Height          =   495
            Left            =   0
            TabIndex        =   117
            Top             =   0
            Width           =   1455
         End
         Begin VB.CheckBox chkDevNames 
            Caption         =   "Default"
            Height          =   255
            Left            =   0
            TabIndex        =   116
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtDevNames 
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   115
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox txtDevNames 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   114
            Top             =   1560
            Width           =   2175
         End
         Begin VB.ComboBox cmbDevNames 
            Height          =   315
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   113
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lbl 
            Caption         =   "Output Port:"
            Height          =   255
            Index           =   38
            Left            =   0
            TabIndex        =   120
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   "Driver Name:"
            Height          =   255
            Index           =   39
            Left            =   0
            TabIndex        =   119
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   "Device Name:"
            Height          =   255
            Index           =   40
            Left            =   0
            TabIndex        =   118
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   1095
         Index           =   6
         Left            =   0
         ScaleHeight     =   1095
         ScaleWidth      =   5415
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   0
         Width           =   5415
         Begin VB.CommandButton cmdShow 
            Caption         =   "Show Dialog"
            Height          =   615
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   2535
         End
         Begin VB.TextBox txtCommon 
            Height          =   285
            Index           =   0
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtCommon 
            Height          =   285
            Index           =   1
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   315
            Width           =   1575
         End
         Begin VB.ComboBox cmbCenter 
            Height          =   315
            ItemData        =   "ucComDlgTest.ctx":0000
            Left            =   840
            List            =   "ucComDlgTest.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Return Value:"
            Height          =   255
            Index           =   36
            Left            =   2760
            TabIndex        =   56
            Top             =   45
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Extended Error:"
            Height          =   255
            Index           =   37
            Left            =   2640
            TabIndex        =   55
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Center To:"
            Height          =   255
            Index           =   13
            Left            =   -60
            TabIndex        =   54
            Top             =   765
            Width           =   855
         End
         Begin VB.Label lbl 
            Caption         =   "(Only if dialog hook is enabled)"
            Height          =   255
            Index           =   14
            Left            =   2565
            TabIndex        =   53
            Top             =   780
            Width           =   2175
         End
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   5295
      Index           =   5
      Left            =   4140
      ScaleHeight     =   5295
      ScaleWidth      =   7695
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ListBox lst 
         Height          =   1410
         Index           =   6
         ItemData        =   "ucComDlgTest.ctx":0028
         Left            =   0
         List            =   "ucComDlgTest.ctx":0088
         Style           =   1  'Checkbox
         TabIndex        =   42
         Top             =   1440
         Width           =   2655
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Pages:"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   46
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "All Pages"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   44
         Top             =   1440
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   315
         Index           =   21
         Left            =   3480
         TabIndex        =   47
         Text            =   "1"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   315
         Index           =   22
         Left            =   4560
         TabIndex        =   49
         Text            =   "5"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   23
         Left            =   3480
         TabIndex        =   48
         Text            =   "1"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   24
         Left            =   4560
         TabIndex        =   50
         Text            =   "5"
         Top             =   2640
         Width           =   375
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Selection"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   45
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox lst 
         Height          =   1410
         Index           =   9
         ItemData        =   "ucComDlgTest.ctx":01FC
         Left            =   0
         List            =   "ucComDlgTest.ctx":025C
         Style           =   1  'Checkbox
         TabIndex        =   43
         Top             =   3240
         Width           =   2655
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   14
         Left            =   4905
         Top             =   2640
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   1
         Max             =   5
         Value           =   5
         Large           =   2
         Buddy           =   "txt(24)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   12
         Left            =   4905
         Top             =   2280
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   1
         Max             =   5
         Value           =   5
         Large           =   2
         Buddy           =   "txt(22)"
         Enabled         =   0   'False
         BuddyProp       =   "Text"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   11
         Left            =   3825
         Top             =   2280
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         Min             =   1
         Max             =   5
         Value           =   1
         Large           =   2
         Buddy           =   "txt(21)"
         Enabled         =   0   'False
         BuddyProp       =   "Text"
      End
      Begin vbComCtlTest.ucEvents evtLog 
         Height          =   2055
         Index           =   5
         Left            =   2760
         TabIndex        =   51
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3625
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   13
         Left            =   3825
         Top             =   2640
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   1
         Max             =   5
         Value           =   1
         Large           =   2
         Buddy           =   "txt(23)"
         BuddyProp       =   "Text"
      End
      Begin VB.Label lbl 
         Caption         =   "To print a test document, select the dlgPrintReturnDC flag."
         Height          =   255
         Index           =   35
         Left            =   0
         TabIndex        =   65
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lbl 
         Caption         =   "Events:"
         Height          =   255
         Index           =   34
         Left            =   2760
         TabIndex        =   64
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         Caption         =   "To:"
         Height          =   255
         Index           =   13
         Left            =   4290
         TabIndex        =   63
         Top             =   2340
         Width           =   315
      End
      Begin VB.Label lblInfo 
         Caption         =   "Min:"
         Height          =   255
         Index           =   14
         Left            =   3135
         TabIndex        =   62
         Top             =   2700
         Width           =   390
      End
      Begin VB.Label lblInfo 
         Caption         =   "Max:"
         Height          =   255
         Index           =   15
         Left            =   4215
         TabIndex        =   61
         Top             =   2670
         Width           =   390
      End
      Begin VB.Label lblInfo 
         Caption         =   "From:"
         Height          =   255
         Index           =   2
         Left            =   3060
         TabIndex        =   60
         Top             =   2310
         Width           =   390
      End
      Begin VB.Label lbl 
         Caption         =   "Return Flags"
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   59
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label lbl 
         Caption         =   "Events require the dlgPrintEnablePrintHook or dlgPrintEnableSetupHook flag depending on the dlgPrintSetup flag."
         Height          =   1095
         Index           =   33
         Left            =   5520
         TabIndex        =   58
         Top             =   3240
         Width           =   2175
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   5295
      Index           =   4
      Left            =   3180
      ScaleHeight     =   5295
      ScaleWidth      =   6975
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   6975
      Begin VB.ListBox lst 
         Height          =   1410
         Index           =   5
         ItemData        =   "ucComDlgTest.ctx":03D0
         Left            =   0
         List            =   "ucComDlgTest.ctx":040E
         Style           =   1  'Checkbox
         TabIndex        =   36
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   13
         Left            =   1560
         TabIndex        =   32
         Text            =   "0.25"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   14
         Left            =   1560
         TabIndex        =   33
         Text            =   "0.25"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   15
         Left            =   1560
         TabIndex        =   34
         Text            =   "0.25"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   16
         Left            =   1560
         TabIndex        =   35
         Text            =   "0.25"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   17
         Left            =   3900
         TabIndex        =   37
         Text            =   "0.10"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   18
         Left            =   3900
         TabIndex        =   38
         Text            =   "0.10"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   19
         Left            =   3900
         TabIndex        =   39
         Text            =   "0.10"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   20
         Left            =   3900
         TabIndex        =   40
         Text            =   "0.10"
         Top             =   2400
         Width           =   495
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   3
         Left            =   2040
         Top             =   1320
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   25
         Small           =   5
         Large           =   25
         Buddy           =   "txt(13)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   4
         Left            =   2040
         Top             =   1680
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   25
         Small           =   5
         Large           =   25
         Buddy           =   "txt(14)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   5
         Left            =   2025
         Top             =   2040
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   25
         Small           =   5
         Large           =   25
         Buddy           =   "txt(15)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   6
         Left            =   2025
         Top             =   2400
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   25
         Small           =   5
         Large           =   25
         Buddy           =   "txt(16)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   7
         Left            =   4365
         Top             =   1320
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   10
         Small           =   5
         Large           =   25
         Buddy           =   "txt(17)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   8
         Left            =   4365
         Top             =   1680
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   10
         Small           =   5
         Large           =   25
         Buddy           =   "txt(18)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   9
         Left            =   4365
         Top             =   2040
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   10
         Small           =   5
         Large           =   25
         Buddy           =   "txt(19)"
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   10
         Left            =   4365
         Top             =   2400
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Min             =   10
         Max             =   500
         Value           =   10
         Small           =   5
         Large           =   25
         Buddy           =   "txt(20)"
      End
      Begin vbComCtlTest.ucEvents evtLog 
         Height          =   2055
         Index           =   4
         Left            =   2760
         TabIndex        =   41
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3625
      End
      Begin VB.Label lbl 
         Caption         =   "Flags:"
         Height          =   255
         Index           =   31
         Left            =   0
         TabIndex        =   76
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label lbl 
         Caption         =   "Events: (requires dlgPPSEnablePageSetupHook)"
         Height          =   255
         Index           =   32
         Left            =   2760
         TabIndex        =   75
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Top Margin:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   74
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Left Margin:"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   73
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Right Margin:"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   72
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Bottom Margin:"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   71
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Top Margin:"
         Height          =   255
         Index           =   7
         Left            =   2580
         TabIndex        =   70
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Left Margin:"
         Height          =   255
         Index           =   8
         Left            =   2580
         TabIndex        =   69
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Right Margin:"
         Height          =   255
         Index           =   9
         Left            =   2460
         TabIndex        =   68
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Bottom Margin:"
         Height          =   255
         Index           =   10
         Left            =   2340
         TabIndex        =   67
         Top             =   2400
         Width           =   1515
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4755
      Index           =   3
      Left            =   2280
      ScaleHeight     =   4755
      ScaleWidth      =   8025
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   8025
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   12
         Left            =   6840
         TabIndex        =   28
         Text            =   "28"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   11
         Left            =   6840
         TabIndex        =   27
         Text            =   "8"
         Top             =   150
         Width           =   855
      End
      Begin VB.PictureBox picFontSample 
         HasDC           =   0   'False
         Height          =   975
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   7935
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1230
         Width           =   7995
      End
      Begin VB.ListBox lst 
         Height          =   2085
         Index           =   4
         ItemData        =   "ucComDlgTest.ctx":0512
         Left            =   0
         List            =   "ucComDlgTest.ctx":058E
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   2640
         Width           =   2415
      End
      Begin VB.ListBox lst 
         Height          =   2085
         Index           =   8
         ItemData        =   "ucComDlgTest.ctx":06E9
         Left            =   2520
         List            =   "ucComDlgTest.ctx":0765
         Style           =   1  'Checkbox
         TabIndex        =   30
         Top             =   2640
         Width           =   2415
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   1
         Left            =   7680
         Top             =   150
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Value           =   6
         BProps          =   327718
         Buddy           =   "txt(11)"
         BuddyProp       =   "Text"
      End
      Begin vbComCtlTest.ucEvents evtLog 
         Height          =   2055
         Index           =   3
         Left            =   5040
         TabIndex        =   31
         Top             =   2640
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3625
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   315
         Index           =   2
         Left            =   7665
         Top             =   600
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Value           =   72
         BProps          =   327718
         Buddy           =   "txt(12)"
         BuddyProp       =   "Text"
      End
      Begin VB.Label lbl 
         Caption         =   "Flags:"
         Height          =   255
         Index           =   30
         Left            =   0
         TabIndex        =   83
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Events: (Must have dlgFontEnableHook)"
         Height          =   255
         Index           =   29
         Left            =   5040
         TabIndex        =   82
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lbl 
         Caption         =   "Min Size:"
         Height          =   255
         Index           =   28
         Left            =   6120
         TabIndex        =   81
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Max Size:"
         Height          =   255
         Index           =   27
         Left            =   6120
         TabIndex        =   80
         Top             =   630
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Return Flags:"
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   79
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   5295
      Index           =   2
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   7815
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   7
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   8
         Left            =   1200
         TabIndex        =   24
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   9
         Left            =   1200
         TabIndex        =   23
         Top             =   2160
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   10
         Left            =   1200
         TabIndex        =   22
         Text            =   "Choose a folder"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.ListBox lst 
         Height          =   1410
         Index           =   3
         ItemData        =   "ucComDlgTest.ctx":08C0
         Left            =   120
         List            =   "ucComDlgTest.ctx":08E7
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   3240
         Width           =   2655
      End
      Begin vbComCtlTest.ucEvents evtLog 
         Height          =   2055
         Index           =   2
         Left            =   3120
         TabIndex        =   26
         Top             =   3240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
      End
      Begin VB.Label lbl 
         Caption         =   "Events: (requires dlgFolderEnableHook)"
         Height          =   255
         Index           =   20
         Left            =   3120
         TabIndex        =   91
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label lbl 
         Caption         =   "Flags:"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   90
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label lbl 
         Caption         =   "Return Path:"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   89
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Initial Folder:"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   88
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Root Folder:"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   87
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Title:"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   86
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "(requires dlgFolderEnableHook)"
         Height          =   255
         Index           =   26
         Left            =   5475
         TabIndex        =   85
         Top             =   2565
         Width           =   2895
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   5415
      Index           =   1
      Left            =   600
      ScaleHeight     =   5415
      ScaleWidth      =   8055
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   6
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optFile 
         Caption         =   "Save"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optFile 
         Caption         =   "Open"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ListBox lst 
         Height          =   840
         Index           =   2
         ItemData        =   "ucComDlgTest.ctx":09A0
         Left            =   5760
         List            =   "ucComDlgTest.ctx":09A2
         TabIndex        =   19
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   0
         TabIndex        =   15
         Text            =   "Choose a File"
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   0
         TabIndex        =   14
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Text            =   "ChooseThisFile.vbp"
         Top             =   3600
         Width           =   2655
      End
      Begin VB.ListBox lst 
         Height          =   1410
         Index           =   1
         ItemData        =   "ucComDlgTest.ctx":09A4
         Left            =   2880
         List            =   "ucComDlgTest.ctx":0A14
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Text            =   "All Files (*.*)|*.*|VB Projects (*.vbp)|*.vbp"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Text            =   "vbp"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox lst 
         Height          =   1410
         Index           =   7
         ItemData        =   "ucComDlgTest.ctx":0BB1
         Left            =   2880
         List            =   "ucComDlgTest.ctx":0C21
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   3600
         Width           =   2655
      End
      Begin vbComCtlTest.ucEvents evtLog 
         Height          =   2055
         Index           =   1
         Left            =   5760
         TabIndex        =   20
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
      End
      Begin vbComCtl.ucUpDown ud 
         Height          =   285
         Index           =   0
         Left            =   960
         Top             =   3000
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         Max             =   5
         BProps          =   327718
         Large           =   3
         Buddy           =   "txt(2)"
         BuddyProp       =   "Text"
      End
      Begin VB.Label lbl 
         Caption         =   "Level 2: +dlgFileExplorerStyle"
         Height          =   255
         Index           =   19
         Left            =   5760
         TabIndex        =   107
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lbl 
         Caption         =   "Level 1: +dlgFileEnableHook"
         Height          =   255
         Index           =   18
         Left            =   5760
         TabIndex        =   106
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "Events:"
         Height          =   255
         Index           =   17
         Left            =   5760
         TabIndex        =   105
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lbl 
         Caption         =   "Folder Returned:"
         Height          =   255
         Index           =   16
         Left            =   5760
         TabIndex        =   104
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblReturnFilterIndex 
         Height          =   255
         Left            =   7200
         TabIndex        =   103
         Top             =   2325
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Return Filter Index:"
         Height          =   255
         Index           =   9
         Left            =   5760
         TabIndex        =   102
         Top             =   2325
         Width           =   1335
      End
      Begin VB.Label lbl 
         Caption         =   "Files Returned:"
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   101
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Title:"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   100
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Initial Path:"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   99
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Initial File Name:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   98
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Flags:"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   97
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Default Index:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   96
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Filter String:                      Separator: |"
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   95
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label lbl 
         Caption         =   "Default Extension:"
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   94
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Return Flags:"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   93
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4455
      Index           =   0
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   6615
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox picColorSample 
         HasDC           =   0   'False
         Height          =   555
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   6495
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   6555
      End
      Begin VB.ListBox lst 
         Height          =   1185
         Index           =   0
         ItemData        =   "ucComDlgTest.ctx":0DBE
         Left            =   0
         List            =   "ucComDlgTest.ctx":0DD6
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin vbComCtlTest.ucEvents evtLog 
         Height          =   2055
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3625
      End
      Begin VB.Label lblInfo 
         Caption         =   "Color:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   111
         Top             =   1200
         Width           =   6555
      End
      Begin VB.Label lbl 
         Caption         =   "Events: (Must have dlgColorEnableHook)"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   110
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label lbl 
         Caption         =   "Flags:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   109
         Top             =   2160
         Width           =   1815
      End
   End
   Begin vbComCtl.ucTabStrip tabstrip 
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BeginProperty Font {6B122147-22A8-11DA-9002-C6F4D6587ECE} 
      EndProperty
      RightJustify    =   -1  'True
   End
   Begin vbComCtl.ucComDlg dlg 
      Left            =   7800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "ucComDlgTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'ucComDlgTest.ctl           3/31/05
'
'            PURPOSE:
'               Provide functionality for testing the common dialogs.
'
'---------------------------------------------------------------------------------------

Option Explicit

#Const bUseArgs = False

'pictureboxes that contain different sets of controls
Private Enum ePic
    picColor
    picFile
    picFolder
    picFont
    picPageSetup
    picPrint
    picCommon   'shown for all dialogs
    picDev      'shown for print and page setup dialogs
End Enum

Private Enum eOptFile   'File option buttons
    optOpen
    optSave
End Enum

Private Enum eOptPrint  'Print option buttons
    optAll
    optSelection
    optPages
End Enum

Private Enum eEvtLog    'event log usercontrols for each dialog
    evtColor
    evtFile
    evtFolder
    evtFont
    evtPageSetup
    evtPrint
End Enum

Private Enum eLst       'listbox controls
    lstColorFlags
    lstFileFlags
    lstFileReturnFiles
    lstFolderFlags
    lstFontFlags
    lstPageSetupFlags
    lstPrintFlags
    lstFileReturnFlags
    lstFontReturnFlags
    lstPrintReturnFlags
End Enum

Private Enum eTxt       'textbox controls
    txtFileDefExt
    txtFileFilter
    txtFileDefFilterIndex
    txtFileInitFileName
    txtFileInitPath
    txtFileTitle
    txtFileReturnPath
    txtFolderReturnPath
    txtFolderInitial
    txtFolderRoot
    txtFolderTitle
    txtFontMinSize
    txtFontMaxSize
    txtPageSetupTopMargin
    txtPageSetupLeftMargin
    txtPageSetupRightMargin
    txtPageSetupBottomMargin
    txtPageSetupMinTopMargin
    txtPageSetupMinLeftMargin
    txtPageSetupMinRightMargin
    txtPageSetupMinBottomMargin
    txtPrintFromPage
    txtPrintToPage
    txtPrintMinPage
    txtPrintMaxPage
End Enum

Private Enum eUd        'updown controls
    udFileFilterIndex
    udFontMinSize
    udFontMaxSize
    udPageSetupTopMargin
    udPageSetupLeftMargin
    udPageSetupRightMargin
    udPageSetupBottomMargin
    udPageSetupMinTopMargin
    udPageSetupMinLeftMargin
    udPageSetupMinRightMargin
    udPageSetupMinBottomMargin
    udPrintFromPage
    udPrintToPage
    udPrintMinPage
    udPrintMaxPage
End Enum

Private Type DOCINFO
    cbSize As Long
    lpszDocName As String
    lpszOutput As String
End Type

'functions for printing a test page.
Private Declare Function StartDoc Lib "gdi32.dll" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Private Declare Function EndDoc Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StartPage Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPage Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private moPrintDevMode      As cDeviceMode
Private moPageSetupDevMode  As cDeviceMode

Private moPrintDevNames     As cDeviceNames
Private moPageSetupDevNames As cDeviceNames

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
        picColorSample.BackColor = dlg.Color
    End Select
End Sub

Private Sub Usercontrol_EnterFocus()
    vbComCtl.EnterFocus Controls
End Sub

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the contituent controls and modular variables.
'---------------------------------------------------------------------------------------
    vbComCtl.ShowAllUIStates hWnd
    lst(lstFontFlags).Selected(0) = True
    
    With tabstrip.Tabs
        .Add "Color"
        .Add "File"
        .Add "Folder"
        .Add "Font"
        .Add "Page Setup"
        .Add "Print"
    End With
    
    tabstrip.SetSelectedTab OneL
    
    Dim p As Printer
    For Each p In Printers
        cmbDevNames.AddItem p.DeviceName
    Next
    
    Set moPrintDevMode = New cDeviceMode
    Set moPrintDevNames = New cDeviceNames
    Set moPageSetupDevMode = New cDeviceMode
    Set moPageSetupDevNames = New cDeviceNames
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Initialize the font and back color properties to match the container.
'---------------------------------------------------------------------------------------
    Set UserControl.Font = Ambient.Font
    UserControl.BackColor = Ambient.BackColor
    vbComCtl.CascadeBackColor UserControl.Controls, Ambient.BackColor
    picColorSample.BackColor = dlg.Color
End Sub

Private Sub UserControl_Resize()
    tabstrip.Move 0, 0, Width, Height
    tabstrip.MoveToClient ucScrollBox1
End Sub

Private Sub dlg_ColorClose(ByVal hDlg As Long)
    evtLog(evtColor).LogItem "Close"
End Sub

Private Sub dlg_ColorInit(ByVal hDlg As Long)
    pCenter hDlg
    evtLog(evtColor).LogItem "Init "
End Sub

Private Sub dlg_ColorOK(ByVal hDlg As Long, bCancel As stdole.OLE_CANCELBOOL)
    bCancel = CBool(MsgBox("Choose this color?", vbYesNo Or vbDefaultButton1) = vbNo)
End Sub

Private Sub dlg_ColorWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)
    evtLog(evtColor).LogItem "WM_COMMAND " & Hex(wParam)
End Sub

Private Sub dlg_FileChangeFile(ByVal hDlg As Long)
    evtLog(evtFile).LogItem "ChangeFile"
End Sub

Private Sub dlg_FileChangeFolder(ByVal hDlg As Long)
    evtLog(evtFile).LogItem "ChangeFolder"
End Sub

Private Sub dlg_FileChangeType(ByVal hDlg As Long)
    evtLog(evtFile).LogItem "ChangeType"
End Sub

Private Sub dlg_FileClose(ByVal hDlg As Long)
    evtLog(evtFile).LogItem "Close"
End Sub

Private Sub dlg_FileHelpClicked(ByVal hDlg As Long)
    evtLog(evtFile).LogItem "HelpClicked"
End Sub

Private Sub dlg_FileInit(ByVal hDlg As Long)
    pCenter hDlg
    evtLog(evtFile).LogItem "Init"
End Sub

Private Sub dlg_FileOK(ByVal hDlg As Long, bCancel As stdole.OLE_CANCELBOOL)
    bCancel = (MsgBox("Choose this file?", vbYesNo Or vbDefaultButton1) = vbNo)
End Sub

Private Sub dlg_FileWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)
    evtLog(evtFile).LogItem "WM_COMMAND " & Hex(wParam)
End Sub

Private Sub dlg_FolderChanged(ByVal hDlg As Long, ByVal sPath As String)
    evtLog(evtFolder).LogItem "Changed " & sPath
End Sub

Private Sub dlg_FolderInit(ByVal hDlg As Long)
    pCenter hDlg
    evtLog(evtFolder).LogItem "Init"
End Sub

Private Sub dlg_FolderValidationFailed(ByVal hDlg As Long, ByVal sPath As String, bCloseDialog As Boolean)
    bCloseDialog = CBool(MsgBox(sPath & vbNewLine & _
                        "Validation Failed!" & vbNewLine & vbNewLine & _
                        "Keep the window open?", vbYesNo + vbDefaultButton1) = vbNo)
End Sub

Private Sub dlg_FontClose(ByVal hDlg As Long)
    evtLog(evtFont).LogItem "Close"
End Sub

Private Sub dlg_FontInit(ByVal hDlg As Long)
    pCenter hDlg
    evtLog(evtFont).LogItem "Init"
End Sub

Private Sub dlg_FontWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)
    evtLog(evtFont).LogItem "WM_COMMAND " & Hex(wParam)
End Sub

Private Sub dlg_PageSetupClose(ByVal hDlg As Long)
    evtLog(evtPageSetup).LogItem "Close"
End Sub

Private Sub dlg_PageSetupInit(ByVal hDlg As Long)
    pCenter hDlg
    evtLog(evtPageSetup).LogItem "Init"
End Sub

Private Sub dlg_PageSetupWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)
    evtLog(evtPageSetup).LogItem "WM_COMMAND " & Hex(wParam)
End Sub

Private Sub dlg_PrintClose(ByVal hDlg As Long)
    evtLog(evtPrint).LogItem "Close"
End Sub

Private Sub dlg_PrintInit(ByVal hDlg As Long)
    pCenter hDlg
    evtLog(evtPrint).LogItem "Init"
End Sub

Private Sub dlg_PrintWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)
    evtLog(evtPrint).LogItem "WM_COMMAND " & Hex(wParam)
End Sub

Private Sub cmdShow_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show one of the common dialogs in one of two ways.  If the bArgs compiler
'             constant is True, use the arguments of the Showxxxxx method.  Otherwise,
'             use the properties of the ucComDlg object.
'---------------------------------------------------------------------------------------
    Dim lbVal As Boolean
    Dim liReturnExtendedError As Long
    Dim liReturnFlags As Long
    
    Select Case tabstrip.SelectedTab.Index - OneL
    Case picColor
        Dim liColor As Long
        
        #If bUseArgs Then
            liColor = picColorSample.BackColor
            lbVal = dlg.ShowColor(liColor, LstFlags(lst(lstColorFlags)), liReturnExtendedError)
        #Else
            dlg.Color = picColorSample.BackColor
            dlg.ColorFlags = LstFlags(lst(lstColorFlags))
            lbVal = dlg.ShowColor()
            liColor = dlg.Color
            liReturnExtendedError = dlg.ExtendedError
        #End If
        
        If lbVal Then picColorSample.BackColor = liColor
        
    Case picFile
        Dim lsFile As String
        Dim liReturnFilterIndex As Long
        
        #If bUseArgs Then
            If optFile(optOpen).Value _
                Then lbVal = dlg.ShowFileOpen(lsFile, LstFlags(lst(lstFileFlags)), txt(txtFileFilter), ud(udFileFilterIndex).Value, txt(txtFileDefExt), txt(txtFileInitPath), txt(txtFileInitFileName), txt(txtFileTitle), , , liReturnFlags, liReturnExtendedError, liReturnFilterIndex) _
                Else lbVal = dlg.ShowFileSave(lsFile, LstFlags(lst(lstFileFlags)), txt(txtFileFilter), ud(udFileFilterIndex).Value, txt(txtFileDefExt), txt(txtFileInitPath), txt(txtFileInitFileName), txt(txtFileTitle), , , liReturnFlags, liReturnExtendedError, liReturnFilterIndex)
            Debug.Assert lsFile = dlg.FilePath
        #Else
            With dlg
                .FileFlags = LstFlags(lst(lstFileFlags))
                .FileFilter = txt(txtFileFilter).Text
                .FileFilterIndex = ud(udFileFilterIndex).Value
                .FileDefaultExt = txt(txtFileDefExt).Text
                .FileInitialPath = txt(txtFileInitPath).Text
                .FileInitialFile = txt(txtFileInitFileName).Text
                .FileTitle = txt(txtFileTitle).Text
                
                If optFile(optOpen).Value _
                    Then lbVal = dlg.ShowFileOpen() _
                    Else lbVal = dlg.ShowFileSave()
                
                liReturnFilterIndex = .FileReturnFilterIndex
                liReturnFlags = .FileReturnFlags
                liReturnExtendedError = .ExtendedError
            End With
        #End If
        
        If lbVal Then
            LstFlags(lst(lstFileReturnFlags)) = liReturnFlags
            lblReturnFilterIndex.Caption = liReturnFilterIndex
            
            Dim lsPath As String
            Dim lsFiles() As String
            Dim i As Long
            
            With lst(lstFileReturnFiles)
                .Clear
                For i = 0 To dlg.FileNames(lsPath, lsFiles) - 1
                    .AddItem lsFiles(i)
                Next
            End With
            txt(txtFileReturnPath).Text = lsPath
        Else
            LstFlags(lst(lstFileReturnFlags)) = ZeroL
            lst(lstFileReturnFiles).Clear
            txt(txtFileReturnPath).Text = vbNullString
            lblReturnFilterIndex.Caption = vbNullString
        End If
        
    Case picFolder
        liReturnExtendedError = ZeroL
        
        Dim lsFolder As String
        
        #If bUseArgs Then
            lbVal = dlg.ShowFolder(lsFolder, LstFlags(lst(lstFolderFlags)), txt(txtFolderTitle), txt(txtFolderInitial), txt(txtFolderRoot))
        #Else
            With dlg
                .FolderFlags = LstFlags(lst(lstFolderFlags))
                .FolderInitialPath = txt(txtFolderInitial).Text
                .FolderRootPath = txt(txtFolderRoot).Text
                .FolderTitle = txt(txtFolderTitle).Text
                lbVal = .ShowFolder()
                lsFolder = .FolderPath
            End With
        #End If
        
        If lbVal Then txt(txtFolderReturnPath).Text = lsFolder
        
    Case picFont
        Dim liForeColor As Long: liForeColor = picFontSample.ForeColor
        
        #If bUseArgs Then
            lbVal = dlg.ShowFont(picFontSample.Font, LstFlags(lst(lstFontFlags)), picFontSample.hdc, ud(udFontMinSize).Value, ud(udFontMaxSize).Value, liForeColor, liReturnFlags, liReturnExtendedError)
        #Else
            With dlg
                Set .Font = picFontSample.Font
                .FontFlags = LstFlags(lst(lstFontFlags))
                .FontColor = liForeColor
                .FontHdc = picFontSample.hdc
                .FontMinSize = ud(udFontMinSize).Value
                .FontMaxSize = ud(udFontMaxSize).Value
                lbVal = .ShowFont()
                liReturnFlags = .FontReturnFlags
                liReturnExtendedError = .ExtendedError
            End With
        #End If
        
        If lbVal Then
            LstFlags(lst(lstFontReturnFlags)) = liReturnFlags
            picFontSample.ForeColor = liForeColor
            picFontSample.Refresh
        Else
            LstFlags(lst(lstFontReturnFlags)) = ZeroL
        End If
        
    Case picPageSetup
        Dim lfLeftMargin As Single
        Dim lfRightMargin As Single
        Dim lfTopMargin As Single
        Dim lfBottomMargin As Single
        
        pSyncDevNames True, moPageSetupDevNames
        
        #If bUseArgs Then
            lfLeftMargin = ud(udPageSetupLeftMargin).Value / 100
            lfRightMargin = ud(udPageSetupRightMargin).Value / 100
            lfTopMargin = ud(udPageSetupTopMargin).Value / 100
            lfBottomMargin = ud(udPageSetupBottomMargin).Value / 100
            lbVal = dlg.ShowPageSetup(dlgPrintInches, lfLeftMargin, lfRightMargin, lfTopMargin, lfBottomMargin, _
                                 LstFlags(lst(lstPageSetupFlags)), _
                                 ud(udPageSetupMinLeftMargin).Value / 100, _
                                 ud(udPageSetupMinRightMargin).Value / 100, _
                                 ud(udPageSetupMinTopMargin).Value / 100, _
                                 ud(udPageSetupMinBottomMargin).Value / 100, _
                                 moPageSetupDevMode, moPageSetupDevNames, _
                                 liReturnExtendedError)
        #Else
            With dlg
                .PageSetupFlags = LstFlags(lst(lstPageSetupFlags))
                .PageSetupUnits = dlgPrintInches
                .PageSetupLeftMargin = ud(udPageSetupLeftMargin).Value / 100
                .PageSetupRightMargin = ud(udPageSetupRightMargin).Value / 100
                .PageSetupTopMargin = ud(udPageSetupTopMargin).Value / 100
                .PageSetupBottomMargin = ud(udPageSetupBottomMargin).Value / 100
                .PageSetupMinLeftMargin = ud(udPageSetupMinLeftMargin).Value / 100
                .PageSetupMinRightMargin = ud(udPageSetupMinRightMargin).Value / 100
                .PageSetupMinTopMargin = ud(udPageSetupMinTopMargin).Value / 100
                .PageSetupMinBottomMargin = ud(udPageSetupMinBottomMargin).Value / 100
                Set .PageSetupDeviceMode = moPageSetupDevMode
                Set .PageSetupDeviceNames = moPageSetupDevNames
                lbVal = .ShowPageSetup()
                lfLeftMargin = .PageSetupLeftMargin
                lfRightMargin = .PageSetupRightMargin
                lfTopMargin = .PageSetupTopMargin
                lfBottomMargin = .PageSetupBottomMargin
                liReturnExtendedError = .ExtendedError
            End With
        #End If
        
        If lbVal Then
            ud(udPageSetupLeftMargin).Value = lfLeftMargin * 100
            ud(udPageSetupRightMargin).Value = lfRightMargin * 100
            ud(udPageSetupTopMargin).Value = lfTopMargin * 100
            ud(udPageSetupBottomMargin).Value = lfBottomMargin * 100
            pSyncDevNames False, moPageSetupDevNames
        End If
        
    Case picPrint
        Dim lhDc As Long
        Dim liRange As ePrintRange
        Dim liToPage As Long
        Dim liFromPage As Long
        pSyncDevNames True, moPrintDevNames
        
        #If bUseArgs Then
            
            liFromPage = ud(udPrintFromPage).Value
            liToPage = ud(udPrintToPage).Value
            liRange = Switch(optPrint(optAll).Value, dlgPrintRangeAll, _
                             optPrint(optSelection).Value, dlgPrintRangeSelection, _
                             optPrint(optPages).Value, dlgPrintRangePageNumbers)
            
            lbVal = dlg.ShowPrint(lhDc, LstFlags(lst(lstPrintFlags)), _
                             Switch(optPrint(optAll).Value, dlgPrintRangeAll, _
                                     optPrint(optSelection).Value, dlgPrintRangeSelection, _
                                     optPrint(optPages).Value, dlgPrintRangePageNumbers), _
                                     liFromPage, liToPage, _
                                     ud(udPrintMinPage).Value, ud(udPrintMaxPage).Value, _
                                     moPrintDevMode, moPrintDevNames, _
                                     liReturnFlags, liReturnExtendedError)
        #Else
            With dlg
                .PrintFlags = LstFlags(lst(lstPrintFlags))
                .PrintFromPage = ud(udPrintFromPage).Value
                .PrintToPage = ud(udPrintToPage).Value
                .PrintMinPage = ud(udPrintMinPage).Value
                .PrintMaxPage = ud(udPrintMaxPage).Value
                .PrintRange = Switch(optPrint(optAll).Value, dlgPrintRangeAll, _
                                     optPrint(optSelection).Value, dlgPrintRangeSelection, _
                                     optPrint(optPages).Value, dlgPrintRangePageNumbers)
                Set .PrintDeviceMode = moPrintDevMode
                Set .PrintDeviceNames = moPrintDevNames
                lbVal = .ShowPrint()
                liFromPage = .PrintFromPage
                liToPage = .PrintToPage
                lhDc = .PrintHdc
                liRange = .PrintRange
                liReturnFlags = .PrintReturnFlags
                liReturnExtendedError = .ExtendedError
            End With
        #End If
        
        If lbVal Then
            LstFlags(lst(lstPrintReturnFlags)) = liReturnFlags
            ud(udPrintFromPage).Value = liFromPage
            ud(udPrintToPage).Value = liToPage
            optPrint(liRange).Value = True
            If liReturnFlags And dlgPrintReturnDc Then pPrintTest lhDc
            pSyncDevNames False, moPrintDevNames
        Else
            LstFlags(lst(lstPrintReturnFlags)) = ZeroL
        End If
        
    End Select
    
    txtCommon(0) = CStr(lbVal)
    txtCommon(1) = IIf(lbVal, vbNullString, pTranslateExtErr(liReturnExtendedError))
    
End Sub

Private Sub cmdDevMode_Click()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Allow the user to edit the cDeviceMode for either the print or page setup
'             dialog.
'---------------------------------------------------------------------------------------
    Select Case tabstrip.SelectedTab.Index - OneL
    Case picPrint:       fDev.GetDevMode Parent, moPrintDevMode
    Case picPageSetup:   fDev.GetDevMode Parent, moPageSetupDevMode
    End Select
End Sub

Private Sub optPrint_Click(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Enable the Print From and Print To Page selection only if the user has selected
'             to print a page range as opposed to printing all or printing the selection.
'---------------------------------------------------------------------------------------
    txt(txtPrintFromPage).Enabled = (Index = optPages)
    txt(txtPrintToPage).Enabled = (Index = optPages)
    ud(udPrintFromPage).Enabled = (Index = optPages)
    ud(udPrintToPage).Enabled = (Index = optPages)
End Sub

Private Sub picFontSample_Paint()
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show the user some sample text in the font that has been chosen.
'---------------------------------------------------------------------------------------
    picFontSample.CurrentX = 0
    picFontSample.CurrentY = 0
    picFontSample.Print "Sample Text"
End Sub

Private Sub tabstrip_Click(ByVal oTab As vbComCtl.cTab)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Move the ui elements to their position on the control.
'---------------------------------------------------------------------------------------
    Dim loPic As PictureBox
    For Each loPic In pic
        If loPic.Index < picCommon Then
            loPic.Visible = False
            Set loPic.Container = Me
        End If
    Next
    Set loPic = pic(oTab.Index - OneL)
    tabstrip.MoveToClient ucScrollBox1
    loPic.Move -ucScrollBox1.ViewportLeft, -ucScrollBox1.ViewportTop
    Set loPic.Container = ucScrollBox1
    loPic.Visible = True
    Set pic(picDev).Container = IIf((oTab.Index - OneL) = picPrint Or (oTab.Index - OneL) = picPageSetup, ucScrollBox1, Me)
    pic(picDev).Move pic(picCommon).Left + pic(picCommon).Width + 75, pic(picCommon).Top
    ucScrollBox1.AutoSize
    pic(picCommon).ZOrder
    pic(picDev).ZOrder
    pic(picDev).Visible = loPic.Index > picFont
    If loPic.Index = picPageSetup Then
        pSyncDevNames False, moPageSetupDevNames
    ElseIf loPic.Index = picPrint Then
        pSyncDevNames False, moPrintDevNames
    End If
    txtCommon(0).Text = vbNullString
    txtCommon(1).Text = vbNullString
    txtCommon(1).Visible = loPic.Index <> picFolder
    lbl(37).Visible = loPic.Index <> picFolder
End Sub

Private Sub txtDevNames_Change(Index As Integer)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the stored DevNames object.
'---------------------------------------------------------------------------------------
    Select Case tabstrip.SelectedTab.Index - OneL
    Case picPageSetup:  pSyncDevNames True, moPageSetupDevNames
    Case picPrint:      pSyncDevNames True, moPrintDevNames
    End Select
End Sub

Private Sub ud_Change(Index As Integer, ByVal iValue As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Since the updown does not handle decimals, handle the update of the buddy
'             control ourselves and scale the updown's value by 100.
'---------------------------------------------------------------------------------------
    If Index >= udPageSetupTopMargin And Index <= udPageSetupMinBottomMargin Then
        txt(txtPageSetupTopMargin + (Index - udPageSetupTopMargin)).Text = Format$(iValue / 100, "0.00")
    End If
End Sub

Private Sub pCenter(ByVal hDlg As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Center the dialog to the screen or the parent form depending on the user's
'             selection.
'---------------------------------------------------------------------------------------
    If cmbCenter.ListIndex > 0 Then dlg.CenterDialog hDlg, Choose(cmbCenter.ListIndex, ContainerHwnd, Screen)
End Sub

Private Function pTranslateExtErr(ByVal iErr As eComDlgExtendedError) As String
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a textual representation of the extended error.
'---------------------------------------------------------------------------------------
    Select Case iErr
    Case Is < dlgErrGeneralCodes:       Debug.Assert False
    Case dlgErrGeneralCodes:            pTranslateExtErr = "[None]"
    Case dlgErrDialogFailure:           pTranslateExtErr = "Dialog Failure"
    Case dlgErrStructsize:              pTranslateExtErr = "Struct Size"
    Case dlgErrInitialization:          pTranslateExtErr = "Initialization"
    Case dlgErrNoTemplate:              pTranslateExtErr = "No Template"
    Case dlgErrNoHInstance:             pTranslateExtErr = "NoHInstance"
    Case dlgErrLoadStrFailure:          pTranslateExtErr = "LoadStrFailure"
    Case dlgErrFindResFailure:          pTranslateExtErr = "FindResFailure"
    Case dlgErrLoadResFailure:          pTranslateExtErr = "LoadResFilure"
    Case dlgErrLockResFailure:          pTranslateExtErr = "LockResFailure"
    Case dlgErrMemAllocFailure:         pTranslateExtErr = "MemAllocFailure"
    Case dlgErrMemlockFailure:          pTranslateExtErr = "MemLockFailure"
    Case dlgErrNoHook:                  pTranslateExtErr = "No Hook"
    Case dlgErrRegisterMsgFail:         pTranslateExtErr = "Register Msg Failure"
    
    Case Is < dlgPrintErrCodes:         pTranslateExtErr = "Unknown General (" & iErr & ")"
    Case dlgPrintErrCodes:              Debug.Assert False
    Case dlgPrintErrSetupFailure:       pTranslateExtErr = "Setup Failure"
    Case dlgPrintErrParseFailure:       pTranslateExtErr = "Parse Failure"
    Case dlgPrintErrRetDefFailure:      pTranslateExtErr = "Ret Def Failure"
    Case dlgPrintErrLoadDrvFailure:     pTranslateExtErr = "Load Drv Failure"
    Case dlgPrintErrGetDevModeFail:     pTranslateExtErr = "Get DevMode Failure"
    Case dlgPrintErrInitFailure:        pTranslateExtErr = "Init Failure"
    Case dlgPrintErrNoDevices:          pTranslateExtErr = "No Devices"
    Case dlgPrintErrNoDefaultPrn:       pTranslateExtErr = "No Default Printer"
    Case dlgPrintErrDNDMMismatch:       pTranslateExtErr = "DN-DM Mismatch"
    Case dlgPrintErrCreateICFailure:    pTranslateExtErr = "Create IC Failure"
    Case dlgPrintErrPrinterNotFound:    pTranslateExtErr = "Printer Not Found"
    Case dlgPrintErrDefaultDifferent:   pTranslateExtErr = "Default Different"
    
    Case Is < dlgFontErrCodes:          pTranslateExtErr = "Unknown Printer (" & iErr & ")"
    Case dlgFontErrCodes:               Debug.Assert False
    Case dlgFontErrNoFonts:             pTranslateExtErr = "No Fonts"
    Case dlgFontErrMaxLessThanMin:      pTranslateExtErr = "Max < Min"
    
    Case Is < dlgFileErrCodes:          pTranslateExtErr = "Unknown Font (" & iErr & ")"
    Case dlgFileErrCodes:               Debug.Assert False
    Case dlgFileErrSubclassFailure:     pTranslateExtErr = "Subclass Failure"
    Case dlgFileErrInvalidFilename:     pTranslateExtErr = "Invalid Filename"
    Case dlgFileErrBufferTooSmall:      pTranslateExtErr = "Buffer Too Small"
    
    Case Is < dlgColorErrCodes:         pTranslateExtErr = "Unknown File (" & iErr & ")"
    
    Case Else:                          pTranslateExtErr = "Unknown (" & iErr & ")"
    End Select
End Function

Public Property Let Themeable(ByVal bNew As Boolean)
    vbComCtl.ThemeControls Controls, bNew
End Property

Private Sub pPrintTest(ByVal hdc As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Print a test document.
'---------------------------------------------------------------------------------------
    If hdc Then
        Dim ltDI As DOCINFO
        ltDI.cbSize = LenB(ltDI)
        ltDI.lpszDocName = "vbComCtl Test Page"
        StartDoc hdc, ltDI
        StartPage hdc
        TextOut hdc, 100, 100, "Testing...", 10
        EndPage hdc
        EndDoc hdc
        DeleteDC hdc
    End If
End Sub

Private Sub pSyncDevNames(ByVal bFromUI As Boolean, ByVal oDevNames As cDeviceNames)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Update the cDeviceNames object from the ui or update the ui from the object.
'---------------------------------------------------------------------------------------
    
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    
    bInHere = True
    
    With oDevNames
        If bFromUI Then
            .Default = chkDevNames.Value
            .DeviceName = cmbDevNames.Text
            .DriverName = txtDevNames(0).Text
            .OutputPort = txtDevNames(1).Text
        Else
            chkDevNames.Value = Abs(.Default)
            cmbDevNames.Text = .DeviceName
            txtDevNames(0).Text = .DriverName
            txtDevNames(1).Text = .OutputPort
        End If
    End With
    
    bInHere = False
    
End Sub

Private Sub chkDevNames_Click()
    Select Case tabstrip.SelectedTab.Index - OneL
    Case picPageSetup:  pSyncDevNames True, moPageSetupDevNames
    Case picPrint:      pSyncDevNames True, moPrintDevNames
    End Select
End Sub

Private Sub cmbDevNames_Click()
    Select Case tabstrip.SelectedTab.Index - OneL
    Case picPageSetup:  pSyncDevNames True, moPageSetupDevNames
    Case picPrint:      pSyncDevNames True, moPrintDevNames
    End Select
End Sub

Private Sub cmbDevNames_Change()
    Select Case tabstrip.SelectedTab.Index - OneL
    Case picPageSetup:  pSyncDevNames True, moPageSetupDevNames
    Case picPrint:      pSyncDevNames True, moPrintDevNames
    End Select
End Sub
