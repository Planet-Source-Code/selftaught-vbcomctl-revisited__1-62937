VERSION 5.00
Begin VB.UserControl ucComDlg 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucComDlg.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "ucComDlg.ctx":0972
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucComDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'ucComDlg.ctl           9/10/05
'
'            PURPOSE:
'               Expose all modal win32 common dialogs, shell browse for folder and html help.
'
'            LINEAGE:
'               See mCommonDialog.bas
'
'---------------------------------------------------------------------------------------

Option Explicit

Implements iComDlgHook

Public Enum eComDlgExtendedError
    dlgErrDialogFailure = CDERR_DIALOGFAILURE

    dlgErrGeneralCodes = CDERR_GENERALCODES
    dlgErrStructsize = CDERR_STRUCTSIZE
    dlgErrInitialization = CDERR_INITIALIZATION
    dlgErrNoTemplate = CDERR_NOTEMPLATE
    dlgErrNoHInstance = CDERR_NOHINSTANCE
    dlgErrLoadStrFailure = CDERR_LOADSTRFAILURE
    dlgErrFindResFailure = CDERR_FINDRESFAILURE
    dlgErrLoadResFailure = CDERR_LOADRESFAILURE
    dlgErrLockResFailure = CDERR_LOCKRESFAILURE
    dlgErrMemAllocFailure = CDERR_MEMALLOCFAILURE
    dlgErrMemlockFailure = CDERR_MEMLOCKFAILURE
    dlgErrNoHook = CDERR_NOHOOK
    dlgErrRegisterMsgFail = CDERR_REGISTERMSGFAIL

    dlgPrintErrCodes = PDERR_PRINTERCODES
    dlgPrintErrSetupFailure = PDERR_SETUPFAILURE
    dlgPrintErrParseFailure = PDERR_PARSEFAILURE
    dlgPrintErrRetDefFailure = PDERR_RETDEFFAILURE
    dlgPrintErrLoadDrvFailure = PDERR_LOADDRVFAILURE
    dlgPrintErrGetDevModeFail = PDERR_GETDEVMODEFAIL
    dlgPrintErrInitFailure = PDERR_INITFAILURE
    dlgPrintErrNoDevices = PDERR_NODEVICES
    dlgPrintErrNoDefaultPrn = PDERR_NODEFAULTPRN
    dlgPrintErrDNDMMismatch = PDERR_DNDMMISMATCH
    dlgPrintErrCreateICFailure = PDERR_CREATEICFAILURE
    dlgPrintErrPrinterNotFound = PDERR_PRINTERNOTFOUND
    dlgPrintErrDefaultDifferent = PDERR_DEFAULTDIFFERENT

    dlgFontErrCodes = CFERR_CHOOSEFONTCODES
    dlgFontErrNoFonts = CFERR_NOFONTS
    dlgFontErrMaxLessThanMin = CFERR_MAXLESSTHANMIN

    dlgFileErrCodes = FNERR_FILENAMECODES
    dlgFileErrSubclassFailure = FNERR_SUBCLASSFAILURE
    dlgFileErrInvalidFilename = FNERR_INVALIDFILENAME
    dlgFileErrBufferTooSmall = FNERR_BUFFERTOOSMALL

    dlgColorErrCodes = CCERR_CHOOSECOLORCODES
End Enum

Public Enum eFileDialog
    dlgFileExplorerStyle = OFN_EXPLORER
    dlgFileMustExist = OFN_FILEMUSTEXIST
    dlgFilePathMustExist = OFN_PATHMUSTEXIST
    dlgFileMultiSelect = OFN_ALLOWMULTISELECT
    dlgFilePromptToCreate = OFN_CREATEPROMPT
    dlgFileEnableSizing = OFN_ENABLESIZING
    dlgFileNoDereferenceLinks = OFN_NODEREFERENCELINKS
    dlgFileHideNetworkButton = OFN_NONETWORKBUTTON
    dlgFileHideReadOnly = OFN_HIDEREADONLY
    dlgFileNoReadOnlyReturn = OFN_NOREADONLYRETURN
    dlgFileNoTestFileCreate = OFN_NOTESTFILECREATE
    dlgFilePromptToOverwrite = OFN_OVERWRITEPROMPT
    dlgFileReadOnly = OFN_READONLY
    dlgFileShowHelpButton = OFN_SHOWHELP
    dlgFileEnableHook = OFN_ENABLEHOOK
    dlgFileEnableTemplate = OFN_ENABLETEMPLATE
    dlgFileDontAddToRecent = OFN_DONTADDTORECENT
    dlgFileExtensionDifferent = OFN_EXTENSIONDIFFERENT
End Enum

Public Enum eColorDialog
    dlgColorFullOpen = CC_FULLOPEN
    dlgColorPreventFullOpen = CC_PREVENTFULLOPEN
    dlgColorSolid = CC_SOLIDCOLOR
    dlgColorAny = CC_ANYCOLOR
    dlgColorEnableHook = CC_ENABLEHOOK
End Enum

Public Enum eFontDialog
    dlgFontScreenFonts = CF_SCREENFONTS
    dlgFontPrinterFonts = CF_PRINTERFONTS
    dlgFontBoth = CF_BOTH
    dlgFontEffects = CF_EFFECTS
    dlgFontNoVectorFonts = CF_NOVECTORFONTS
    dlgFontNoSimulations = CF_NOSIMULATIONS
    dlgFontFixedPitchOnly = CF_FIXEDPITCHONLY
    dlgFontWysiwyg = CF_WYSIWYG ' Must also have ScreenFonts And PrinterFonts
    dlgFontForceExist = CF_FORCEFONTEXIST
    dlgFontScalableOnly = CF_SCALABLEONLY
    dlgFontTTOnly = CF_TTONLY
    dlgFontNoFaceSel = CF_NOFACESEL
    dlgFontNoStyleSel = CF_NOSTYLESEL
    dlgFontNoSizeSel = CF_NOSIZESEL
    
    dlgFontSelectScript = CF_SELECTSCRIPT
    dlgFontNoScriptSel = CF_NOSCRIPTSEL
    dlgFontNoVertFonts = CF_NOVERTFONTS
    
    dlgFontEnableHook = CF_ENABLEHOOK
End Enum

Public Enum ePrintPageSetup
    dlgPPSDefaultMinMargins = PSD_DEFAULTMINMARGINS  ' Default (printer's)
    dlgPPSMinMargins = PSD_MINMARGINS
    dlgPPSMargins = PSD_MARGINS
    dlgPPSDisableMargins = PSD_DISABLEMARGINS
    dlgPPSDisablePrinter = PSD_DISABLEPRINTER
    dlgPPSNoWarning = PSD_NOWARNING
    dlgPPSDisableOrientation = PSD_DISABLEORIENTATION
    dlgPPSReturnDefault = PSD_RETURNDEFAULT
    dlgPPSDisablePaper = PSD_DISABLEPAPER
    dlgPPSShowHelp = PSD_SHOWHELP
    dlgPPSEnablePageSetupHook = PSD_ENABLEPAGESETUPHOOK
    dlgPPSDisablePagePainting = PSD_DISABLEPAGEPAINTING
End Enum

Public Enum ePrintPageSetupUnits
    dlgPrintInches = PSD_UNITS_Inches
    dlgPrintMillimeters = PSD_UNITS_Millimeters
End Enum

Public Enum ePrintDialog
    dlgPrintAllPages = PD_ALLPAGES
    dlgPrintSelection = PD_SELECTION
    dlgPrintPageNums = PD_PAGENUMS
    dlgPrintNoSelection = PD_NOSELECTION
    dlgPrintNoPageNums = PD_NOPAGENUMS
    dlgPrintCollate = PD_COLLATE
    dlgPrintToFile = PD_PRINTTOFILE
    dlgPrintSetup = PD_PRINTSETUP
    dlgPrintNoWarning = PD_NOWARNING
    dlgPrintReturnDc = PD_RETURNDC
    dlgPrintReturnIc = PD_RETURNIC
    dlgPrintReturnDefault = PD_RETURNDEFAULT
    dlgPrintShowHelp = PD_SHOWHELP
    dlgPrintEnablePrintHook = PD_ENABLEPRINTHOOK
    dlgPrintEnableSetupHook = PD_ENABLESETUPHOOK
    dlgPrintDisablePrintToFile = PD_DISABLEPRINTTOFILE
    dlgPrintHidePrintToFile = PD_HIDEPRINTTOFILE
    dlgPrintNoNetworkButton = PD_NONETWORKBUTTON
End Enum

Public Enum ePrintRange
    dlgPrintRangeAll = PD_ALLPAGES
    dlgPrintRangePageNumbers = PD_PAGENUMS
    dlgPrintRangeSelection = PD_SELECTION
End Enum

Public Enum eFolderDialog
    dlgFolderReturnOnlyFSDirs = BIF_RETURNONLYFSDIRS        'Only returns file system directories
    dlgFolderDontGoBelowDomain = BIF_DONTGOBELOWDOMAIN      'Does not include network folders below domain level
    dlgFolderStatusText = BIF_STATUSTEXT                    'Includes status area in the dialog for use with callback
    dlgFolderReturnFSAncestors = BIF_RETURNFSANCESTORS      'Only returns file system ancestors.
    dlgFolderEditBox = BIF_EDITBOX                          'allows user to rename selection
    dlgFolderValidate = BIF_VALIDATE                        'insist on valid editbox result (or CANCEL)
    dlgFolderUseNewUI = BIF_USENEWUI                        'Version 5.0. Use the new user-interface. Setting
                                                            'this flag provides the user with a larger dialog box
                                                            'that can be resized. It has several new capabilities
                                                            'including: drag and drop capability within the
                                                            'dialog box, reordering, context menus, new folders,
                                                            'delete, and other context menu commands. To use
                                                            'this flag, you must call OleInitialize or
                                                            'CoInitialize before calling SHBrowseForFolder.
    'dlgFolderBrowseForComputer = BIF_BROWSEFORCOMPUTER      'Only returns computers.
    'dlgFolderBrowseForPrinter = BIF_BROWSEFORPRINTER        'Only returns printers.
    'dlgFolderBrowseIncludeFiles = BIF_BROWSEINCLUDEFILES    'Browse for everything
    dlgFolderEnableHook = BIF_EnableHook
End Enum

Public Enum eSpecialFolders
    'dlgFolderDesktop = CSIDL_DESKTOP                        '(desktop)
    dlgFolderInternet = CSIDL_INTERNET                      'Internet Explorer (icon on desktop)
    dlgFolderPrograms = CSIDL_PROGRAMS                      'Start Menu\Programs
    dlgFolderControls = CSIDL_CONTROLS                      'My Computer\Control Panel
    dlgFolderPrinters = CSIDL_PRINTERS                      'My Computer\Printers
    dlgFolderPersonal = CSIDL_PERSONAL                      'My Documents
    dlgFolderFavorites = CSIDL_FAVORITES                    '(user name)\Favorites
    dlgFolderStartup = CSIDL_STARTUP                        'Start Menu\Programs\Startup
    dlgFolderRecent = CSIDL_RECENT                          '(user name)\Recent
    dlgFolderSendTo = CSIDL_SENDTO                          '(user name)\SendTo
    dlgFolderBitBucket = CSIDL_BITBUCKET                    '(desktop)\Recycle Bin
    dlgFolderStartMenu = CSIDL_STARTMENU                    '(user name)\Start Menu
    dlgFolderDesktopDirectory = CSIDL_DESKTOPDIRECTORY      '(user name)\Desktop
    'dlgFolderDrives = CSIDL_DRIVES                          'My Computer
    'dlgFolderNetwork = CSIDL_NETWORK                        'Network Neighborhood
    dlgFolderNethood = CSIDL_NETHOOD                        '(user name)\nethood
    dlgFolderFonts = CSIDL_FONTS                            'windows\fonts
    dlgFolderTemplates = CSIDL_TEMPLATES
    dlgFolderCommonStartMenu = CSIDL_COMMON_STARTMENU       'All Users\Start Menu
    dlgFolderCommonPrograms = CSIDL_COMMON_PROGRAMS         'All Users\Programs
    dlgFolderCommonStartup = CSIDL_COMMON_STARTUP           'All Users\Startup
    dlgFolderCommonDesktopDirectory = CSIDL_COMMON_DESKTOPDIRECTORY  'All Users\Desktop
    dlgFolderAppData = CSIDL_APPDATA                        '(user name)\Application Data
    dlgFolderPrinthood = CSIDL_PRINTHOOD                    '(user name)\PrintHood
    dlgFolderLocalAppData = CSIDL_LOCAL_APPDATA             '(user name)\Local Settings\Applicaiton Data (non roaming)
    dlgFolderAltStartup = CSIDL_ALTSTARTUP                  'non localized startup
    dlgFolderCommonAltStartup = CSIDL_COMMON_ALTSTARTUP     'non localized common startup
    dlgFolderCommonFavorites = CSIDL_COMMON_FAVORITES
    dlgFolderInternetCache = CSIDL_INTERNET_CACHE
    dlgFolderCookies = CSIDL_COOKIES
    dlgFolderHistory = CSIDL_HISTORY
    dlgFolderCommonAppData = CSIDL_COMMON_APPDATA           'All Users\Application Data
    dlgFolderWindows = CSIDL_WINDOWS                        'GetWindowsDirectory()
    dlgFolderSystem = CSIDL_SYSTEM                          'GetSystemDirectory()
    dlgFolderProgramFiles = CSIDL_PROGRAM_FILES             'C:\Program Files
    dlgFolderMyPictures = CSIDL_MYPICTURES                  'C:\Program Files\My Pictures
    dlgFolderProfile = CSIDL_PROFILE                        'USERPROFILE
    dlgFolderProgramFilesCommon = CSIDL_PROGRAM_FILES_COMMON 'C:\Program Files\Common
    dlgFolderCommonTemplates = CSIDL_COMMON_TEMPLATES       'All Users\Templates
    dlgFolderCommonDocuments = CSIDL_COMMON_DOCUMENTS       'All Users\Documents
    dlgFolderCommonAdminTools = CSIDL_COMMON_ADMINTOOLS     'All Users\Start Menu\Programs\Administrative Tools
    dlgFolderAdminTools = CSIDL_ADMINTOOLS                  '(user name)\Start Menu\Programs\Administrative Tools
    
    dlgFolderFlagCreate = CSIDL_FLAG_CREATE
    'dlgFolderFlagDontVerify = CSIDL_FLAG_DONT_VERIFY
    dlgFolderFlagMask = CSIDL_FLAG_MASK
End Enum

Public Enum eHelpDialog
    dlgHelpTopic = HH_DISPLAY_TOPIC
    dlgHelpContents = HH_DISPLAY_TOC
    dlgHelpIndex = HH_DISPLAY_INDEX
    dlgHelpSearch = HH_DISPLAY_SEARCH
    dlgHelpContext = HH_HELP_CONTEXT
    dlgHelpCloseAll = HH_CLOSE_ALL
End Enum

Event ColorInit(ByVal hDlg As Long)
Event ColorOK(ByVal hDlg As Long, ByRef bCancel As OLE_CANCELBOOL)
Event ColorClose(ByVal hDlg As Long)
Event ColorWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)

Event FileInit(ByVal hDlg As Long)
Event FileOK(ByVal hDlg As Long, ByRef bCancel As OLE_CANCELBOOL)
Event FileClose(ByVal hDlg As Long)
Event FileChangeFile(ByVal hDlg As Long)
Event FileChangeFolder(ByVal hDlg As Long)
Event FileChangeType(ByVal hDlg As Long)
Event FileHelpClicked(ByVal hDlg As Long)
Event FileWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)

Event FolderInit(ByVal hDlg As Long)
Event FolderChanged(ByVal hDlg As Long, ByVal sPath As String)
Event FolderValidationFailed(ByVal hDlg As Long, ByVal sPath As String, ByRef bCloseDialog As Boolean)

Event FontInit(ByVal hDlg As Long)
Event FontClose(ByVal hDlg As Long)
Event FontWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)

Event PageSetupInit(ByVal hDlg As Long)
Event PageSetupClose(ByVal hDlg As Long)
Event PageSetupWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)

Event PrintInit(ByVal hDlg As Long)
Event PrintClose(ByVal hDlg As Long)
Event PrintWMCommand(ByVal hDlg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lReturn As Long)

Private Const DEF_RaiseCancelErrors As Boolean = False
Private Const DEF_RaiseExtendedErrors As Boolean = False

Private Const DEF_File_Flags                    As Long = OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
Private Const DEF_File_Filter                   As String = "All Files (*.*)" & OFN_FilterDelim & "*.*"
Private Const DEF_File_FilterIndex              As Long = 0
Private Const DEF_File_DefaultExt               As String = vbNullString
Private Const DEF_File_InitPath                 As String = vbNullString
Private Const DEF_File_InitFile                 As String = vbNullString
Private Const DEF_File_Title                    As String = vbNullString

Private Const DEF_Folder_Title                  As String = vbNullString
Private Const DEF_Folder_InitPath               As String = vbNullString
Private Const DEF_Folder_RootPath               As String = vbNullString
Private Const DEF_Folder_Flags                  As Long = BIF_USENEWUI Or BIF_STATUSTEXT Or BIF_RETURNONLYFSDIRS
    
Private Const DEF_Font_Flags                    As Long = CF_SCREENFONTS
Private Const DEF_Font_MinSize                  As Long = 6
Private Const DEF_Font_MaxSize                  As Long = 72
Private Const DEF_Font_Color                    As Long = 0
    
Private Const DEF_Color                         As Long = 0
Private Const DEF_Color_Flags                   As Long = CC_ANYCOLOR
    
Private Const DEF_Print_Flags                   As Long = PD_ALLPAGES
Private Const DEF_Print_Range                   As Long = PD_ALLPAGES
Private Const DEF_Print_MinPage                 As Long = 0
Private Const DEF_Print_MaxPage                 As Long = 0

Private Const DEF_PageSetup_Flags               As Long = PSD_DEFAULTMINMARGINS
Private Const DEF_PageSetup_Units               As Long = PSD_UNITS_Inches
    
Private Const DEF_PageSetup_LeftMargin          As Single = 1
Private Const DEF_PageSetup_RightMargin         As Single = 1
Private Const DEF_PageSetup_TopMargin           As Single = 1
Private Const DEF_PageSetup_BottomMargin        As Single = 1
    
Private Const DEF_PageSetup_MinLeftMargin       As Single = 0
Private Const DEF_PageSetup_MinRightMargin      As Single = 0
Private Const DEF_PageSetup_MinTopMargin        As Single = 0
Private Const DEF_PageSetup_MinBottomMargin     As Single = 0

Private Const DEF_Help_File                     As String = vbNullString


Private Const PROP_RaiseCancelErrors            As String = "RaiseCancel"
Private Const PROP_RaiseExtendedErrors          As String = "RaiseExt"

Private Const PROP_File_Flags                   As String = "FileFlags"
Private Const PROP_File_Filter                  As String = "FileFilter"
Private Const PROP_File_FilterIndex             As String = "FileIndex"
Private Const PROP_File_DefaultExt              As String = "FileExt"
Private Const PROP_File_InitPath                As String = "FilePath"
Private Const PROP_File_InitFile                As String = "FileFile"
Private Const PROP_File_Title                   As String = "FileTitle"

Private Const PROP_Folder_Title                 As String = "FolderTitle"
Private Const PROP_Folder_InitPath              As String = "FolderPath"
Private Const PROP_Folder_RootPath              As String = "FolderRoot"
Private Const PROP_Folder_Flags                 As String = "FolderFlags"
    
Private Const PROP_Font_Flags                   As String = "FontFlags"
Private Const PROP_Font_MinSize                 As String = "FontMinSize"
Private Const PROP_Font_MaxSize                 As String = "FontMaxSize"
Private Const PROP_Font_Color                   As String = "FontColor"
    
Private Const PROP_Color                        As String = "Color"
Private Const PROP_Color_Flags                  As String = "ColorFlags"
    
Private Const PROP_Print_Flags                  As String = "PrintFlags"
Private Const PROP_Print_Range                  As String = "PrintRange"
Private Const PROP_Print_MinPage                As String = "PrintMinPage"
Private Const PROP_Print_MaxPage                As String = "PrintMaxPage"

Private Const PROP_PageSetup_Flags              As String = "PSFlags"
Private Const PROP_PageSetup_Units              As String = "PSUnits"
    
Private Const PROP_PageSetup_LeftMargin         As String = "PSLeft"
Private Const PROP_PageSetup_RightMargin        As String = "PSRight"
Private Const PROP_PageSetup_TopMargin          As String = "PSTop"
Private Const PROP_PageSetup_BottomMargin       As String = "PSBottom"
    
Private Const PROP_PageSetup_MinLeftMargin      As String = "PSMinLeft"
Private Const PROP_PageSetup_MinRightMargin     As String = "PSMinRight"
Private Const PROP_PageSetup_MinTopMargin       As String = "PSMinTop"
Private Const PROP_PageSetup_MinBottomMargin    As String = "PSMinBottom"

Private Const PROP_Help_File                    As String = "HelpFile"

Private Const ClassName                         As String = "ucComDlg"

Private Const NMHDR_code                        As Long = 8

Private mtFile                  As tFileDialog
Private mtFolder                As tFolderDialog
Private mtFont                  As tFontDialog
Private mtColor                 As tColorDialog
Private mtPrint                 As tPrintDialog
Private mtPageSetup             As tPageSetupDialog

Private msHelpFile              As String
Private mbHelpWasShown          As Boolean

Private miExtendedError         As eComDlgExtendedError

Private mbRaiseExtendedErrors   As Boolean
Private mbRaiseCancelErrors     As Boolean

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Initialize modular variables and load the shell library to prevent crashes
'             at shutdown when linked to CC 6.0.
'---------------------------------------------------------------------------------------
    LoadShellMod
    Dim i As Long
    For i = 0 To 15
        mtColor.iColors(i) = QBColor(i)
    Next
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Initialize property values.
'---------------------------------------------------------------------------------------
    mbRaiseExtendedErrors = DEF_RaiseExtendedErrors
    mbRaiseCancelErrors = DEF_RaiseCancelErrors
    
    With mtFile
        .iFlags = DEF_File_Flags
        .sFilter = DEF_File_Filter
        .iFilterIndex = DEF_File_FilterIndex
        .sDefaultExt = DEF_File_DefaultExt
        .sInitFile = DEF_File_InitPath
        .sInitFile = DEF_File_InitFile
        .sTitle = DEF_File_Title
    End With

    With mtFolder
        .sTitle = DEF_Folder_Title
        .sInitialPath = DEF_Folder_InitPath
        .sRootPath = DEF_Folder_RootPath
        .iFlags = DEF_Folder_Flags
    End With

    With mtFont
        .iFlags = DEF_Font_Flags
        .iMinSize = DEF_Font_MinSize
        .iMaxSize = DEF_Font_MaxSize
        .iColor = DEF_Font_Color
    End With

    With mtColor
        .iColor = DEF_Color
        .iFlags = DEF_Color_Flags
    End With

    With mtPrint
        .iFlags = DEF_Print_Flags
        .iRange = DEF_Print_Range
        .iMinPage = DEF_Print_MinPage
        .iMaxPage = DEF_Print_MaxPage
    End With
    
    With mtPageSetup
        .iFlags = DEF_PageSetup_Flags
        .iUnits = DEF_PageSetup_Units
        .fLeftMargin = DEF_PageSetup_LeftMargin
        .fRightMargin = DEF_PageSetup_RightMargin
        .fTopMargin = DEF_PageSetup_TopMargin
        .fBottomMargin = DEF_PageSetup_BottomMargin
        
        .fMinLeftMargin = DEF_PageSetup_MinLeftMargin
        .fMinRightMargin = DEF_PageSetup_MinRightMargin
        .fMinTopMargin = DEF_PageSetup_MinTopMargin
        .fMinBottomMargin = DEF_PageSetup_MinBottomMargin
    End With
    
    msHelpFile = DEF_Help_File
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Read property values from a previously persisted instance.
'---------------------------------------------------------------------------------------
    mbRaiseExtendedErrors = PropBag.ReadProperty(PROP_RaiseExtendedErrors, DEF_RaiseExtendedErrors)
    mbRaiseCancelErrors = PropBag.ReadProperty(PROP_RaiseExtendedErrors, DEF_RaiseCancelErrors)
    
    With mtFile
        .iFlags = PropBag.ReadProperty(PROP_File_Flags, DEF_File_Flags)
        .sFilter = PropBag.ReadProperty(PROP_File_Filter, DEF_File_Filter)
        .iFilterIndex = PropBag.ReadProperty(PROP_File_FilterIndex, DEF_File_FilterIndex)
        .sDefaultExt = PropBag.ReadProperty(PROP_File_DefaultExt, DEF_File_DefaultExt)
        .sInitFile = PropBag.ReadProperty(PROP_File_InitPath, DEF_File_InitPath)
        .sInitFile = PropBag.ReadProperty(PROP_File_InitFile, DEF_File_InitFile)
        .sTitle = PropBag.ReadProperty(PROP_File_Title, DEF_File_Title)
    End With

    With mtFolder
        .sTitle = PropBag.ReadProperty(PROP_Folder_Title, DEF_Folder_Title)
        .sInitialPath = PropBag.ReadProperty(PROP_Folder_InitPath, DEF_Folder_InitPath)
        .sRootPath = PropBag.ReadProperty(PROP_Folder_RootPath, DEF_Folder_RootPath)
        .iFlags = PropBag.ReadProperty(PROP_Folder_Flags, DEF_Folder_Flags)
    End With

    With mtFont
        .iFlags = PropBag.ReadProperty(PROP_Font_Flags, DEF_Font_Flags)
        .iMinSize = PropBag.ReadProperty(PROP_Font_MinSize, DEF_Font_MinSize)
        .iMaxSize = PropBag.ReadProperty(PROP_Font_MaxSize, DEF_Font_MaxSize)
        .iColor = PropBag.ReadProperty(PROP_Font_Color, DEF_Font_Color)
    End With

    With mtColor
        .iColor = PropBag.ReadProperty(PROP_Color, DEF_Color)
        .iFlags = PropBag.ReadProperty(PROP_Color_Flags, DEF_Color_Flags)
    End With

    With mtPrint
        .iFlags = PropBag.ReadProperty(PROP_Print_Flags, DEF_Print_Flags)
        .iRange = PropBag.ReadProperty(PROP_Print_Range, DEF_Print_Range)
        .iMinPage = PropBag.ReadProperty(PROP_Print_MinPage, DEF_Print_MinPage)
        .iMaxPage = PropBag.ReadProperty(PROP_Print_MaxPage, DEF_Print_MaxPage)
    End With
    
    With mtPageSetup
        .iFlags = PropBag.ReadProperty(PROP_PageSetup_Flags, DEF_PageSetup_Flags)
        .iUnits = PropBag.ReadProperty(PROP_PageSetup_Units, DEF_PageSetup_Units)
        .fLeftMargin = PropBag.ReadProperty(PROP_PageSetup_LeftMargin, DEF_PageSetup_LeftMargin)
        .fRightMargin = PropBag.ReadProperty(PROP_PageSetup_RightMargin, DEF_PageSetup_RightMargin)
        .fTopMargin = PropBag.ReadProperty(PROP_PageSetup_TopMargin, DEF_PageSetup_TopMargin)
        .fBottomMargin = PropBag.ReadProperty(PROP_PageSetup_BottomMargin, DEF_PageSetup_BottomMargin)
        
        .fMinLeftMargin = PropBag.ReadProperty(PROP_PageSetup_MinLeftMargin, DEF_PageSetup_MinLeftMargin)
        .fMinRightMargin = PropBag.ReadProperty(PROP_PageSetup_MinRightMargin, DEF_PageSetup_MinRightMargin)
        .fMinTopMargin = PropBag.ReadProperty(PROP_PageSetup_MinTopMargin, DEF_PageSetup_MinTopMargin)
        .fMinBottomMargin = PropBag.ReadProperty(PROP_PageSetup_MinBottomMargin, DEF_PageSetup_MinBottomMargin)
    End With
    
    msHelpFile = PropBag.ReadProperty(PROP_Help_File, DEF_Help_File)
End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Size to a fixed width and height.
'---------------------------------------------------------------------------------------
    UserControl.SIZE ScaleX(28, vbPixels, vbTwips), ScaleY(28, vbPixels, vbTwips)
End Sub

Private Sub UserControl_Terminate()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Release the shell module handle.
'---------------------------------------------------------------------------------------
    ReleaseShellMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Save property values between instances.
'---------------------------------------------------------------------------------------
    PropBag.WriteProperty PROP_RaiseExtendedErrors, mbRaiseExtendedErrors, DEF_RaiseExtendedErrors
    PropBag.WriteProperty PROP_RaiseExtendedErrors, mbRaiseCancelErrors, DEF_RaiseCancelErrors
    
    With mtFile
        PropBag.WriteProperty PROP_File_Flags, .iFlags, DEF_File_Flags
        PropBag.WriteProperty PROP_File_Filter, .sFilter, DEF_File_Filter
        PropBag.WriteProperty PROP_File_FilterIndex, .iFilterIndex, DEF_File_FilterIndex
        PropBag.WriteProperty PROP_File_DefaultExt, .sDefaultExt, DEF_File_DefaultExt
        PropBag.WriteProperty PROP_File_InitPath, .sInitPath, DEF_File_InitPath
        PropBag.WriteProperty PROP_File_InitFile, .sInitFile, DEF_File_InitFile
        PropBag.WriteProperty PROP_File_Title, .sTitle, DEF_File_Title
    End With

    With mtFolder
        PropBag.WriteProperty PROP_Folder_Title, .sTitle, DEF_Folder_Title
        PropBag.WriteProperty PROP_Folder_InitPath, .sInitialPath, DEF_Folder_InitPath
        PropBag.WriteProperty PROP_Folder_RootPath, .sRootPath, DEF_Folder_RootPath
        PropBag.WriteProperty PROP_Folder_Flags, .iFlags, DEF_Folder_Flags
    End With

    With mtFont
        PropBag.WriteProperty PROP_Font_Flags, .iFlags, DEF_Font_Flags
        PropBag.WriteProperty PROP_Font_MinSize, .iMinSize, DEF_Font_MinSize
        PropBag.WriteProperty PROP_Font_MaxSize, .iMaxSize, DEF_Font_MaxSize
        PropBag.WriteProperty PROP_Font_Color, .iColor, DEF_Font_Color
    End With

    With mtColor
        PropBag.WriteProperty PROP_Color, .iColor, DEF_Color
        PropBag.WriteProperty PROP_Color_Flags, .iFlags, DEF_Color_Flags
    End With

    With mtPrint
        PropBag.WriteProperty PROP_Print_Flags, .iFlags, DEF_Print_Flags
        PropBag.WriteProperty PROP_Print_Range, .iRange, DEF_Print_Range
        PropBag.WriteProperty PROP_Print_MinPage, .iMinPage, DEF_Print_MinPage
        PropBag.WriteProperty PROP_Print_MaxPage, .iMaxPage, DEF_Print_MaxPage
    End With
    
    With mtPageSetup
        PropBag.WriteProperty PROP_PageSetup_Flags, .iFlags, DEF_PageSetup_Flags
        PropBag.WriteProperty PROP_PageSetup_Units, .iUnits, DEF_PageSetup_Units
        PropBag.WriteProperty PROP_PageSetup_LeftMargin, .fLeftMargin, DEF_PageSetup_LeftMargin
        PropBag.WriteProperty PROP_PageSetup_RightMargin, .fRightMargin, DEF_PageSetup_RightMargin
        PropBag.WriteProperty PROP_PageSetup_TopMargin, .fTopMargin, DEF_PageSetup_TopMargin
        PropBag.WriteProperty PROP_PageSetup_BottomMargin, .fBottomMargin, DEF_PageSetup_BottomMargin
        
        PropBag.WriteProperty PROP_PageSetup_MinLeftMargin, .fMinLeftMargin, DEF_PageSetup_MinLeftMargin
        PropBag.WriteProperty PROP_PageSetup_MinRightMargin, .fMinRightMargin, DEF_PageSetup_MinRightMargin
        PropBag.WriteProperty PROP_PageSetup_MinTopMargin, .fMinTopMargin, DEF_PageSetup_MinTopMargin
        PropBag.WriteProperty PROP_PageSetup_MinBottomMargin, .fMinBottomMargin, DEF_PageSetup_MinBottomMargin
    End With
    
    PropBag.WriteProperty PROP_Help_File, msHelpFile, DEF_Help_File
End Sub



Private Sub pErr()
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Raise errors only if property values dictate that we do so.
'---------------------------------------------------------------------------------------
    If miExtendedError Then
        If mbRaiseExtendedErrors Then gErr vbccComDlgExtendedError, ClassName, "Common dialog extended error: " & miExtendedError
    Else
        If mbRaiseCancelErrors Then gErr vbccUserCanceled, ClassName
    End If
End Sub

Private Sub iComDlgHook_Proc(ByVal iDlgType As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Handle hook procedure callbacks for the dialogs.
'---------------------------------------------------------------------------------------
    Dim lbCancel As OLE_CANCELBOOL
    Select Case iDlgType
    Case dlgTypeColor
        Select Case uMsg
        Case WM_INITDIALOG:         RaiseEvent ColorInit(hWnd)
        Case WM_COMMAND:            RaiseEvent ColorWMCommand(hWnd, wParam, lParam, lReturn)
        Case WM_DESTROY:            RaiseEvent ColorClose(hWnd)
        Case Color_OKMsg:           RaiseEvent ColorOK(hWnd, lbCancel)
                                    lReturn = Abs(lbCancel)
        End Select
    Case dlgTypeFont
        Select Case uMsg
        Case WM_INITDIALOG:         RaiseEvent FontInit(hWnd)
        Case WM_COMMAND:            RaiseEvent FontWMCommand(hWnd, wParam, lParam, lReturn)
        Case WM_DESTROY:            RaiseEvent FontClose(hWnd)
        End Select
    Case dlgTypeFolder
        Select Case uMsg
        Case BFFM_INITIALIZED:      RaiseEvent FolderInit(hWnd)
                                    If lParam Then SendMessage hWnd, BFFM_SETSELECTIONA, OneL, lParam
        Case BFFM_SELCHANGED:       RaiseEvent FolderChanged(hWnd, mCommonDialog.Folder_PathFromPidl(wParam))
        Case BFFM_VALIDATEFAILEDA:  RaiseEvent FolderValidationFailed(hWnd, lstrToStringAFunc(wParam), lbCancel)
                                    lReturn = (lbCancel + OneL)
        End Select
    Case dlgTypePageSetup
        Select Case uMsg
        Case WM_INITDIALOG:         RaiseEvent PageSetupInit(hWnd)
        Case WM_COMMAND:            RaiseEvent PageSetupWMCommand(hWnd, wParam, lParam, lReturn)
        Case WM_DESTROY:            RaiseEvent PageSetupClose(hWnd)
        End Select
    Case dlgTypePrint
        Select Case uMsg
        Case WM_INITDIALOG:         RaiseEvent PrintInit(hWnd)
        Case WM_COMMAND:            RaiseEvent PrintWMCommand(hWnd, wParam, lParam, lReturn)
        Case WM_DESTROY:            RaiseEvent PrintClose(hWnd)
        End Select
    Case dlgTypeFile
        Select Case uMsg
        Case WM_INITDIALOG:         RaiseEvent FileInit(hWnd)
        Case WM_COMMAND:            RaiseEvent FileWMCommand(hWnd, wParam, lParam, lReturn)
        Case WM_DESTROY:            RaiseEvent FileClose(hWnd)
        Case WM_NOTIFY
            Select Case MemOffset32(lParam, NMHDR_code)
            Case CDN_SELCHANGE:     RaiseEvent FileChangeFile(hWnd)
            Case CDN_FOLDERCHANGE:  RaiseEvent FileChangeFolder(hWnd)
            Case CDN_HELP:          RaiseEvent FileHelpClicked(hWnd)
            Case CDN_TYPECHANGE:    RaiseEvent FileChangeType(hWnd)
            Case CDN_FILEOK:        RaiseEvent FileOK(hWnd, lbCancel)
                                    lReturn = Abs(lbCancel)
                                    SetWindowLong hWnd, DWL_MSGRESULT, lReturn
            End Select
        End Select
    End Select
End Sub

Private Sub pPropChanged(ByRef sName As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Call the PropertyChange method at design time.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode = False Then PropertyChanged sName
End Sub

Public Property Get ExtendedError() As eComDlgExtendedError
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the extended error returned by the last common dialog.
'---------------------------------------------------------------------------------------
    ExtendedError = miExtendedError
End Property

Public Property Get RaiseCancelErrors() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a value indicating whether an error will be raised when the user
'             cancels a dialog.
'---------------------------------------------------------------------------------------
    RaiseCancelErrors = mbRaiseCancelErrors
End Property
Public Property Let RaiseCancelErrors(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set a value indicating whether an error will be raised when the user
'             cancels a dialog.
'---------------------------------------------------------------------------------------
    mbRaiseCancelErrors = bNew
    pPropChanged PROP_RaiseCancelErrors
End Property

Public Property Get RaiseExtendedErrors() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a value indicating whether an error will be raised when an error
'             occurs that prevents the dialog from showing.
'---------------------------------------------------------------------------------------
    RaiseExtendedErrors = mbRaiseExtendedErrors
End Property
Public Property Let RaiseExtendedErrors(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set a value indicating whether an error will be raised when an error
'             occurs that prevents the dialog from showing.
'---------------------------------------------------------------------------------------
    mbRaiseExtendedErrors = bNew
    pPropChanged PROP_RaiseExtendedErrors
End Property

Public Sub CenterDialog(ByVal hDlg As Long, ByVal vCenterTo As Variant)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Center the dialog over the window or form provided.
'---------------------------------------------------------------------------------------
    Const liBufferLen As Long = 256
    
    Dim lpBuffer As Long
    lpBuffer = MemAllocFromString(ZeroL, liBufferLen)
    
    If lpBuffer Then
        
        Dim lhDlgParent As Long:    lhDlgParent = GetParent(hDlg)
        Dim lsDlgClass As String:   lsDlgClass = StrConv(WC_DIALOG & vbNullChar, vbFromUnicode)
        Dim lpDlgClass As Long:     lpDlgClass = StrPtr(lsDlgClass)
        
        Do While CBool(lhDlgParent)
            If GetClassName(lhDlgParent, ByVal lpBuffer, liBufferLen) = ZeroL Then Exit Do
            If lstrcmp(lpBuffer, lpDlgClass) Then Exit Do
            hDlg = lhDlgParent
            lhDlgParent = GetParent(hDlg)
        Loop
        
        MemFree lpBuffer
        
    End If
    
    Dialog_CenterWindow hDlg, vCenterTo
End Sub

Public Property Get ColorCustom(ByVal iIndex As Long) As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a custom color from the array.
'---------------------------------------------------------------------------------------
    If iIndex > NegOneL And iIndex < 16& Then ColorCustom = mtColor.iColors(iIndex)
End Property

Public Property Let ColorCustom(ByVal iIndex As Long, ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set a custom color from the array.
'---------------------------------------------------------------------------------------
    If iIndex > NegOneL And iIndex < 16& Then mtColor.iColors(iIndex) = iNew
End Property

Public Property Get ColorCustomCount() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the number of custom colors available.
'---------------------------------------------------------------------------------------
    ColorCustomCount = 16&
End Property

Public Property Get Color() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color returned by the color dialog.
'---------------------------------------------------------------------------------------
    Color = mtColor.iColor
End Property
Public Property Let Color(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color initially displayed by the color dialog..
'---------------------------------------------------------------------------------------
    mtColor.iColor = iNew
    pPropChanged PROP_Color
End Property

Public Property Get ColorFlags() As eColorDialog
Attribute ColorFlags.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags passed to the color dialog.
'---------------------------------------------------------------------------------------
    ColorFlags = mtColor.iFlags
End Property
Public Property Let ColorFlags(ByVal iNew As eColorDialog)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the flags passed to the color dialog.
'---------------------------------------------------------------------------------------
    mtColor.iFlags = iNew
    pPropChanged PROP_Color_Flags
End Property

Public Function ShowColor( _
                Optional ByRef iColor As OLE_COLOR = NegOneL, _
                Optional ByVal iFlags As eColorDialog = NegOneL, _
                Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                    As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a color dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
    Dim ltDialog As tColorDialog
    
    With ltDialog
        Set .oHookCallback = Me
        .hWndOwner = RootParent(UserControl.ContainerHwnd)
        CopyMemory .iColors(0), mtColor.iColors(0), 64&
        If iColor <> NegOneL Then .iColor = iColor Else .iColor = mtColor.iColor
        If iFlags <> NegOneL Then .iFlags = iFlags Else .iFlags = mtColor.iFlags
    End With
    
    ShowColor = mCommonDialog.Color_ShowIndirect(ltDialog)
    
    miExtendedError = ltDialog.iReturnExtendedError
    iReturnExtendedError = miExtendedError
    
    CopyMemory mtColor.iColors(0), ltDialog.iColors(0), 64&
    
    If ShowColor Then
        iColor = ltDialog.iColor
        mtColor.iColor = iColor
    Else
        pErr
    End If
    
End Function



Private Sub pFileGetUDT( _
            ByRef tFileDialog As tFileDialog, _
            ByRef sTitle As String, _
            ByVal iFlags As eFileDialog, _
            ByRef sFilter As String, _
            ByVal iFilterIndex As Long, _
            ByRef sDefaultExt As String, _
            ByRef sInitPath As String, _
            ByRef sInitFile As String, _
            ByRef vTemplate As Variant, _
            ByVal hInstance As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the members of a file dialog udt to the given arguments or modular
'             variables.
'---------------------------------------------------------------------------------------
    With tFileDialog
        .hWndOwner = RootParent(UserControl.ContainerHwnd)
        Set .oHookCallback = Me
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtFile.iFlags
        
        If LenB(sTitle) _
            Then .sTitle = CStr(sTitle) _
            Else .sTitle = mtFile.sTitle
        
        If LenB(sInitFile) _
            Then .sInitFile = CStr(sInitFile) _
            Else .sInitFile = mtFile.sInitFile
        
        If LenB(sInitPath) _
            Then .sInitPath = CStr(sInitPath) _
            Else .sInitPath = mtFile.sInitPath
        
        If LenB(sDefaultExt) _
            Then .sDefaultExt = CStr(sDefaultExt) _
            Else .sDefaultExt = mtFile.sDefaultExt
        
        If hInstance <> NegOneL _
            Then .hInstance = hInstance _
            Else .hInstance = mtFile.hInstance
        
        If IsMissing(vTemplate) _
            Then .vTemplate = mtFile.vTemplate _
            Else .vTemplate = vTemplate
        
        If LenB(sFilter) _
            Then .sFilter = CStr(sFilter) _
            Else .sFilter = mtFile.sFilter
        
        If iFilterIndex <> NegOneL _
            Then .iFilterIndex = iFilterIndex _
            Else .iFilterIndex = mtFile.iFilterIndex
        
    End With
End Sub

Public Function ShowFileOpen( _
   Optional ByRef sReturnFileName As String, _
   Optional ByVal iFlags As eFileDialog = NegOneL, _
   Optional ByRef sFilter As String, _
   Optional ByVal iFilterIndex As Long = NegOneL, _
   Optional ByRef sDefaultExt As String, _
   Optional ByRef sInitPath As String, _
   Optional ByRef sInitFile As String, _
   Optional ByRef sTitle As String, _
   Optional ByVal vTemplate As Variant, _
   Optional ByVal hInstance As Long, _
   Optional ByRef iReturnFlags As eFileDialog, _
   Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
   Optional ByRef iReturnFilterIndex As Long) _
                As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a file open dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
                
    Dim ltDialog As tFileDialog
    pFileGetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, vTemplate, hInstance
    
    ShowFileOpen = mCommonDialog.File_ShowOpenIndirect(ltDialog)
    
    iReturnExtendedError = ltDialog.iReturnExtendedError
    miExtendedError = ltDialog.iReturnExtendedError
    
    If ShowFileOpen Then
    
        With ltDialog
            
            iReturnFilterIndex = .iReturnFilterIndex
            iReturnFlags = .iReturnFlags
            sReturnFileName = .sReturnFileName
            
            mtFile.iReturnFilterIndex = .iReturnFilterIndex
            mtFile.iReturnFlags = .iReturnFlags
            mtFile.sReturnFileName = .sReturnFileName
        End With

    Else
        pErr
        
    End If

End Function

Public Function ShowFileSave( _
   Optional ByRef sReturnFileName As String, _
   Optional ByVal iFlags As eFileDialog = NegOneL, _
   Optional ByRef sFilter As String, _
   Optional ByVal iFilterIndex As Long = NegOneL, _
   Optional ByRef sDefaultExt As String, _
   Optional ByRef sInitPath As String, _
   Optional ByRef sInitFile As String, _
   Optional ByRef sTitle As String, _
   Optional ByVal vTemplate As Variant, _
   Optional ByVal hInstance As Long, _
   Optional ByRef iReturnFlags As eFileDialog, _
   Optional ByRef iReturnExtendedError As eComDlgExtendedError, _
   Optional ByRef iReturnFilterIndex As Long) _
                As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a file save dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
    Dim ltDialog As tFileDialog
    pFileGetUDT ltDialog, sTitle, iFlags, sFilter, iFilterIndex, sDefaultExt, sInitPath, sInitFile, vTemplate, hInstance
    
    ShowFileSave = mCommonDialog.File_ShowSaveIndirect(ltDialog)
    
    miExtendedError = ltDialog.iReturnExtendedError
    iReturnExtendedError = miExtendedError
    
    If ShowFileSave Then
        With ltDialog
            iReturnFilterIndex = .iReturnFilterIndex
            iReturnFlags = .iReturnFlags
            sReturnFileName = .sReturnFileName
            
            mtFile.iReturnFilterIndex = .iReturnFilterIndex
            mtFile.iReturnFlags = .iReturnFlags
            mtFile.sReturnFileName = .sReturnFileName
        End With
        
    Else
        pErr
        
    End If
    
End Function

Public Function FileGetFilter(ParamArray vFilters() As Variant) As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return a filter string given separate parameters.
'---------------------------------------------------------------------------------------
    Dim i As Long
    For i = LBound(vFilters) To UBound(vFilters)
        If Not IsMissing(vFilters(i)) _
            Then FileGetFilter = FileGetFilter & vFilters(i) & OFN_FilterDelim
    Next
End Function

Public Property Get FileNames(ByRef sFilePath As String, ByRef sFileNames() As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Parse a return file string for multiple file names.
'---------------------------------------------------------------------------------------
    FileNames = mCommonDialog.File_GetMultiFileNames(mtFile.sReturnFileName, sFilePath, sFileNames)
End Property

Public Property Get FileFlags() As eFileDialog
Attribute FileFlags.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileFlags = mtFile.iFlags
End Property
Public Property Let FileFlags(ByVal iNew As eFileDialog)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the flags passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.iFlags = iNew
    pPropChanged PROP_File_Flags
End Property

Public Property Get FileFilter() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the filter passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileFilter = mtFile.sFilter
End Property
Public Property Let FileFilter(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the filter passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.sFilter = sNew
    pPropChanged PROP_File_Filter
End Property

Public Property Get FileFilterIndex() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the filter index passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileFilterIndex = mtFile.iFilterIndex
End Property
Public Property Let FileFilterIndex(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the filter index passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.iFilterIndex = iNew
    pPropChanged PROP_File_FilterIndex
End Property

Public Property Get FileDefaultExt() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the default extension passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileDefaultExt = mtFile.sDefaultExt
End Property
Public Property Let FileDefaultExt(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the default extension passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.sDefaultExt = sNew
    pPropChanged PROP_File_DefaultExt
End Property

Public Property Get FileInitialPath() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the initial path passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileInitialPath = mtFile.sInitPath
End Property
Public Property Let FileInitialPath(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the initial path passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.sInitPath = sNew
    pPropChanged PROP_File_InitPath
End Property

Public Property Get FileInitialFile() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the initial file passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileInitialFile = mtFile.sInitFile
End Property
Public Property Let FileInitialFile(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the initial file passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.sInitFile = sNew
    pPropChanged PROP_File_InitFile
End Property

Public Property Get FileTitle() As String
Attribute FileTitle.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the window title passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileTitle = mtFile.sTitle
End Property
Public Property Let FileTitle(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the window title passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.sTitle = sNew
    pPropChanged PROP_File_Title
End Property

Public Property Get FileTemplate() As Variant
Attribute FileTemplate.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the template handle or name passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileTemplate = mtFile.vTemplate
End Property
Public Property Let FileTemplate(ByRef vNew As Variant)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the template handle or name passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.vTemplate = vNew
End Property

Public Property Get FileTemplateHInstance() As Long
Attribute FileTemplateHInstance.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the template hInstance passed to the file dialog.
'---------------------------------------------------------------------------------------
    FileTemplateHInstance = mtFile.hInstance
End Property
Public Property Let FileTemplateHInstance(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the template hInstance passed to the file dialog.
'---------------------------------------------------------------------------------------
    mtFile.hInstance = iNew
End Property

Public Property Get FileReturnFlags() As eFileDialog
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags returned by the file dialog.
'---------------------------------------------------------------------------------------
    FileReturnFlags = mtFile.iReturnFlags
End Property

Public Property Get FileReturnFilterIndex() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the filter index returned by the file dialog.
'---------------------------------------------------------------------------------------
    FileReturnFilterIndex = mtFile.iReturnFilterIndex
End Property

Public Property Get FilePath() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the file path returned by the file dialog.
'---------------------------------------------------------------------------------------
    FilePath = mtFile.sReturnFileName
End Property

Public Property Get FolderFlags() As eFolderDialog
Attribute FolderFlags.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags passed to the folder dialog.
'---------------------------------------------------------------------------------------
    FolderFlags = mtFolder.iFlags
End Property
Public Property Let FolderFlags(ByVal iNew As eFolderDialog)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the flags passed to the folder dialog.
'---------------------------------------------------------------------------------------
    mtFolder.iFlags = iNew
    pPropChanged PROP_Folder_Flags
End Property

Public Property Get FolderTitle() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the dialog title passed to the folder dialog.
'---------------------------------------------------------------------------------------
   FolderTitle = mtFolder.sTitle
End Property
Public Property Let FolderTitle(ByVal sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the dialog title passed to the folder dialog.  The folder dialog
'             always has the same window title, the "dialog title" is displayed in a
'             label at the top of the dialog.
'---------------------------------------------------------------------------------------
   mtFolder.sTitle = sNew
   pPropChanged PROP_Folder_Title
End Property

Public Property Get FolderInitialPath() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the initial path passed to the folder dialog.
'---------------------------------------------------------------------------------------
   FolderInitialPath = mtFolder.sInitialPath
End Property
Public Property Let FolderInitialPath(ByVal sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the initial path passed to the folder dialog.  This functionality
'             is implemented by way of the WM_INITDIALOG notification, and accordingly
'             is only effectual when a dialog hook is enabled.
'---------------------------------------------------------------------------------------
   mtFolder.sInitialPath = sNew
   pPropChanged PROP_Folder_InitPath
End Property

Public Property Get FolderRootPath() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the root path passed to the folder dialog.
'---------------------------------------------------------------------------------------
   FolderRootPath = mtFolder.sRootPath
End Property
Public Property Let FolderRootPath(ByVal sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the root path passed to the folder dialog.
'---------------------------------------------------------------------------------------
   mtFolder.sRootPath = sNew
   pPropChanged PROP_Folder_RootPath
End Property

Public Property Get FolderPath() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the path returned by the folder dialog.
'---------------------------------------------------------------------------------------
    FolderPath = mtFolder.sReturnPath
End Property

Public Function FolderGetSpecial(ByVal iFolder As eSpecialFolders) As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the path to a given operating system folder.
'---------------------------------------------------------------------------------------
    FolderGetSpecial = mCommonDialog.Folder_GetSpecial(RootParent(UserControl.ContainerHwnd), iFolder)
End Function

Public Sub FolderSetCurrent(ByVal hDlg As Long, ByRef sFolder As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the current folder in the dialog.  Useful only if a dialog hook is enabled.
'---------------------------------------------------------------------------------------
    Dim lsAnsi As String
    lsAnsi = StrConv(sFolder & vbNullChar, vbFromUnicode)
    SendMessage hDlg, BFFM_SETSELECTIONA, OneL, StrPtr(lsAnsi)
End Sub
Public Sub FolderSetStatus(ByVal hDlg As Long, ByRef sText As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the current status string in the dialog.  Useful only if a dialog hook is enabled.
'---------------------------------------------------------------------------------------
    Dim lsAnsi As String
    lsAnsi = StrConv(sText & vbNullChar, vbFromUnicode)
    SendMessage hDlg, BFFM_SETSTATUSTEXTA, ZeroL, StrPtr(lsAnsi)
End Sub

Public Sub FolderEnableOK(ByVal hDlg As Long, ByVal bEnabled As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the enabled state of the OK button.  Useful only if a dialog hook is enabled.
'---------------------------------------------------------------------------------------
    SendMessage hDlg, BFFM_ENABLEOK, ZeroL, Abs(bEnabled)
End Sub

    
Public Function ShowFolder( _
   Optional ByRef sReturnPath As String, _
   Optional ByVal iFlags As eFolderDialog = NegOneL, _
   Optional ByRef sTitle As String, _
   Optional ByRef sInitialPath As String, _
   Optional ByRef sRootPath As String) _
                As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a folder dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
    
    Dim ltDialog As tFolderDialog
    
    With ltDialog
        .hWndOwner = RootParent(UserControl.ContainerHwnd)
        Set .oHookCallback = Me
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtFolder.iFlags
        
        If LenB(sTitle) _
            Then .sTitle = CStr(sTitle) _
            Else .sTitle = mtFolder.sTitle
        
        If LenB(sInitialPath) _
            Then .sInitialPath = sInitialPath _
            Else .sInitialPath = mtFolder.sInitialPath
        
        If LenB(sRootPath) _
            Then .sRootPath = CStr(sRootPath) _
            Else .sRootPath = mtFolder.sRootPath
        
    End With
    
    ShowFolder = mCommonDialog.Folder_ShowIndirect(ltDialog)
    
    miExtendedError = 0
    
    If ShowFolder Then
        sReturnPath = ltDialog.sReturnPath
        mtFolder.sReturnPath = sReturnPath
        
    Else
        pErr
        
    End If
    
End Function




Public Property Get Font() As Object
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the font returned from the folder dialog.
'---------------------------------------------------------------------------------------
    Set Font = mtFont.oFont
End Property

Public Property Set Font(ByVal oNew As Object)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the font passed to the folder dialog. This can be a vbComCtl.cFont or
'             a stdole.StdFont object.
'---------------------------------------------------------------------------------------
    Set mtFont.oFont = oNew
End Property

Public Property Get FontFlags() As eFontDialog
Attribute FontFlags.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags passed to the font dialog.
'---------------------------------------------------------------------------------------
    FontFlags = mtFont.iFlags
End Property
Public Property Let FontFlags(ByVal iNew As eFontDialog)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the flags passed to the font dialog.
'---------------------------------------------------------------------------------------
    mtFont.iFlags = iNew
    pPropChanged PROP_Font_Flags
End Property

Public Property Get FontHdc() As Long
Attribute FontHdc.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the hdc passed to the font dialog.
'---------------------------------------------------------------------------------------
    FontHdc = mtFont.hDc
End Property
Public Property Let FontHdc(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the hdc passed to the font dialog.
'---------------------------------------------------------------------------------------
    mtFont.hDc = iNew
End Property

Public Property Get FontColor() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the color passed to the font dialog or returned from the font dialog.
'---------------------------------------------------------------------------------------
    FontColor = mtFont.iColor
End Property
Public Property Let FontColor(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the color passed to the font dialog.
'---------------------------------------------------------------------------------------
    mtFont.iColor = iNew
    pPropChanged PROP_Font_Color
End Property

Public Property Get FontMinSize() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minumum size passed to the font dialog.
'---------------------------------------------------------------------------------------
    FontMinSize = mtFont.iMinSize
End Property
Public Property Let FontMinSize(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minumum size passed to the font dialog.
'---------------------------------------------------------------------------------------
    mtFont.iMinSize = iNew
    pPropChanged PROP_Font_MaxSize
End Property

Public Property Get FontMaxSize() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the maximum size passed to the font dialog.
'---------------------------------------------------------------------------------------
    FontMaxSize = mtFont.iMaxSize
End Property
Public Property Let FontMaxSize(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the maximum size passed to the font dialog.
'---------------------------------------------------------------------------------------
    mtFont.iMaxSize = iNew
    pPropChanged PROP_Font_MaxSize
End Property

Public Property Get FontReturnFlags() As eFontDialog
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags returned by the font dialog.
'---------------------------------------------------------------------------------------
    FontReturnFlags = mtFont.iReturnFlags
End Property

Public Function ShowFont( _
                Optional ByVal oFont As Object, _
                Optional ByVal iFlags As eFontDialog = NegOneL, _
                Optional ByVal hDc As Long = NegOneL, _
                Optional ByVal iMinSize As Long = NegOneL, _
                Optional ByVal iMaxSize As Long = NegOneL, _
                Optional ByRef iColor As OLE_COLOR, _
                Optional ByRef iReturnFlags As eFontDialog, _
                Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                    As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a font dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
    
    Dim ltDialog As tFontDialog
    
    With ltDialog
        
        .hWndOwner = RootParent(UserControl.ContainerHwnd)
        Set .oHookCallback = Me
        
        If oFont Is Nothing _
            Then Set .oFont = mtFont.oFont _
            Else Set .oFont = oFont
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtFont.iFlags
        
        If hDc <> NegOneL _
            Then .hDc = hDc _
            Else .hDc = mtFont.hDc
        
        If iMinSize <> NegOneL _
            Then .iMinSize = iMinSize _
            Else .iMinSize = mtFont.iMinSize
        
        If iMaxSize <> NegOneL _
            Then .iMaxSize = iMaxSize _
            Else .iMaxSize = mtFont.iMaxSize
        
        If iColor <> NegOneL _
            Then .iColor = TranslateColor(iColor) _
            Else .iColor = mtFont.iColor
        
    End With
    
    ShowFont = mCommonDialog.Font_ShowIndirect(ltDialog)
    
    iReturnExtendedError = ltDialog.iReturnExtendedError
    miExtendedError = ltDialog.iReturnExtendedError
    
    If ShowFont Then
        mtFont.iColor = ltDialog.iColor
        mtFont.iReturnFlags = ltDialog.iReturnFlags
        
        iColor = ltDialog.iColor
        iReturnFlags = ltDialog.iReturnFlags
    Else
        pErr
    End If
    
End Function




Public Property Get PageSetupFlags() As ePrintPageSetup
Attribute PageSetupFlags.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupFlags = mtPageSetup.iFlags
End Property
Public Property Let PageSetupFlags(ByVal iNew As ePrintPageSetup)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the flags passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.iFlags = iNew
    pPropChanged PROP_PageSetup_Flags
End Property

Public Property Get PageSetupLeftMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the left margin passed to the page setup dialog or returned from the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupLeftMargin = mtPageSetup.fLeftMargin
End Property
Public Property Let PageSetupLeftMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the left margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fLeftMargin = rNew
    pPropChanged PROP_PageSetup_LeftMargin
End Property

Public Property Get PageSetupMinLeftMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minimum left margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupMinLeftMargin = mtPageSetup.fMinLeftMargin
End Property
Public Property Let PageSetupMinLeftMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minimum left margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fMinLeftMargin = rNew
    pPropChanged PROP_PageSetup_MinLeftMargin
End Property

Public Property Get PageSetupRightMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the right margin passed to the page setup dialog or returned from the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupRightMargin = mtPageSetup.fRightMargin
End Property
Public Property Let PageSetupRightMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the right margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fRightMargin = rNew
    pPropChanged PROP_PageSetup_RightMargin
End Property

Public Property Get PageSetupMinRightMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minimum right margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupMinRightMargin = mtPageSetup.fMinRightMargin
End Property
Public Property Let PageSetupMinRightMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minimum right margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fMinRightMargin = rNew
    pPropChanged PROP_PageSetup_MinRightMargin
End Property

Public Property Get PageSetupTopMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the top margin passed to the page setup dialog or returned from the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupTopMargin = mtPageSetup.fTopMargin
End Property
Public Property Let PageSetupTopMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the top margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fTopMargin = rNew
    pPropChanged PROP_PageSetup_TopMargin
End Property

Public Property Get PageSetupMinTopMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minimum top margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupMinTopMargin = mtPageSetup.fMinTopMargin
End Property
Public Property Let PageSetupMinTopMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minimum top margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fMinTopMargin = rNew
    pPropChanged PROP_PageSetup_MinTopMargin
End Property

Public Property Get PageSetupBottomMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the bottom margin passed to the page setup dialog or returned from the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupBottomMargin = mtPageSetup.fBottomMargin
End Property
Public Property Let PageSetupBottomMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the bottom margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fBottomMargin = rNew
    pPropChanged PROP_PageSetup_BottomMargin
End Property

Public Property Get PageSetupMinBottomMargin() As Single
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minimum bottom margin passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    PageSetupMinBottomMargin = mtPageSetup.fMinBottomMargin
End Property
Public Property Let PageSetupMinBottomMargin(ByRef rNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minimum bottom passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    mtPageSetup.fMinBottomMargin = rNew
    pPropChanged PROP_PageSetup_MinBottomMargin
End Property

Public Property Get PageSetupUnits() As ePrintPageSetupUnits
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the units in which the page setup margins are specified.
'---------------------------------------------------------------------------------------
    PageSetupUnits = mtPageSetup.iUnits
End Property
Public Property Let PageSetupUnits(ByVal iNew As ePrintPageSetupUnits)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the units in which the page setup margins are specified.
'---------------------------------------------------------------------------------------
    mtPageSetup.iUnits = iNew
    pPropChanged PROP_PageSetup_Units
End Property

Public Property Get PageSetupDeviceMode() As cDeviceMode
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the object representing the DEVMODE structure passed to and returned
'             from the page setup dialog.
'---------------------------------------------------------------------------------------
    If mtPageSetup.oDeviceMode Is Nothing Then Set mtPageSetup.oDeviceMode = New cDeviceMode
    Set PageSetupDeviceMode = mtPageSetup.oDeviceMode
End Property

Public Property Set PageSetupDeviceMode(ByVal oNew As cDeviceMode)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the object representing the DEVMODE structure passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set mtPageSetup.oDeviceMode = New cDeviceMode _
        Else Set mtPageSetup.oDeviceMode = oNew
End Property

Public Property Get PageSetupDeviceNames() As cDeviceNames
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the object representing the DEVNAMES structure passed to and returned
'             from the page setup dialog.
'---------------------------------------------------------------------------------------
    If mtPageSetup.oDeviceNames Is Nothing Then Set mtPageSetup.oDeviceNames = New cDeviceNames
    Set PageSetupDeviceNames = mtPageSetup.oDeviceNames
End Property

Public Property Set PageSetupDeviceNames(ByVal oNew As cDeviceNames)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the object representing the DEVNAMES structure passed to the page setup dialog.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set mtPageSetup.oDeviceNames = New cDeviceNames _
        Else Set mtPageSetup.oDeviceNames = oNew
End Property

Public Function ShowPageSetup( _
            Optional ByVal iUnits As ePrintPageSetupUnits = NegOneL, _
            Optional ByRef fLeftMargin As Single = NegOneF, _
            Optional ByRef fRightMargin As Single = NegOneF, _
            Optional ByRef fTopMargin As Single = NegOneF, _
            Optional ByRef fBottomMargin As Single = NegOneF, _
            Optional ByVal iFlags As ePrintPageSetup = NegOneL, _
            Optional ByVal fMinLeftMargin As Single = NegOneF, _
            Optional ByVal fMinRightMargin As Single = NegOneF, _
            Optional ByVal fMinTopMargin As Single = NegOneF, _
            Optional ByVal fMinBottomMargin As Single = NegOneF, _
            Optional ByRef oDeviceMode As cDeviceMode, _
            Optional ByRef oDeviceNames As cDeviceNames, _
            Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a page setup dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
    
    Dim ltDialog As tPageSetupDialog
    With ltDialog
        .hWndOwner = RootParent(UserControl.ContainerHwnd)
        Set .oHookCallback = Me

        If fLeftMargin <> NegOneF _
            Then .fLeftMargin = fLeftMargin _
            Else .fLeftMargin = mtPageSetup.fLeftMargin
        
        If fMinLeftMargin <> NegOneF _
            Then .fMinLeftMargin = fMinLeftMargin _
            Else .fMinLeftMargin = mtPageSetup.fMinLeftMargin
        
        If fRightMargin <> NegOneF _
            Then .fRightMargin = fRightMargin _
            Else .fRightMargin = mtPageSetup.fRightMargin
        
        If fMinRightMargin <> NegOneF _
            Then .fMinRightMargin = fMinRightMargin _
            Else .fMinRightMargin = mtPageSetup.fMinRightMargin
        
        If fTopMargin <> NegOneF _
            Then .fTopMargin = fTopMargin _
            Else .fTopMargin = mtPageSetup.fTopMargin
        
        If fMinTopMargin <> NegOneF _
            Then .fMinTopMargin = fMinTopMargin _
            Else .fMinTopMargin = mtPageSetup.fMinTopMargin
        
        If fBottomMargin <> NegOneF _
            Then .fBottomMargin = fBottomMargin _
            Else .fBottomMargin = mtPageSetup.fBottomMargin
        
        If fMinBottomMargin <> NegOneF _
            Then .fMinBottomMargin = fMinBottomMargin _
            Else .fMinBottomMargin = mtPageSetup.fMinBottomMargin
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtPageSetup.iFlags
                
        If iUnits <> iUnits _
            Then .iUnits = iUnits _
            Else .iUnits = mtPageSetup.iUnits
        
        If oDeviceMode Is Nothing _
            Then Set .oDeviceMode = mtPageSetup.oDeviceMode _
            Else Set .oDeviceMode = oDeviceMode
        
        If oDeviceNames Is Nothing _
            Then Set .oDeviceNames = mtPageSetup.oDeviceNames _
            Else Set .oDeviceNames = oDeviceNames
    
    End With
    
    ShowPageSetup = mCommonDialog.PageSetup_ShowIndirect(ltDialog)
    
    iReturnExtendedError = miExtendedError
    miExtendedError = ltDialog.iReturnExtendedError
    
    If ShowPageSetup Then
        With ltDialog
            fLeftMargin = .fLeftMargin
            fRightMargin = .fRightMargin
            fTopMargin = .fTopMargin
            fBottomMargin = .fBottomMargin
        End With
        With mtPageSetup
            If .oDeviceMode Is Nothing And oDeviceMode Is Nothing _
                Then Set .oDeviceMode = ltDialog.oDeviceMode
            .fLeftMargin = fLeftMargin
            .fRightMargin = fRightMargin
            .fTopMargin = fTopMargin
            .fBottomMargin = fBottomMargin
        End With
        Set oDeviceMode = ltDialog.oDeviceMode
        Set oDeviceNames = ltDialog.oDeviceNames
    Else
        pErr
    End If
    
End Function







Public Property Get PrintFlags() As ePrintDialog
Attribute PrintFlags.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags passed to the print dialog.
'---------------------------------------------------------------------------------------
    PrintFlags = mtPrint.iFlags
End Property
Public Property Let PrintFlags(ByVal iNew As ePrintDialog)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the flags passed to the print dialog.
'---------------------------------------------------------------------------------------
    mtPrint.iFlags = iNew
    pPropChanged PROP_Print_Flags
End Property

Public Property Get PrintReturnFlags() As ePrintDialog
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the flags returned from the print dialog.
'---------------------------------------------------------------------------------------
    PrintReturnFlags = mtPrint.iReturnFlags
End Property

Public Property Get PrintRange() As ePrintRange
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the range passed to the print dialog.
'---------------------------------------------------------------------------------------
    PrintRange = mtPrint.iRange
End Property
Public Property Let PrintRange(ByVal iNew As ePrintRange)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the range passed to the print dialog.
'---------------------------------------------------------------------------------------
    mtPrint.iRange = iNew
    pPropChanged PROP_Print_Range
End Property

Public Property Get PrintFromPage() As Long
Attribute PrintFromPage.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the from page passed to the print dialog.
'---------------------------------------------------------------------------------------
    PrintFromPage = mtPrint.iFromPage
End Property
Public Property Let PrintFromPage(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the from page passed to the print dialog.
'---------------------------------------------------------------------------------------
    mtPrint.iFromPage = iNew
End Property

Public Property Get PrintToPage() As Long
Attribute PrintToPage.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the to page passed to the print dialog.
'---------------------------------------------------------------------------------------
    PrintToPage = mtPrint.iToPage
End Property
Public Property Let PrintToPage(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the to page passed to the print dialog.
'---------------------------------------------------------------------------------------
    mtPrint.iToPage = iNew
End Property

Public Property Get PrintMinPage() As Long
Attribute PrintMinPage.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the minimum page passed to the print dialog.
'---------------------------------------------------------------------------------------
    PrintMinPage = mtPrint.iMinPage
End Property
Public Property Let PrintMinPage(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the minimum page passed to the print dialog.
'---------------------------------------------------------------------------------------
    mtPrint.iMinPage = iNew
    pPropChanged PROP_Print_MinPage
End Property

Public Property Get PrintMaxPage() As Long
Attribute PrintMaxPage.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the maximum  page passed to the print dialog.
'---------------------------------------------------------------------------------------
    PrintMaxPage = mtPrint.iMaxPage
End Property
Public Property Let PrintMaxPage(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the maximum page passed to the print dialog.
'---------------------------------------------------------------------------------------
    mtPrint.iMaxPage = iNew
    pPropChanged PROP_Print_MaxPage
End Property

Public Property Get PrintDeviceNames() As cDeviceNames
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the object representing the DEVNAMES structure passed to and return
'             from the print dialog.
'---------------------------------------------------------------------------------------
    If mtPrint.oDeviceNames Is Nothing Then Set mtPrint.oDeviceNames = New cDeviceNames
    PrintDeviceNames = mtPrint.oDeviceNames
End Property
Public Property Set PrintDeviceNames(ByVal oNew As cDeviceNames)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the object representing the DEVNAMES structure passed to the print dialog.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set mtPrint.oDeviceNames = New cDeviceNames _
        Else Set mtPrint.oDeviceNames = oNew
End Property

Public Property Get PrintDeviceMode() As cDeviceMode
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the object representing the DEVNAMES structure passed to and return
'             from the print dialog.
'---------------------------------------------------------------------------------------
    If mtPrint.oDeviceMode Is Nothing Then Set mtPrint.oDeviceMode = New cDeviceMode
    Set PrintDeviceMode = mtPrint.oDeviceMode
End Property

Public Property Set PrintDeviceMode(ByVal oNew As cDeviceMode)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the object representing the DEVMODE structure passed to the print dialog.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set mtPrint.oDeviceMode = New cDeviceMode _
        Else Set mtPrint.oDeviceMode = oNew
End Property

Public Property Get PrintHdc() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the hdc returned from the print dialog.
'---------------------------------------------------------------------------------------
    PrintHdc = mtPrint.hDc
End Property

Public Function ShowPrint( _
                Optional ByRef hDc As Long, _
                Optional ByVal iFlags As ePrintDialog = NegOneL, _
                Optional ByRef iRange As ePrintRange = NegOneL, _
                Optional ByRef iFromPage As Long = NegOneL, _
                Optional ByRef iToPage As Long = NegOneL, _
                Optional ByVal iMinPage As Long = NegOneL, _
                Optional ByVal iMaxPage As Long = NegOneL, _
                Optional ByRef oDeviceMode As cDeviceMode, _
                Optional ByVal oDeviceNames As cDeviceNames, _
                Optional ByRef iReturnFlags As ePrintDialog, _
                Optional ByRef iReturnExtendedError As eComDlgExtendedError) _
                    As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a print dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'---------------------------------------------------------------------------------------
    Dim ltDialog As tPrintDialog
    
    With ltDialog
        Set .oHookCallback = Me
        .hWndOwner = RootParent(UserControl.ContainerHwnd)
        
        If iFlags <> NegOneL _
            Then .iFlags = iFlags _
            Else .iFlags = mtPrint.iFlags
        
        If iRange <> NegOneL _
            Then .iRange = iRange _
            Else .iRange = mtPrint.iRange
            
        If iFromPage <> NegOneL _
            Then .iFromPage = iFromPage _
            Else .iFromPage = mtPrint.iFromPage
        
        If iToPage <> NegOneL _
            Then .iToPage = iToPage _
            Else .iToPage = mtPrint.iToPage
            
        If iMinPage <> NegOneL _
            Then .iMinPage = iMinPage _
            Else .iMinPage = mtPrint.iMinPage
            
        If iMaxPage <> NegOneL _
            Then .iMaxPage = iMaxPage _
            Else .iMaxPage = mtPrint.iMaxPage
        
        If oDeviceMode Is Nothing _
            Then Set .oDeviceMode = mtPrint.oDeviceMode _
            Else Set .oDeviceMode = oDeviceMode
            
        If oDeviceNames Is Nothing _
            Then Set .oDeviceNames = mtPrint.oDeviceNames _
            Else Set .oDeviceNames = oDeviceNames
        
    End With
    
    ShowPrint = mCommonDialog.Print_ShowIndirect(ltDialog)
    
    miExtendedError = ltDialog.iReturnExtendedError
    iReturnExtendedError = miExtendedError
    
    If ShowPrint Then
        With ltDialog
            hDc = .hDc
            iRange = .iRange
            iFromPage = .iFromPage
            iToPage = .iToPage
            iReturnFlags = .iReturnFlags
        End With
        With mtPrint
            If Not .oDeviceMode Is ltDialog.oDeviceMode And oDeviceMode Is Nothing Then Set .oDeviceMode = ltDialog.oDeviceMode
            If Not .oDeviceNames Is ltDialog.oDeviceNames And oDeviceNames Is Nothing Then Set .oDeviceNames = ltDialog.oDeviceNames
            If oDeviceMode Is Nothing Then Set oDeviceMode = ltDialog.oDeviceMode
            If oDeviceNames Is Nothing Then Set oDeviceNames = ltDialog.oDeviceNames
            .hDc = hDc
            .iRange = iRange
            .iFromPage = iFromPage
            .iToPage = iToPage
            .iReturnFlags = iReturnFlags
        End With
        Set oDeviceMode = ltDialog.oDeviceMode
        Set oDeviceNames = ltDialog.oDeviceNames
    Else
        pErr
        
    End If
    
End Function


Public Property Get HelpFile() As String
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Return the filename passed to the help dialog.
'---------------------------------------------------------------------------------------
    HelpFile = msHelpFile
End Property
Public Property Let HelpFile(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Set the filename passed to the help dialog.
'---------------------------------------------------------------------------------------
    msHelpFile = sNew
    pPropChanged PROP_Help_File
End Property

Public Function ShowHelp(Optional ByVal iCmdShow As eHelpDialog = dlgHelpContents, Optional ByVal vTopicNameOrId As Variant, Optional ByRef sHelpFile As String) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/10/05
' Purpose   : Show a help dialog.  For arguments not provided, the corresponding
'             modular variables are used.
'
'             Showing the help dialog from the IDE causes a bit of a predicament.  You must
'             decide whether or not to call dlgHelpCloseAll when your program is closing.
'
'                   If you DO call dlgHelpCloseAll:
'                       If the help window was previously shown and had been closed manually
'                       then you may crash after returning to design mode.  To prevent this,
'                       don't close the help file using the window menu or buttons if you display
'                       it unless you are running the compiled version.
'
'                       This is a very quirky error.  It has produced both 'unknown software
'                       exception' and 'access violation'.  If after returning to design mode
'                       as described above you keep VB as the active task, the crash
'                       does not always immediately appear.  It is often only upon starting up
'                       another task, expecially IE or Explorer, that the IDE crashes.
'
'                   If you DO NOT call dlgHelpCloseAll:
'                       If the help window is still visible when returning to design mode,
'                       the IDE will almost certainly crash.
'
'---------------------------------------------------------------------------------------

    If iCmdShow <> dlgHelpCloseAll Or mbHelpWasShown Then
        ShowHelp = mCommonDialog.Help_Show(IIf(LenB(sHelpFile), sHelpFile, msHelpFile), iCmdShow, vTopicNameOrId)
        miExtendedError = 0
        If ShowHelp Then mbHelpWasShown = (iCmdShow <> dlgHelpCloseAll)
    End If
End Function
