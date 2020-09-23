VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gComCtl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'gComCtl.cls                    12/15/04
'
'           PURPOSE:
'               Properties and methods added to the global namespace.
'
'               Enums that are associated with more than one usercontrol.
'
'==================================================================================================

Option Explicit

Public Enum evbComCtlError
    vbccInvalidProcedureCall = 5
    vbccOutOfMemory = 7
    vbccTypeMismatch = 13
    vbccLetSetNoRunTime = 382
    vbccLetSetNoDesignTime = 383
    vbccGetNoRunTime = 393
    vbccGetNoDesignTime = 394
    vbccItemDetached = 35605
    vbccCollectionChangedDuringEnum = 35606
    vbccKeyOrIndexNotFound = 35601
    vbccKeyAlreadyExists = 35602
    vbccUnsupported = 35603
    vbccUserCanceled = 32755
    vbccComDlgExtendedError = 39512
End Enum

Public Enum evbComCtlBorderStyle
    vbccBorderNone
    vbccBorderSingle
    vbccBorderThin
    vbccBorderSunken
End Enum

Public Enum evbComCtlKeyboardState
    vbccShiftMask = vbShiftMask
    vbccControlMask = vbCtrlMask
    vbccAltMask = vbAltMask
End Enum

Public Enum evbComCtlMouseButton
    vbccMouseLButton = MK_LBUTTON
    vbccMouseRButton = MK_RBUTTON
    vbccMouseMButton = MK_MBUTTON
    vbccMouseXButton1 = MK_XBUTTON1
    vbccMouseXButton2 = MK_XBUTTON2
End Enum

Public Enum evbComCtlOleDropEffect
    vbccOleDropCopy = vbDropEffectCopy
    vbccOleDropMove = vbDropEffectMove
    vbccOleDropNone = vbDropEffectNone
End Enum

Public Enum evbComCtlOleDragOverState
    vbccOleDragOver = vbOver
    vbccOleDragLeave = vbLeave
    vbccOleDragEnter = vbEnter
End Enum

Public Enum evbComCtlAlignment
    vbccAlignNone = vbAlignNone
    vbccAlignTop = vbAlignTop
    vbccAlignBottom = vbAlignBottom
    vbccAlignLeft = vbAlignLeft
    vbccAlignRight = vbAlignRight
End Enum

Public Enum evbComCtlControlInfoFlags
    vbccEatsReturn = CTRLINFO_EATS_RETURN
    vbccEatsEscape = CTRLINFO_EATS_ESCAPE
End Enum

#If bDebug Then
    
    Public Property Get DEBUG_Grid(ByVal x As Long, ByVal y As Long) As String
        DEBUG_Grid = mDebug.DEBUG_Grid(x, y)
    End Property
    
    Public Property Get DEBUG_GridCountY() As Long
        DEBUG_GridCountY = mDebug.DEBUG_GridCountY
    End Property
    
    Public Property Get DEBUG_GridCountX() As Long
        DEBUG_GridCountX = mDebug.DEBUG_GridCountX
    End Property
        
#End If

Public Function NewImageList(ByVal iIconWidth As Long, ByVal iIconHeight As Long, Optional ByVal iColorDepth As eImlColorDepth = imlColor) As cImageList
    Set NewImageList = New cImageList
    If NewImageList.fCreate(iIconWidth, iIconHeight, iColorDepth) = False Then gErr vbccOutOfMemory, "gComCtl.NewImageList", "Failed to create the imagelist."
End Function

Public Function NewSysImageList(Optional ByVal bLargeIcons As Boolean) As cSysImageList
    Set NewSysImageList = New cSysImageList
    If NewSysImageList.fCreate(bLargeIcons) = False Then gErr vbccOutOfMemory, "gComCtl.NewImageList", "Failed to create the imagelist."
End Function

Public Property Get SystemColorDepth() As eImlColorDepth
    SystemColorDepth = mComCtlShared.SystemColorDepth
End Property

Public Sub GetComCtlVersion(ByRef iMajor As Long, Optional ByRef iMinor As Long, Optional ByRef iBuild As Long)
    mComCtlVersion.GetCCVersion iMajor, iMinor, iBuild
End Sub

Public Property Get CheckComCtlVersion(ByVal iMajor As Long, Optional ByVal iMinor As Long, Optional ByVal iBuild As Long) As Boolean
    CheckComCtlVersion = mComCtlVersion.CheckCCVersion(iMajor, iMinor, iBuild)
End Property

Public Sub InitComCtl()
    mComCtlVersion.InitCC
End Sub

Public Function GetGComCtl() As gComCtl
    Set GetGComCtl = Me
End Function

Public Property Get dlgFileFilterDelim() As String
    dlgFileFilterDelim = OFN_FilterDelim
End Property

Public Property Get IsAppThemed() As Boolean
    IsAppThemed = mComCtlShared.IsAppThemed
End Property

Public Sub CallMeBack(ByVal oCallee As Object, ByRef sProcName As String, ByVal iProcType As VbCallType, Optional ByVal vArg As Variant)
    Dim loCallback As pcCallbackObject
    Set loCallback = New pcCallbackObject
    loCallback.fInit oCallee, sProcName, iProcType, vArg
End Sub

Public Sub EnterFocus(ByVal oControls As Object)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : A usercontrol will set the focus to the first tabindex even if the user
'             is tabbing backwards by holding the shift button.  Instead, we determine
'             if the user is tabbing backwards and if so, set the focus to the last tabindex.
'---------------------------------------------------------------------------------------
    If KeyIsDown(VK_TAB, False) And KeyIsDown(VK_SHIFT, False) Then
        Dim loControl As Object
        Dim loControlMax As Object
        Dim liTabIndexMax As Long
        
        On Error GoTo iterate
        
        For Each loControl In oControls
            If loControl.Visible Then
                If loControl.TabIndex > liTabIndexMax Then
                    If loControl.TabStop Then
                        If GetDispId(loControl, "SetFocus") Then
                            liTabIndexMax = loControl.TabIndex
                            Set loControlMax = loControl
                        End If
                    End If
                End If
            End If

            If False Then
iterate:
                Resume hereandnow
hereandnow:
            End If
        Next
        
        If Not loControlMax Is Nothing Then
            On Error Resume Next
            loControlMax.SetFocus
        End If
        
        On Error GoTo 0
    End If
End Sub

Public Sub ThemeControls(ByVal oControls As Object, ByVal bEnable As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Enable or disable xp themes for a collection of controls.
'---------------------------------------------------------------------------------------
    Dim ls As String: If bEnable Then ls = vbNullString Else ls = vbNullChar
    Dim lp As Long: lp = StrPtr(ls)
    
    If CheckComCtlVersion(6&) Then
        On Error Resume Next
        Dim o As Object
        For Each o In oControls
            If GetDispId(o, "Themeable") = NegOneL Then
                SetWindowTheme o.hWnd, lp, lp
            Else
                o.Themeable = bEnable
            End If
        Next
        On Error GoTo 0
    End If
End Sub

Public Sub CascadeBackColor(ByVal oControls As Object, ByVal iColor As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Cascade updates of the BackColor property, ignoring controls which should
'             not have the default backcolor, like a listbox or textbox.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim o As Object
    For Each o In oControls
        If Not (TypeOf o Is VB.TextBox Or _
                TypeOf o Is VB.ComboBox Or _
                TypeOf o Is VB.ListBox Or _
                TypeOf o Is VB.Timer Or _
                TypeOf o Is ucComDlg Or _
                TypeOf o Is ucDateTimePicker Or _
                TypeOf o Is ucHotKey Or _
                TypeOf o Is ucListView Or _
                TypeOf o Is ucMaskedEdit Or _
                TypeOf o Is ucMonthCalendar Or _
                TypeOf o Is ucProgressBar Or _
                TypeOf o Is ucPopupMenus Or _
                TypeOf o Is ucRichEdit Or _
                TypeOf o Is ucTabStrip Or _
                TypeOf o Is ucToolbar Or _
                TypeOf o Is ucTreeView Or _
                TypeOf o Is ucUpDown) _
                    Then
            Select Case True
            Case TypeOf o Is ucAnimation, _
                 TypeOf o Is ucComboBoxEx, _
                 TypeOf o Is ucFrame, _
                 TypeOf o Is ucRebar, _
                 TypeOf o Is ucScrollBox, _
                 TypeOf o Is ucStatusBar, _
                 TypeOf o Is ucTrackbar
                    o.ColorBack = iColor
            Case TypeOf o Is VB.CommandButton, _
                 TypeOf o Is VB.CheckBox, _
                 TypeOf o Is VB.Frame, _
                 TypeOf o Is VB.Label, _
                 TypeOf o Is VB.OptionButton, _
                 TypeOf o Is VB.PictureBox, _
                 TypeOf o Is VB.Shape
                    o.BackColor = iColor
            End Select
        End If
    Next
    On Error GoTo 0
End Sub


Public Sub ShowAllUIStates(ByVal hWnd As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Force a VB form or usercontrol to show keyboard mnemonics and focus rectangles.
'---------------------------------------------------------------------------------------
    ForceWindowToShowAllUIStates hWnd
End Sub

Public Sub TbarContainerKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ParamArray oToolbars() As Variant)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Inform constituent toolbars of the container's keypress event.  This is necessary
'             for keyboard mnemonics to work since the toolbar cannot receive focus.  (One might
'             think that the mnemonics would work anyway, since IOLEControl::GetControlInfo is still
'             called, but apparently vb usercontrols cannot receive mnemonics unless the CanGetFocus
'             property is True.
'
'             This also allows the toolbars to detect Alt and F10 keypresses to enter menu tracking mode.
'
'             Calling this sub from the form's keypress event requires the form's KeyPreview
'             property to be True if there are any controls on the form which can receive focus.
'
'             The paramarray argument can consist of any number of toolbars or collections of
'             toolbars.  You can pass the form's Controls collection, but for effeciency you may
'             pass the toolbars explicitly.
'---------------------------------------------------------------------------------------

    Dim loToolbar As ucToolbar
    Dim loObject As Variant
    Dim loCollObject As Variant
    
    On Error GoTo iterate
    
    For Each loObject In oToolbars
        If VarType(loObject) = vbObject Then
            If TypeOf loObject Is ucToolbar Then
                Set loToolbar = loObject
                loToolbar.ContainerKeyDown KeyCode, Shift
            Else
                For Each loCollObject In loObject
                    If VarType(loCollObject) = vbObject Then
                        If TypeOf loCollObject Is ucToolbar Then
                            Set loToolbar = loCollObject
                            loToolbar.ContainerKeyDown KeyCode, Shift
                        End If
                    End If
                    If KeyCode = 0 Then Exit For
                Next
            End If
        End If
        If KeyCode = 0 Then Exit For
        
        If False Then
iterate:
            Resume hereandnow
hereandnow:
        End If
        
    Next
    
    On Error GoTo 0
    
End Sub

Public Sub ShowMenuAtButton(ByVal oMenu As cPopupMenu, ByVal oButton As cButton)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Show a popup menu next to and excluding a toolbar button.
'---------------------------------------------------------------------------------------
    Dim fLeft As Single, fTop As Single, fExcludeLeft As Single, fExcludeTop As Single, fExcludeWidth As Single, fExcludeHeight As Single
    Dim bPreserveVertAlign As Boolean
    oButton.GetIdealPopup fLeft, fTop, fExcludeLeft, fExcludeTop, fExcludeWidth, fExcludeHeight, bPreserveVertAlign
    oMenu.Show mnuPreserveVertAlign * -bPreserveVertAlign, fLeft, fTop, fExcludeLeft, fExcludeTop, fExcludeWidth, fExcludeHeight
End Sub
