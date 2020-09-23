VERSION 5.00
Begin VB.UserControl ucRebar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   PropertyPages   =   "ucRebar.ctx":0000
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucRebar.ctx":000D
End
Attribute VB_Name = "ucRebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucRebar.ctl        12/15/04
'
'           PURPOSE:
'               Implement the comctl32.dll rebar control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Toolbar/vbAccelerator_ToolBar_and_CoolMenu_Control/VB6_Toolbar_Complete_Source.asp
'               cRebar.ctl
'
'==================================================================================================

Option Explicit

Public Enum eRebarHitTest
    rbarHitNowhere = RBHT_NOWHERE
    rbarHitCaption = RBHT_CAPTION
    rbarHitClient = RBHT_CLIENT
    rbarHitGrabber = RBHT_GRABBER
    rbarHitChevron = RBHT_CHEVRON
End Enum


Event BeginDrag(ByVal oBand As cBand, ByRef bCancel As OLE_CANCELBOOL)
Event EndDrag(ByVal oBand As cBand)
Event ChevronPushed(ByVal oBand As cBand, ByVal fLeft As Single, ByVal fTop As Single, ByVal fWidth As Single, ByVal fHeight As Single)
Event Resize()

Private Type tBand
    iId             As Long
    oChild          As Object
    hwndOldParent   As Long
    sKey            As String
    iWidth          As Long
    iHeight         As Long
End Type

Implements iSubclass

Private Const PROP_Font = "Font"
Private Const PROP_Themeable = "Themeable"

Private Const DEF_Themeable = True

Private WithEvents moFont As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private mhFont As Long

Private mhWnd As Long

Private mbThemeable As Boolean

Private mtBands() As tBand
Private miBandCount As Long
Private miBandUbound As Long
Private miBandControl As Long

Private mbClearing As Boolean
Private mbRedraw As Boolean
Private miBandChevronPressed As Long

Private mtBand As REBARBANDINFO
Private msTextBuffer As String * 130

Const ucRebar = "ucRebar"
Const cBands = "cBands"
Const cBand = "cBand"

Const NMREBARCHEVRON_uBand As Long = 12
Const NMREBARCHEVRON_lpRect As Long = 24

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Respond to notifications from the rebar.
'---------------------------------------------------------------------------------------
    Const NMHDR_code As Long = 8
    Const NMREBAR_uBand As Long = 16
    
    Dim bCancel As OLE_CANCELBOOL
    Select Case uMsg
    Case WM_NOTIFY
        Select Case MemOffset32(lParam, NMHDR_code)
        Case RBN_HEIGHTCHANGE
            If Not mbClearing Then pResize
        Case RBN_BEGINDRAG
            RaiseEvent BeginDrag(pBand(MemOffset32(lParam, NMREBAR_uBand)), bCancel)
        Case RBN_ENDDRAG
            RaiseEvent EndDrag(pBand(MemOffset32(lParam, NMREBAR_uBand)))
        Case RBN_CHILDSIZE
            pChildResize lParam
        Case RBN_CHEVRONPUSHED
            miBandChevronPressed = MemOffset32(lParam, NMREBARCHEVRON_uBand)
            pChevronPushed lParam
            miBandChevronPressed = NegOneL
        End Select
    Case UM_SIZEBAND
        pSizeBand wParam, loword(lParam), hiword(lParam)
        bHandled = True
    End Select

End Sub


Private Sub pSizeBand(ByVal hWndChild As Long, ByVal iWidth As Integer, ByVal iHeight As Integer)
    Dim i As Long
    Dim liIndex As Long
    
    If mhWnd Then
        mtBand.fMask = RBBIM_CHILD Or RBBIM_CHILDSIZE
        
        For i = ZeroL To miBandCount
            liIndex = SendMessage(mhWnd, RB_IDTOINDEX, mtBands(i).iId, ZeroL)
            If liIndex > NegOneL Then
                If SendMessage(mhWnd, RB_GETBANDINFO, liIndex, VarPtr(mtBand)) Then
                    If mtBand.hWndChild = hWndChild Then
                        With mtBands(i)
                            If iWidth Then .iWidth = iWidth
                            If iHeight Then .iHeight = iHeight
                            pEvaluateChild mtBands(i).oChild, .iId, .iWidth, .iHeight, pBand_Style(liIndex, RBBS_FIXEDSIZE)
                            mtBand.fMask = RBBIM_CHILDSIZE
                            SendMessage mhWnd, RB_SETBANDINFO, liIndex, VarPtr(mtBand)
                            Exit For
                        End With
                    End If
                End If
            End If
        Next
        
    End If
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub moFont_Changed()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Update the font displayed by the control.
'---------------------------------------------------------------------------------------
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    PropertyChanged PROP_Font
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Update the font source and remeasure the bands if necessary.
'---------------------------------------------------------------------------------------
    If StrComp("Font", PropertyName) = ZeroL Then
        moFont.OnAmbientFontChanged Ambient.Font
        
        If moFont.Source = fntSourceAmbient Then
            
            Dim liIndex As Long
            Dim i As Long
            
            For i = ZeroL To miBandCount - OneL
                With mtBands(i)
                    If TypeOf .oChild Is ucToolbar Then
                        liIndex = SendMessage(mhWnd, RB_IDTOINDEX, mtBands(i).iId, ZeroL)
                        If liIndex > NegOneL Then
                            pEvaluateChild .oChild, .iId, .iWidth, .iHeight, CBool(pBand_Info(liIndex, RBBIM_STYLE) And RBBS_FIXEDSIZE)
                            mtBand.fMask = RBBIM_CHILDSIZE Or RBBIM_IDEALSIZE
                            SendMessage mhWnd, RB_SETBANDINFO, liIndex, VarPtr(mtBand)
                        End If
                    End If
                End With
            Next
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    ForceWindowToShowAllUIStates hWnd
    LoadShellMod
    InitCC ICC_COOL_CLASSES
    mtBand.cbSize = LenB(mtBand)
    Set moFontPage = New pcSupportFontPropPage
    miBandChevronPressed = NegOneL
End Sub

Private Sub UserControl_InitProperties()
    Set moFont = Font_CreateDefault(Ambient.Font)
    mbThemeable = DEF_Themeable
    pCreate
    mbRedraw = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    pCreate
    mbRedraw = True
End Sub

Private Sub UserControl_Resize()
    pResize
End Sub

Private Sub UserControl_Terminate()
    Set moFontPage = Nothing
    pDestroy
    If mhFont Then moFont.ReleaseHandle mhFont
    ReleaseShellMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Function pBand(ByVal iIndex As Long) As cBand
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : return a cBand object representing a given rebar band.
'---------------------------------------------------------------------------------------
    Debug.Assert iIndex > NegOneL And iIndex < pBands_Count
    If iIndex > NegOneL Then
        Dim iId As Long
        iId = pBand_Info(iIndex, RBBIM_ID)
        Set pBand = New cBand
        pBand.fInit iId, pBands_FindId(iId), Me
    End If
    
End Function

Private Sub pChevronPushed(ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Raise the chevron push event from the NMREBARCHEVRON structure pointed to
'             by lParam.
'---------------------------------------------------------------------------------------
   
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    
    Dim tR As RECT
    CopyMemory tR, ByVal UnsignedAdd(lParam, NMREBARCHEVRON_lpRect), 16&
    MapWindowPoints mhWnd, UserControl.ContainerHwnd, tR, 2&
    
    Dim lhWnd As Long
    Dim lhWndNext As Long
    
    lhWnd = UserControl.ContainerHwnd
    lhWndNext = GetParent(lhWnd)
    
    Do While lhWndNext
        lhWnd = lhWndNext
        lhWndNext = GetParent(lhWnd)
    Loop
    
    If GetActiveWindow <> lhWnd Then SetActiveWindow lhWnd
    
    RaiseEvent ChevronPushed(pBand(MemOffset32(lParam, NMREBARCHEVRON_uBand)), _
                             ScaleX(tR.Left, vbPixels, vbContainerPosition), _
                             ScaleY(tR.Top, vbPixels, vbContainerPosition), _
                             ScaleX(tR.Right - tR.Left, vbPixels, vbContainerPosition), _
                             ScaleY(tR.bottom - tR.Top, vbPixels, vbContainerPosition))
                             
    bInHere = False
End Sub

Private Sub pChildResize(ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Move the toolbar usercontrol to a position exactly the same as the toolbar window.
'---------------------------------------------------------------------------------------
    Const NMREBARCHILDSIZE_wId As Long = 16
    Const NMREBARCHILDSIZE_lpRects As Long = 20
    Dim tR As RECT
    
    Dim liIndex As Long
    
    liIndex = pBands_FindId(MemOffset32(lParam, NMREBARCHILDSIZE_wId))
    
    If liIndex > NegOneL Then
        If Not mtBands(liIndex).oChild Is Nothing Then
        If TypeOf mtBands(liIndex).oChild Is ucToolbar Then
            CopyMemory tR, ByVal UnsignedAdd(lParam, NMREBARCHILDSIZE_lpRects), 16&
            
            With tR
                mtBands(liIndex).oChild.Move ScaleX(.Left, vbPixels, vbContainerPosition) + Extender.Left, _
                                             ScaleY(.Top, vbPixels, vbContainerPosition) + Extender.Top, _
                                             ScaleX(.Right - .Left, vbPixels, vbContainerSize), _
                                             ScaleY(.bottom - .Top, vbPixels, vbContainerSize)
                                             
            End With
        End If
        End If
    End If
End Sub

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Notify the Rebar window of a new font.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    If mhWnd Then
        Dim hFont As Long
        hFont = moFont.GetHandle()
        If hFont Then
            SendMessage mhWnd, WM_SETFONT, hFont, OneL
            If mhFont Then moFont.ReleaseHandle mhFont
            mhFont = hFont
        End If
        On Error GoTo 0
    End If
    Exit Sub
handler:
    Resume Next
End Sub

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Create the rebar window and install the subclass for notification.
'---------------------------------------------------------------------------------------
    pDestroy
    
    If Ambient.UserMode Then
    
        Dim liStyle As Long
        liStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or _
                  RBS_DBLCLKTOGGLE Or RBS_VARHEIGHT Or RBS_BANDBORDERS Or RBS_AUTOSIZE Or CCS_NODIVIDER
        
        Select Case pAlignment
        Case vbAlignRight, vbAlignLeft
            liStyle = liStyle Or CCS_RIGHT
        Case Else
            liStyle = liStyle Or CCS_TOP
        End Select
        
        
        Dim lsAnsi As String
        lsAnsi = StrConv(WC_REBAR & vbNullChar, vbFromUnicode)
        
        mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, liStyle, ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
        
        If mhWnd Then
            EnableWindowTheme mhWnd, mbThemeable
        
            If CheckCCVersion(5&) Then
                SendMessage mhWnd, RB_SETEXTENDEDSTYLE, ZeroL, RBS_EX_OFFICE9
            End If
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, UM_SIZEBAND)
            
            pResize
            pSetFont
        End If
    End If
    
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Destroy the rebar and subclass.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        fBands_Clear
        Subclass_Remove Me, UserControl.hWnd
        Subclass_Remove Me, mhWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Private Sub pResize()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Resize the rebar and account for changes in alignment.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        
        Static bInHere As Boolean
        If bInHere Then Exit Sub
        
        On Error Resume Next
        
        bInHere = True
        
        Static siLastAlignment As Long
        Dim liAlignment As Long
        
        liAlignment = pAlignment()
        
        If liAlignment <> siLastAlignment Then
            Select Case liAlignment
            Case vbAlignLeft
                SetWindowStyle mhWnd, CCS_LEFT, CCS_BOTTOM Or CCS_VERT
            Case vbAlignRight
                SetWindowStyle mhWnd, CCS_RIGHT, CCS_BOTTOM Or CCS_VERT
            Case vbAlignTop
                SetWindowStyle mhWnd, CCS_TOP, CCS_BOTTOM Or CCS_VERT
            Case vbAlignBottom
                SetWindowStyle mhWnd, CCS_BOTTOM, CCS_BOTTOM Or CCS_VERT
            End Select
            
            If CBool(liAlignment > vbAlignBottom) Xor CBool(siLastAlignment > vbAlignBottom) Then
                
                Dim i As Long
                Dim liIndex As Long
                Dim loChild As Object
                
                For i = ZeroL To miBandCount - OneL
                    If mtBands(i).iId Then
                        liIndex = SendMessage(mhWnd, RB_IDTOINDEX, mtBands(i).iId, ZeroL)
                        If liIndex > NegOneL Then
                            Set loChild = mtBands(i).oChild
                            If TypeOf loChild Is ucToolbar Then
                                pEvaluateToolbar loChild, CBool(pBand_Info(liIndex, RBBIM_STYLE) And RBBS_FIXEDSIZE)
                                mtBand.fMask = RBBIM_CHILDSIZE Or RBBIM_IDEALSIZE
                                SendMessage mhWnd, RB_SETBANDINFO, liIndex, VarPtr(mtBand)
                            End If
                        End If
                    End If
                Next
                
            End If
        End If
        
        SendMessage mhWnd, WM_SIZE, ZeroL, ZeroL
        
        siLastAlignment = liAlignment
        
        If siLastAlignment > vbAlignBottom Then
            Width = ScaleX(SendMessage(mhWnd, RB_GETBARHEIGHT, ZeroL, ZeroL), vbPixels, vbTwips)
        Else
            Height = ScaleX(SendMessage(mhWnd, RB_GETBARHEIGHT, ZeroL, ZeroL), vbPixels, vbTwips)
        End If
        
        If mbRedraw Then RaiseEvent Resize
        
        bInHere = False
        
        On Error GoTo 0
        
    End If
End Sub

Private Function pAlignment() As evbComCtlAlignment
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the current alignment.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    Dim loExtender As VBControlExtender
    Set loExtender = Extender
    'early bound!
    pAlignment = loExtender.Align
    On Error GoTo 0
    Exit Function
handler:
    On Error Resume Next
    'Maybe it's still supported.
    'even if it is, it could have different return values. Oh well.
    pAlignment = Extender.Align
    On Error GoTo 0
End Function

Private Sub pEvaluateChild(ByRef oBandChild As Object, ByVal iId As Long, ByVal iWidth As Long, ByVal iHeight As Long, ByVal bFixedSize As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Evaluate the band child.  ucToolbar controls are done automatically, and
'             an event is raised for other controls.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    
    If TypeOf oBandChild Is ucToolbar Then
        Dim loToolbar As ucToolbar
        Set loToolbar = oBandChild
        loToolbar.fRebar_Attach Me, iId
        pEvaluateToolbar oBandChild, bFixedSize
        'Toolbar usercontrol must be visible to receive mnemonics!
        oBandChild.ZOrder vbSendToBack
        'oBandChild.Visible = False
    Else
        With mtBand
            .hWndChild = oBandChild.hWnd
            
            If CBool(GetWindowLong(mhWnd, GWL_STYLE) And CCS_VERT) Then
                .cyChild = iWidth
               .cxIdeal = iHeight
            Else
                .cyChild = iHeight
                .cxIdeal = iWidth
            End If
            
            .cyMinChild = .cyChild
            .cyMaxChild = .cyChild
        End With
    End If
    
    Exit Sub
handler:
    Debug.Print "Child warning!", Err.Number, Err.Description
    Debug.Assert False
    Resume Next
End Sub

Private Sub pEvaluateToolbar(ByVal oToolbar As ucToolbar, ByVal bFixedSize As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Retreive dimensions for a ucToolbar control.
'---------------------------------------------------------------------------------------
    With mtBand
        oToolbar.fRebar_Vertical = CBool(GetWindowLong(mhWnd, GWL_STYLE) And CCS_VERT)
        .hWndChild = oToolbar.hWndToolbar
        .cxIdeal = oToolbar.fRebar_cxIdeal()
        .cyChild = oToolbar.fRebar_cyChild()
        .cyMaxChild = .cyChild
        .cyMinChild = .cyChild
        If bFixedSize Then .cxMinChild = .cxIdeal Else .cxMinChild = ZeroL
    End With
End Sub


Friend Property Get fToolbar_ChevronHitTest(ByVal oToolbar As ucToolbar, ByVal iId As Long, ByRef tP As POINT) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the point is over the toolbar's chevron.
'---------------------------------------------------------------------------------------
    
    Dim ltHT As RBHITTESTINFO
    LSet ltHT.pt = tP
    
    If mhWnd Then
        MapWindowPoints ZeroL, mhWnd, ltHT.pt, OneL
        SendMessage mhWnd, RB_HITTEST, ZeroL, VarPtr(ltHT)
        If SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL) = ltHT.iBand _
            Then fToolbar_ChevronHitTest = CBool(ltHT.Flags = RBHT_CHEVRON)
    End If
End Property

Friend Sub fToolbar_ShowChevron(ByVal oToolbar As ucToolbar, ByVal iId As Long)
    Dim liIndex As Long
    liIndex = SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL)
    If liIndex > NegOneL Then SendMessage mhWnd, RB_PUSHCHEVRON, liIndex, ZeroL
End Sub

Friend Sub fToolbar_Resize(ByVal oToolbar As ucToolbar, ByVal iId As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Notify the rebar if the toolbar needs to change size.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex As Long
        liIndex = SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL)
        If liIndex > NegOneL Then
            With mtBand
                .fMask = RBBIM_CHILDSIZE Or RBBIM_IDEALSIZE
                .cxIdeal = oToolbar.fRebar_cxIdeal()
                .cyChild = oToolbar.fRebar_cyChild()
                .cyMaxChild = .cyChild
                .cyMinChild = .cyChild
                SendMessage mhWnd, RB_SETBANDINFO, liIndex, VarPtr(mtBand)
            End With
        End If
    End If
End Sub


Friend Sub fBands_Enum_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the next cBand object in an enumeration.
'---------------------------------------------------------------------------------------
    tEnum.iIndex = tEnum.iIndex + OneL
    If tEnum.iIndex < pBands_Count() Then
        If tEnum.iControl <> miBandControl Then gErr vbccCollectionChangedDuringEnum, cBands
        Dim liId As Long
        Dim loItem As cBand
        
        liId = pBand_Info(tEnum.iIndex, RBBIM_ID)
        Set loItem = New cBand
        loItem.fInit liId, pBands_FindId(liId), Me
        Set vNextItem = loItem
        bNoMoreItems = False
    Else
        bNoMoreItems = True
    End If
End Sub


Friend Function fBands_Add( _
        ByRef oBandChild As Object, _
        ByRef sKey As String, _
        ByRef sText As String, _
        ByVal bUseChevron As Boolean, _
        ByVal bBreakLine As Boolean, _
        ByVal bGripper As Boolean, _
        ByVal bVisible As Boolean, _
        ByVal bFixedSize As Boolean, _
        ByVal iItemData As Long, _
        ByVal fWidth As Single, _
        ByVal fHeight As Single, _
        ByRef vBandBefore As Variant) _
            As cBand
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add a band to the rebar.
'---------------------------------------------------------------------------------------
    With mtBand
        Dim liOldParent As Long
        Dim liWidth As Long: liWidth = ScaleX(fWidth, vbContainerSize, vbPixels)
        Dim liHeight As Long: liHeight = ScaleY(fHeight, vbContainerSize, vbPixels)
        
        If liWidth = ZeroL Then
            On Error Resume Next
            liWidth = ScaleX(oBandChild.Width, vbContainerSize, vbPixels)
            On Error GoTo 0
        End If
        If liHeight = ZeroL Then
            On Error Resume Next
            liHeight = ScaleY(oBandChild.Height, vbContainerSize, vbPixels)
            On Error GoTo 0
        End If
        
        If LenB(sKey) Then
            If pBands_FindKey(sKey) <> NegOneL Then gErr vbccKeyAlreadyExists, cBands
        End If
        
        Dim liIndex As Long
        liIndex = NegOneL
        
        If Not IsMissing(vBandBefore) Then
            liIndex = pBands_GetIndex(vBandBefore)
            If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cBands
        End If
        
        .fMask = RBBIM_CHILD Or RBBIM_CHILDSIZE Or RBBIM_ID Or RBBIM_IDEALSIZE Or RBBIM_LPARAM Or RBBIM_STYLE
        
        .fStyle = RBBS_CHILDEDGE Or RBBS_FIXEDBMP
        If bBreakLine Then .fStyle = .fStyle Or RBBS_BREAK
        If Not bGripper Then .fStyle = .fStyle Or RBBS_NOGRIPPER
        If bUseChevron Then .fStyle = .fStyle Or RBBS_CHEVRON
        If Not bVisible Then .fStyle = .fStyle Or RBBS_HIDDEN
        If bFixedSize Then .fStyle = .fStyle Or RBBS_FIXEDSIZE
        .wID = NextItemIdShort()
        .lParam = iItemData
        
        pEvaluateChild oBandChild, .wID, liWidth, liHeight, bFixedSize
        liOldParent = GetParent(.hWndChild)
        
        .cx = .cxIdeal
        
        If LenB(sText) Then
            .fMask = .fMask Or RBBIM_TEXT
            Dim lsText As String
            lsText = StrConv(sText & vbNullChar, vbFromUnicode)
            MidB$(msTextBuffer, OneL, LenB(lsText)) = lsText
            MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = ZeroL
            .cch = LenB(msTextBuffer)
            .lpText = StrPtr(msTextBuffer)
        End If
        
        If SendMessage(mhWnd, RB_INSERTBAND, liIndex, VarPtr(mtBand)) Then
            Dim liDataIndex As Long
            pBands_Init sKey, oBandChild, liOldParent, liDataIndex, liWidth, liHeight
        Else
            If TypeOf oBandChild Is ucToolbar Then
                Dim loToolbar As ucToolbar
                Set loToolbar = oBandChild
                loToolbar.fRebar_Detach
            End If
        End If
        
        Set fBands_Add = New cBand
        fBands_Add.fInit .wID, liDataIndex, Me
    End With
    
    Incr miBandControl
End Function

Friend Sub fBands_Remove(ByRef vBand As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove a band from the rebar.
'---------------------------------------------------------------------------------------
    Dim i As Long
    Dim liDataIndex As Long
    i = pBands_GetIndex(vBand, , liDataIndex)
    If i = NegOneL Then gErr vbccKeyOrIndexNotFound, cBands
    pBands_Term liDataIndex
    If mhWnd Then SendMessage mhWnd, RB_DELETEBAND, i, ZeroL
    Incr miBandControl
End Sub

Friend Sub fBands_Clear()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove all bands from the rebar.
'---------------------------------------------------------------------------------------
    mbClearing = True
    Dim i As Long
    For i = miBandCount - OneL To ZeroL Step NegOneL
        pBands_Term i
    Next
    If mhWnd Then
        For i = pBands_Count To ZeroL Step NegOneL
            SendMessage mhWnd, RB_DELETEBAND, i, ZeroL
        Next
    End If
    
    mbClearing = False
    miBandCount = ZeroL
    Incr miBandControl
    
    If mbRedraw Then RaiseEvent Resize

End Sub

Friend Function fBands_Item(ByRef vBand As Variant) As cBand
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cBand object representing a given band.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    Dim liId As Long
    Dim liDataIndex As Long
    liIndex = pBands_GetIndex(vBand, liId, liDataIndex)
    If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cBands
    Set fBands_Item = New cBand
    fBands_Item.fInit liId, liDataIndex, Me
End Function

Friend Property Get fBands_Count() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of bands in the rebar.
'---------------------------------------------------------------------------------------
    fBands_Count = pBands_Count()
End Property

Friend Property Get fBands_Exists(ByRef vBand As Variant) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether a band exists in the rebar.
'---------------------------------------------------------------------------------------
    fBands_Exists = pBands_GetIndex(vBand) <> NegOneL
End Property

Friend Property Get fBands_Enum_Control() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an identity variable for the current cBands collection.
'---------------------------------------------------------------------------------------
    fBands_Enum_Control = miBandControl
End Property

Private Function pBands_GetIndex(ByRef vBand As Variant, Optional ByRef iId As Long, Optional ByRef iDataIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get an index to a band given its key, object or index.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    iId = ZeroL
    iDataIndex = NegOneL
    pBands_GetIndex = NegOneL
    
    If VarType(vBand) = vbString Then
        pBands_GetIndex = pBands_FindKey(CStr(vBand), iId, iDataIndex) - OneL
    Else
        If VarType(vBand) = vbObject Then
            If TypeOf vBand Is cBand Then
                Dim loBand As cBand
                Set loBand = vBand
                On Error Resume Next
                If loBand.fIsMine(Me) Then pBands_GetIndex = loBand.Index - OneL
                On Error GoTo 0
            Else
                For pBands_GetIndex = ZeroL To miBandCount - OneL
                    If mtBands(pBands_GetIndex).iId Then
                        If mtBands(pBands_GetIndex).oChild Is vBand Then Exit For
                    End If
                Next
                If pBands_GetIndex < miBandCount _
                    Then pBands_GetIndex = SendMessage(mhWnd, RB_IDTOINDEX, mtBands(pBands_GetIndex).iId, ZeroL)
            End If
        Else
            pBands_GetIndex = CLng(vBand) - OneL
        End If
        
        If pBands_GetIndex > NegOneL And pBands_GetIndex < pBands_Count() Then
            iId = pBand_Info(pBands_GetIndex, RBBIM_ID)
            iDataIndex = pBands_FindId(iId)
        End If
    End If
    
    If pBands_GetIndex < ZeroL Or pBands_GetIndex >= pBands_Count() _
        Or iId = ZeroL _
        Or iDataIndex < ZeroL Or iDataIndex >= miBandCount Then

handler:
        pBands_GetIndex = NegOneL
    
    End If
    
    On Error GoTo 0
    
End Function

Private Function pBands_FindId(ByVal iId As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Find a band by id.
'---------------------------------------------------------------------------------------
    If iId Then
        For pBands_FindId = ZeroL To miBandCount - OneL
            If mtBands(pBands_FindId).iId = iId Then Exit Function
        Next
    End If
    pBands_FindId = NegOneL
End Function

Private Function pBands_FindKey(ByRef sKey As String, Optional ByRef iId As Long, Optional ByRef iDataIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Find a band by key.
'---------------------------------------------------------------------------------------
    For iDataIndex = ZeroL To miBandCount - OneL
        If mtBands(iDataIndex).iId Then
            If StrComp(sKey, mtBands(iDataIndex).sKey) = ZeroL Then Exit For
        End If
    Next
    If iDataIndex < miBandCount Then
        If mhWnd Then
            iId = mtBands(iDataIndex).iId
            pBands_FindKey = SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL)
            If pBands_FindKey <> NegOneL Then Exit Function
        End If
    End If
    pBands_FindKey = NegOneL
    iId = ZeroL
    iDataIndex = NegOneL
End Function

Private Function pBands_Count() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of bands in the rebar.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        pBands_Count = SendMessage(mhWnd, RB_GETBANDCOUNT, ZeroL, ZeroL)
        'Debug.Assert pBands_Count = miBandCount
    End If
End Function

Private Sub pBands_Init(ByRef sKey As String, ByVal oBandChild As Object, ByVal iOldParent As Long, ByRef iDataIndex As Long, ByVal iWidth As Long, ByVal iHeight As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Initialize a band on the first available index in the data structure.
'             If no index is available, resize the array.
'---------------------------------------------------------------------------------------
    For iDataIndex = ZeroL To miBandCount - OneL
        If mtBands(iDataIndex).iId = ZeroL Then Exit For
    Next
    If iDataIndex = miBandCount Then
        Dim liUBound As Long
        liUBound = RoundToInterval(miBandCount)
        If liUBound > miBandUbound Then
            miBandUbound = liUBound
            ReDim Preserve mtBands(ZeroL To miBandUbound)
        End If
        miBandCount = miBandCount + OneL
    End If
    With mtBands(iDataIndex)
        .iId = mtBand.wID
        .sKey = sKey
        Set .oChild = oBandChild
        .hwndOldParent = iOldParent
        .iWidth = iWidth
        .iHeight = iHeight
    End With
End Sub

Private Sub pBands_Term(ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Clean up a band before it is to be removed.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim iId As Long
    iId = pBand_Info(iIndex, RBBIM_ID)
    If iId Then
        Dim liIndex As Long
        liIndex = pBands_FindId(iId)
        If liIndex > NegOneL Then
            With mtBands(liIndex)
                Debug.Assert IsWindow(.hwndOldParent)
                If IsWindow(.hwndOldParent) Then SetParent pBand_Info(iIndex, RBBIM_CHILD), .hwndOldParent
                .iId = ZeroL
                If TypeOf .oChild Is ucToolbar Then
                    Dim loToolbar As ucToolbar
                    Set loToolbar = .oChild
                    loToolbar.fRebar_Attach Nothing, ZeroL
                End If
                Set .oChild = Nothing
            End With
            If liIndex = miBandCount - OneL Then
                For liIndex = liIndex - OneL To ZeroL Step NegOneL
                    If mtBands(liIndex).iId Then Exit For
                Next
                miBandCount = (liIndex + OneL)
            End If
        End If
    End If
    On Error GoTo 0
End Sub


Friend Property Get fBand_Child(ByVal iId As Long, ByRef iIndex As Long) As Object
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the child object of the band.
'---------------------------------------------------------------------------------------
    If pBand_Verify(iId, iIndex) Then
        Set fBand_Child = mtBands(iIndex).oChild
    End If
End Property

Friend Property Get fBand_IdealWidth(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the ideal width that was previously specified for this band.
'---------------------------------------------------------------------------------------
    If pBand_Verify(iId, iIndex) Then
        fBand_IdealWidth = ScaleX(mtBands(iIndex).iWidth, vbPixels, vbContainerSize)
    End If
End Property

Friend Property Let fBand_IdealWidth(ByVal iId As Long, ByRef iIndex As Long, ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the ideal width for this band and resize the rebar.
'---------------------------------------------------------------------------------------
    If pBand_Verify(iId, iIndex) Then
        Dim liIndex As Long
        mtBand.fMask = RBBIM_CHILDSIZE Or RBBIM_IDEALSIZE
        liIndex = SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL)
        If liIndex > NegOneL Then
            With mtBands(iIndex)
                .iWidth = ScaleX(fNew, vbContainerSize, vbPixels)
                pEvaluateChild .oChild, .iId, .iWidth, .iHeight, CBool(pBand_Info(liIndex, RBBIM_STYLE) And RBBS_FIXEDSIZE)
            End With
            mtBand.fMask = RBBIM_CHILDSIZE Or RBBIM_IDEALSIZE
            SendMessage mhWnd, RB_SETBANDINFO, liIndex, VarPtr(mtBand)
        End If
    End If
End Property

Friend Property Get fBand_IdealHeight(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the ideal height that was previously specified for this band.
'---------------------------------------------------------------------------------------
    If pBand_Verify(iId, iIndex) Then
        fBand_IdealHeight = ScaleX(mtBands(iIndex).iHeight, vbPixels, vbContainerSize)
    End If
End Property

Friend Property Let fBand_IdealHeight(ByVal iId As Long, ByRef iIndex As Long, ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the ideal height for this band and resize the rebar.
'---------------------------------------------------------------------------------------
    If pBand_Verify(iId, iIndex) Then
        Dim liIndex As Long
        liIndex = SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL)
        If liIndex > NegOneL Then
            With mtBands(iIndex)
                .iHeight = ScaleX(fNew, vbContainerSize, vbPixels)
                pEvaluateChild .oChild, .iId, .iWidth, .iHeight, CBool(pBand_Info(liIndex, RBBIM_STYLE) And RBBS_FIXEDSIZE)
            End With
            mtBand.fMask = RBBIM_CHILDSIZE Or RBBIM_IDEALSIZE
            SendMessage mhWnd, RB_SETBANDINFO, liIndex, VarPtr(mtBand)
        End If
    End If
End Property

Friend Property Get fBand_Key(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the band's key.
'---------------------------------------------------------------------------------------
    If pBand_Verify(iId, iIndex) Then
        fBand_Key = mtBands(iIndex).sKey
    End If
End Property

Friend Property Get fBand_Text(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the band's text.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        fBand_Text = pBand_Text(liBandIndex)
    End If
End Property
Friend Property Let fBand_Text(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the band's text.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        pBand_Text(liBandIndex) = sNew
    End If
End Property

Friend Property Get fBand_Index(ByVal iId As Long, ByRef iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the index of the band.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        fBand_Index = liBandIndex
    End If
End Property
Friend Property Let fBand_Index(ByVal iId As Long, ByRef iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the index of the band.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        iNew = iNew - OneL
        If iNew < pBands_Count And iNew > NegOneL Then
            If SendMessage(mhWnd, RB_MOVEBAND, liBandIndex, iNew) Then Exit Property
        End If
        gErr vbccKeyOrIndexNotFound, cBand
    End If
End Property

Friend Property Get fBand_Visible(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the band is visible.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        fBand_Visible = Not pBand_Style(liBandIndex, RBBS_HIDDEN)
    End If
End Property
Friend Property Let fBand_Visible(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the band is visible.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        pBand_Style(liBandIndex, RBBS_HIDDEN) = Not bNew
    End If
End Property

Friend Property Get fBand_Gripper(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether a gripper is visible and the band can be dragged.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        fBand_Gripper = Not pBand_Style(liBandIndex, RBBS_NOGRIPPER)
    End If
End Property
Friend Property Let fBand_Gripper(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether a gripper is visible and the band can be dragged.
'---------------------------------------------------------------------------------------
    Dim liBandIndex As Long
    If pBand_Verify(iId, iIndex, liBandIndex) Then
        Dim liStyle As Long
        liStyle = pBand_Info(liBandIndex, RBBIM_STYLE)
        If CBool(liStyle And RBBS_NOGRIPPER) Xor (Not bNew) Then
            pBand_Info(liBandIndex, RBBIM_STYLE) = (liStyle And Not RBBS_NOGRIPPER) Or (RBBS_NOGRIPPER * (bNew + OneL))
            If Not CBool(liStyle And RBBS_HIDDEN) Then SendMessage mhWnd, RB_SHOWBAND, liBandIndex, OneL
        End If
    End If
End Property

Private Function pBand_Verify(ByVal iId As Long, ByRef iIndex As Long, Optional ByRef iBandIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Verify that a band still exists in the band collection.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        iBandIndex = SendMessage(mhWnd, RB_IDTOINDEX, iId, ZeroL)
        If iBandIndex > NegOneL Then
            If iIndex < miBandCount And iIndex > NegOneL Then
                pBand_Verify = (mtBands(iIndex).iId = iId)
                If Not pBand_Verify Then
                    iIndex = pBands_FindId(iId)
                    pBand_Verify = iIndex > NegOneL
                End If
            End If
        End If
    End If
    If Not pBand_Verify Then gErr vbccItemDetached, cBand
End Function

Private Property Let pBand_Text(ByVal iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of the band.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim ls As String
        ls = StrConv(sNew & vbNullChar, vbFromUnicode)
        MidB$(msTextBuffer, OneL, LenB(ls)) = ls
        MidB$(msTextBuffer, LenB(msTextBuffer), OneL) = vbNullChar
        With mtBand
            .fMask = RBBIM_TEXT
            .lpText = StrPtr(msTextBuffer)
            .cch = LenB(msTextBuffer)
        End With
        SendMessage mhWnd, RB_SETBANDINFO, iIndex, VarPtr(mtBand)
    End If
End Property
Private Property Get pBand_Text(ByVal iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the text of the band.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtBand
            .fMask = RBBIM_TEXT
            .lpText = StrPtr(msTextBuffer)
            .cch = LenB(msTextBuffer)
            SendMessage mhWnd, RB_GETBANDINFO, iIndex, VarPtr(mtBand)
            lstrToStringA .lpText, pBand_Text
        End With
    End If
End Property

Private Property Get pBand_Info(ByVal iIndex As Long, ByVal iMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value from the tagREBARBANDINFO structure of a band.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtBand
            .fMask = iMask
            If SendMessage(mhWnd, RB_GETBANDINFO, iIndex, VarPtr(mtBand)) Then
                If iMask = RBBIM_ID Then
                    pBand_Info = .wID
                ElseIf iMask = RBBIM_LPARAM Then
                    pBand_Info = .lParam
                ElseIf iMask = RBBIM_STYLE Then
                    pBand_Info = .fStyle
                ElseIf iMask = RBBIM_CHILD Then
                    pBand_Info = .hWndChild
                Else
                    Debug.Assert False
                End If
            End If
        End With
    End If
End Property
Private Property Let pBand_Info(ByVal iIndex As Long, ByVal iMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value in the tagREBARBANDINFO structure of a band.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtBand
            .fMask = iMask
            
            If iMask = RBBIM_ID Then
                .wID = iNew
            ElseIf iMask = RBBIM_LPARAM Then
                .lParam = iNew
            ElseIf iMask = RBBIM_STYLE Then
                .fStyle = iNew
            ElseIf iMask = RBBIM_CHILD Then
                .hWndChild = iNew
            Else
                Debug.Assert False
            End If
            
            SendMessage mhWnd, RB_SETBANDINFO, iIndex, VarPtr(mtBand)

        End With
    End If
End Property

Private Property Get pBand_Style(ByVal iIndex As Long, ByVal iStyle As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value indicating whether a band has a given mask.
'---------------------------------------------------------------------------------------
    pBand_Style = CBool(pBand_Info(iIndex, RBBIM_STYLE) And iStyle)
End Property
Private Property Let pBand_Style(ByVal iIndex As Long, ByVal iStyle As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether a band has a given mask.
'---------------------------------------------------------------------------------------
    Dim liStyle As Long
    liStyle = pBand_Info(iIndex, RBBIM_STYLE)
    If bNew Then liStyle = liStyle Or iStyle Else liStyle = liStyle And Not iStyle
    pBand_Info(iIndex, RBBIM_STYLE) = liStyle
End Property


Public Property Get Bands() As cBands
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a collection of bands.
'---------------------------------------------------------------------------------------
    Set Bands = New cBands
    Bands.fInit Me
End Property


Public Sub MaximixeBand(ByVal vBand As Variant, Optional ByVal bIdeal As Boolean = True)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Make the band as long as space permits.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex As Long
        liIndex = pBands_GetIndex(vBand)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucRebar
        SendMessage mhWnd, RB_MAXIMIZEBAND, liIndex, -bIdeal
    End If
End Sub

Public Sub MinimizeBand(ByVal vBand As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Make the band as short as space permits.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex As Long
        liIndex = pBands_GetIndex(vBand)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucRebar
        SendMessage mhWnd, RB_MINIMIZEBAND, liIndex, ZeroL
    End If
End Sub

Public Sub SetAlignment(ByVal iNew As evbComCtlAlignment)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the alignment of the control.  Calling this sub instead of setting the
'             align property allows the control to resize properly.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim vbe As VBControlExtender
    Set vbe = Extender
    Dim lbOldVisible As Boolean
    lbOldVisible = vbe.Visible
    vbe.Visible = False
    vbe.Align = iNew
    vbe.Visible = lbOldVisible
    pResize
    On Error GoTo 0
End Sub

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the font used by the control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property

Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the font used by the control.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing Then Set oNew = Font_CreateDefault(Ambient.Font)
    Set moFont = oNew
    pSetFont
    PropertyChanged PROP_Font
End Property

Public Property Get HitTest(ByVal x As Single, ByVal y As Single, Optional ByRef iFlags As eRebarHitTest) As cBand
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the band at the given index.
'---------------------------------------------------------------------------------------
    Dim ltHT As RBHITTESTINFO
    ltHT.pt.x = ScaleX(x, vbContainerPosition, vbPixels)
    ltHT.pt.y = ScaleY(y, vbContainerPosition, vbPixels)
    If mhWnd Then
        SendMessage mhWnd, RB_HITTEST, ZeroL, VarPtr(ltHT)
        iFlags = ltHT.Flags
        If iFlags <> rbarHitNowhere Then Set HitTest = pBand(ltHT.iBand)
    End If
End Property

Public Property Get Themeable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return a value indicating whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        PropertyChanged PROP_Themeable
        mbThemeable = bNew
        If mhWnd Then
            EnableWindowTheme mhWnd, mbThemeable
            pResize
        End If
    End If
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Get whether the the control redraws and raises the resize event.
'---------------------------------------------------------------------------------------
    Redraw = mbRedraw
End Property
Public Property Let Redraw(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set whether the the control redraws and raises the resize event.
'---------------------------------------------------------------------------------------
    If mbRedraw Xor bNew Then
        mbRedraw = bNew
        SendMessage mhWnd, WM_SETREDRAW, CLng(-mbRedraw), ZeroL
        If mbRedraw Then RaiseEvent Resize
    End If
End Property

Public Property Get hWndRebar() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the hwnd of the rebar.
'---------------------------------------------------------------------------------------
    If mhWnd Then hWndRebar = mhWnd
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property
