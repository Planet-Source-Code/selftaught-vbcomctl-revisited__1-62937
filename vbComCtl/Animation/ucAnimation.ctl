VERSION 5.00
Begin VB.UserControl ucAnimation 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucAnimation.ctx":0000
End
Attribute VB_Name = "ucAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucAnimation.ctl                  9/06/05
'
'           PURPOSE:
'               Implement ANIMATE_CLASS from ComCtl32.dll.
'
'               Load from a resource or a file.
'
'               Avi files must not have an audio stream.
'
'               "Transparent" does not work as expected. The upper-left pixel of the first frame is the
'               transparent color, but all pixels of that color are changed to the ColorBack property value,
'               and graphics behind the animation do not show through.
'
'               Timer property sets whether the control uses a timer or a thread,
'               but does not work under common controls 6. (always uses a timer).
'
'               Changing the property values of Timer, Transparent or Center will result
'               in the loss of any currently loaded avi.  You must then reload the avi
'               by calling LoadFromFile or LoadFromResource.
'
'==================================================================================================

Option Explicit

Private Const PROP_Transparent  As String = "Trans"
Private Const PROP_Center       As String = "Ctr"
Private Const PROP_AutoPlay     As String = "AutPl"
Private Const PROP_Timer        As String = "Tmr"
Private Const PROP_BackColor    As String = "BckClr"

Private Const DEF_Transparent   As Boolean = True
Private Const DEF_Center        As Boolean = True
Private Const DEF_AutoPlay      As Boolean = True
Private Const DEF_Timer         As Boolean = False
Private Const DEF_Backcolor     As Long = vbButtonFace

Private mhWnd           As Long

Private mbTransparent   As Boolean
Private mbCenter        As Boolean
Private mbAutoPlay      As Boolean
Private mbTimer         As Boolean

Private Sub UserControl_Initialize()
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Initialize the common control library and load the shell32 module to prevent
'             crashes at shutdown when linked to CC 6.0.
'---------------------------------------------------------------------------------------
    LoadShellMod
    InitCC ICC_ANIMATE_CLASS
End Sub

Private Sub UserControl_InitProperties()
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Initialize properties to their default values.
'---------------------------------------------------------------------------------------
    mbTransparent = DEF_Transparent
    mbCenter = DEF_Center
    mbAutoPlay = DEF_AutoPlay
    mbTimer = DEF_Timer
    UserControl.BackColor = DEF_Backcolor
    pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Read the property values from a previously saved instance.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    mbTransparent = PropBag.ReadProperty(PROP_Transparent, DEF_Transparent)
    mbCenter = PropBag.ReadProperty(PROP_Center, DEF_Center)
    mbAutoPlay = PropBag.ReadProperty(PROP_AutoPlay, DEF_AutoPlay)
    mbTimer = PropBag.ReadProperty(PROP_Timer, DEF_Timer)
    UserControl.BackColor = PropBag.ReadProperty(PROP_BackColor, DEF_Backcolor)
    On Error GoTo 0
    pCreate
End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Move the animation window to the same size as the usercontrol.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        MoveWindow mhWnd, ZeroL, ZeroL, ScaleWidth, ScaleHeight, OneL
    End If
End Sub

Private Sub UserControl_Terminate()
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Destroy the animation control and release the shell32.dll handle.
'---------------------------------------------------------------------------------------
    pDestroy
    ReleaseShellMod
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Write the property values to be saved between instances.
'---------------------------------------------------------------------------------------
    PropBag.WriteProperty PROP_Transparent, mbTransparent, DEF_Transparent
    PropBag.WriteProperty PROP_Center, mbCenter, DEF_Center
    PropBag.WriteProperty PROP_AutoPlay, mbAutoPlay, DEF_AutoPlay
    PropBag.WriteProperty PROP_Timer, mbTimer, DEF_Timer
    PropBag.WriteProperty PROP_BackColor, UserControl.BackColor, DEF_Backcolor
End Sub

Private Function pStyle() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return the window style according to our property values.
'---------------------------------------------------------------------------------------
    pStyle = WS_CHILD
    If mbTransparent Then pStyle = pStyle Or ACS_TRANSPARENT
    If mbCenter Then pStyle = pStyle Or ACS_CENTER
    If mbAutoPlay Then pStyle = pStyle Or ACS_AUTOPLAY
    If mbTimer Then pStyle = pStyle Or ACS_TIMER
End Function

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Create the sysanimate32 window.
'---------------------------------------------------------------------------------------
    pDestroy
    If Ambient.UserMode Then
        Dim lsAnsi As String
        lsAnsi = StrConv(WC_ANIMATION & vbNullChar, vbFromUnicode)
        mhWnd = CreateWindowEx(ZeroL, ByVal StrPtr(lsAnsi), ByVal ZeroL, pStyle(), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    End If
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Destroy the sysanimate32 window.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Public Function LoadFromFile(ByRef sFileName As String) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Wrap the ACM_OPEN message to open an avi file.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lsAnsi As String
        lsAnsi = StrConv(sFileName & vbNullChar, vbFromUnicode)
        LoadFromFile = CBool(SendMessage(mhWnd, ACM_OPENA, ZeroL, StrPtr(lsAnsi)))
        If LoadFromFile Then ShowWindow mhWnd, SW_SHOWNORMAL
    End If
End Function

Public Function LoadFromResource(ByVal hInstance As Long, ByVal iId As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Wrap the ACM_OPEN message to open an avi resource.
'             Works only on compiled resources, not *.res files included in the vbp.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        LoadFromResource = CBool(SendMessage(mhWnd, ACM_OPENA, hInstance, (iId And &HFFFF&)))
        If LoadFromResource Then ShowWindow mhWnd, SW_SHOWNORMAL
    End If
End Function

Public Function Play(Optional ByVal iFrameStart As Long = ZeroL, Optional ByVal iFrameEnd As Long = NegOneL, Optional ByVal iRepeatCount As Long = NegOneL) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Wrap the ACM_PLAY message to play an avi that was previously loaded.
'             Calling this function is not necessary when the AutoPlay property is True.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Play = CBool(SendMessage(mhWnd, ACM_PLAY, iRepeatCount, MakeLong(iFrameStart, iFrameEnd)))
        If Play Then ShowWindow mhWnd, SW_SHOWNORMAL
    End If
End Function

Public Function StopPlaying() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Wrap the ACM_STOP message to stop playing the avi.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        StopPlaying = CBool(SendMessage(mhWnd, ACM_STOP, ZeroL, ZeroL))
    End If
End Function

Public Property Get Transparent() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return whether the sysanimate32 window is transparent.
'---------------------------------------------------------------------------------------
    Transparent = mbTransparent
End Property
Public Property Let Transparent(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Set whether the sysanimate32 window is transparent.
'             'Transparent' is not really transparent.  It means that
'             the backcolor of the animation is changed to match the backcolor
'             of the usercontrol.  Graphics behind the animation will not be visible.
'
'             After changing this property, any currently loaded avi will be lost.
'---------------------------------------------------------------------------------------
    mbTransparent = bNew
    PropertyChanged PROP_Transparent
    If Ambient.UserMode Then pCreate
End Property

Public Property Get Timer() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return whether the animation control uses a timer as opposed
'             to a thread for playback.
'---------------------------------------------------------------------------------------
    Timer = mbTimer
End Property
Public Property Let Timer(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Set whether the animation control uses a timer as opposed
'             to a thread for playback.  XP always uses a timer regardless of this setting.
'
''             After changing this property, any currently loaded avi will be lost.
'---------------------------------------------------------------------------------------
    mbTimer = bNew
    PropertyChanged PROP_Timer
    If Ambient.UserMode Then pCreate
End Property

Public Property Get AutoPlay() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return whether an animation is played automatically upon loading.
'---------------------------------------------------------------------------------------
    AutoPlay = mbAutoPlay
End Property
Public Property Let AutoPlay(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Set whether an animation is played automatically upon loading.
'---------------------------------------------------------------------------------------
    mbAutoPlay = bNew
    If mhWnd Then SetWindowStyle mhWnd, ACS_AUTOPLAY * -bNew, ACS_AUTOPLAY
    PropertyChanged PROP_AutoPlay
End Property

Public Property Get Center() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return whether the animation is centered in the control.
'---------------------------------------------------------------------------------------
    Center = mbCenter
End Property
Public Property Let Center(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Set whether the animation is centered in the control.
'
'             After changing this property, any currently loaded avi will be lost.
'---------------------------------------------------------------------------------------
    mbCenter = bNew
    PropertyChanged PROP_Center
    If Ambient.UserMode Then pCreate
End Property

Public Property Get ColorBack() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Get the backcolor of the control.
'---------------------------------------------------------------------------------------
    ColorBack = UserControl.BackColor
End Property
Public Property Let ColorBack(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Set the backcolor of the control.
'---------------------------------------------------------------------------------------
    UserControl.BackColor = iNew
    PropertyChanged PROP_BackColor
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return the hWnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndAnimation() As Long
'---------------------------------------------------------------------------------------
' Date      : 9/06/05
' Purpose   : Return the hWnd of the animation control.
'---------------------------------------------------------------------------------------
    hWndAnimation = mhWnd
End Property
