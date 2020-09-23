Attribute VB_Name = "mTest"
'---------------------------------------------------------------------------------------
'mTest.bas           3/31/05
'
'            PURPOSE:
'               General procedures.
'
'---------------------------------------------------------------------------------------
Option Explicit

Public Const ZeroL As Long = 0&, NegOneL As Long = -1&, OneL As Long = 1&, TwoL As Long = 2&

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
    Public Const VK_TAB As Long = &H9, VK_SHIFT As Long = &H10, VK_ESCAPE As Long = &H1B

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Const SWP_NOACTIVATE As Long = &H10, SWP_NOMOVE As Long = &H2, SWP_NOSIZE As Long = &H1, HWND_TOPMOST As Long = -1, HWND_NOTOPMOST As Long = -2

Public Function KeyIsDown( _
            ByVal iVirtKey As Long) _
                As Boolean
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a value indicating whether the given key is pressed.
'---------------------------------------------------------------------------------------
    KeyIsDown = CBool(GetAsyncKeyState(iVirtKey) And &H8000)
End Function

Public Sub OnTop(ByVal hWnd As Long, ByVal bVal As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Move a window to or from the foreground.
'---------------------------------------------------------------------------------------
    SetWindowPos hWnd, IIf(bVal, HWND_TOPMOST, HWND_NOTOPMOST), ZeroL, ZeroL, ZeroL, ZeroL, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function RandIcon(ByVal oImageList As cImageList) As Long
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Select a random icon from the imagelist.
'---------------------------------------------------------------------------------------
    RandIcon = Rnd * (oImageList.IconCount - OneL)
End Function

Public Property Get LstFlags(ByVal lst As ListBox) As Long
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Logically Or the itemdatas of the selected listbox items and return the result.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    For liIndex = ZeroL To lst.ListCount - OneL
        If lst.Selected(liIndex) Then LstFlags = lst.ItemData(liIndex) Or LstFlags
    Next
End Property

Public Property Let LstFlags(ByVal lst As ListBox, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Set the selected state of each listbox item based on whether its itemdata
'             is a bitmask contained in iNew.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    For liIndex = ZeroL To lst.ListCount - OneL
        lst.Selected(liIndex) = CBool(iNew And lst.ItemData(liIndex))
    Next
End Property

Public Function IsLoaded(ByVal oForm As Form) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return a value indicating whether the given form is loaded
'---------------------------------------------------------------------------------------
    Dim o As Form
    For Each o In Forms
        If o Is oForm Then Exit For
    Next
    IsLoaded = Not o Is Nothing
End Function
