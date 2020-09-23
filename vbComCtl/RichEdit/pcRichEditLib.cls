VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcRichEditLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcRichEditLib.cls                      5/1/05
'
'           PURPOSE:
'               Load the available richedit library.  One might think that the library should
'               be freed when the last richedit is destroyed and reloaded when another richedit
'               control is created, but apparently this causes a memory leak. So we load it
'               once and keep it loaded until the component is unloaded.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================
Option Explicit

Private mhLib As Long
Private mbVersion2 As Boolean

Public Function Init() As Boolean
    
    If mhLib = ZeroL Then
        Dim lsAnsi As String
        lsAnsi = StrConv("RICHED20.DLL" & vbNullChar, vbFromUnicode)
        mhLib = LoadLibrary(ByVal StrPtr(lsAnsi))
        mbVersion2 = CBool(mhLib)
        If Not mbVersion2 Then
            lsAnsi = StrConv("RICHED32.DLL" & vbNullChar, vbFromUnicode)
            mhLib = LoadLibrary(ByVal StrPtr(lsAnsi))
        End If
        
    End If
    
    Init = mbVersion2
   
End Function

Private Sub Class_Terminate()
    If mhLib Then FreeLibrary mhLib
End Sub
