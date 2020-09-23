VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcRichEditStream_FileIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcRichEditStream_FileIn.cls                      5/1/05
'
'           PURPOSE:
'               Open a file and copy its data into a richedit stream.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================
Option Explicit

Private mhFile As Long
Private moOwner As ucRichEdit

Private miProgress As Long
Private miTotal As Long

Public Sub Proc(ByVal lPtrPbBuff As Long, ByVal iBytesRequested As Long, ByRef iBytesProcessed As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Read a chunk from the file to the given destination.
'---------------------------------------------------------------------------------------
    lReturn = Not CBool(ReadFile(mhFile, ByVal lPtrPbBuff, iBytesRequested, iBytesProcessed, ByVal ZeroL))
    Debug.Assert lReturn = ZeroL
    If lReturn = ZeroL Then
        miProgress = miProgress + iBytesProcessed
        moOwner.fStream_InProgress miProgress, miTotal
    End If
End Sub

Public Function Init(ByVal oOwner As ucRichEdit, sFileName As String) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Store the callback object, open the file and get the total size.
'---------------------------------------------------------------------------------------
    Set moOwner = oOwner
    
    Dim lsAnsi As String
    lsAnsi = StrConv(sFileName & vbNullChar, vbFromUnicode)
    
    mhFile = CreateFile(ByVal StrPtr(lsAnsi), GENERIC_READ, FILE_SHARE_READ, ByVal ZeroL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, ZeroL)
    If mhFile <> INVALID_HANDLE_VALUE Then
        miTotal = GetFileSize(mhFile, ZeroL)
        miProgress = ZeroL
        'file sizes >= 2^31 not supported
        Init = Not CBool(miTotal And &H80000000)
        If Not Init Then CloseHandle mhFile
    End If
End Function

Private Sub Class_Terminate()
    If mhFile <> INVALID_HANDLE_VALUE Then CloseHandle mhFile
End Sub
