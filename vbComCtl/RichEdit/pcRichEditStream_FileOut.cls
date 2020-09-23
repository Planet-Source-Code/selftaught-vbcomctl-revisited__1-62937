VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcRichEditStream_FileOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcRichEditStream_FileOut.cls                      5/1/05
'
'           PURPOSE:
'               Copy a stream from a richedit to a file.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================
Option Explicit

Private moOwner As ucRichEdit
Private mhFile As Long
Private miTotal As Long

Public Sub Proc(ByVal lPtrPbBuff As Long, ByVal iBytesRequested As Long, ByRef iBytesProcessed As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Write a chunk of data from the given stream into the file.
'---------------------------------------------------------------------------------------
    lReturn = Not CBool(WriteFile(mhFile, ByVal lPtrPbBuff, iBytesRequested, iBytesProcessed, ByVal ZeroL))
    Debug.Assert lReturn = ZeroL
    If lReturn = ZeroL Then
        miTotal = miTotal + iBytesProcessed
        moOwner.fStream_OutProgress miTotal
    End If
End Sub

Public Function Init(ByVal oOwner As ucRichEdit, sFileName As String) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Store the callback object and open the file.
'---------------------------------------------------------------------------------------
    Set moOwner = oOwner
    
    Dim lsAnsi As String
    lsAnsi = StrConv(sFileName & vbNullChar, vbFromUnicode)
    mhFile = CreateFile(ByVal StrPtr(lsAnsi), GENERIC_WRITE, ZeroL, ByVal ZeroL, CREATE_ALWAYS, FILE_FLAG_SEQUENTIAL_SCAN, ZeroL)
    miTotal = ZeroL
    Init = CBool(mhFile <> INVALID_HANDLE_VALUE)
End Function

Private Sub Class_Terminate()
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Close the file.
'---------------------------------------------------------------------------------------
    If mhFile <> INVALID_HANDLE_VALUE Then CloseHandle mhFile
End Sub
