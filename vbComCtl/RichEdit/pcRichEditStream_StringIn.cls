VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcRichEditStream_StringIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcRichEditStream_StringIn.cls      5/1/05
'
'           PURPOSE:
'             Copy a string into a richedit stream in chunks.
'
'           LINEAGE:
'             N/A
'
'==================================================================================================

Option Explicit

Private moOwner As ucRichEdit
Private msANSI As String
Private mpString As Long
Private miProgress As Long
Private miTotal As Long

Public Sub Proc(ByVal lPtrPbBuff As Long, ByVal iBytesRequested As Long, ByRef iBytesProcessed As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Copy a chunk of data into the stream.
'---------------------------------------------------------------------------------------
    Dim liLen As Long
    liLen = miTotal - miProgress
    If iBytesRequested > liLen Then iBytesRequested = liLen
    
    If iBytesRequested > ZeroL Then
        Debug.Assert mpString
        CopyMemory ByVal lPtrPbBuff, ByVal UnsignedAdd(mpString, miProgress), iBytesRequested
        miProgress = miProgress + iBytesRequested
        moOwner.fStream_InProgress miProgress, miTotal
    Else
        iBytesRequested = ZeroL
    End If
    
    iBytesProcessed = iBytesRequested
    
    lReturn = ZeroL
End Sub

Public Function Init(ByVal oOwner As ucRichEdit, sString As String) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Store the length of and a pointer the string in ansi format.
'---------------------------------------------------------------------------------------
    Set moOwner = oOwner
    msANSI = StrConv(sString, vbFromUnicode)
    mpString = StrPtr(msANSI)
    miTotal = LenB(msANSI)
    miProgress = ZeroL
    Init = True
End Function
