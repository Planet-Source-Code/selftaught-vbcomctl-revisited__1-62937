VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcRichEditStream_StringOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcRichEditStream_StringOut.cls                      5/1/05
'
'           PURPOSE:
'               Build a return string from chunks of a data stream supplied by the rich edit.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================
Option Explicit

Private moOwner As ucRichEdit       'callback for progress events
Private mpLinkedList As Long        'pointer to the first node in the linked list
Private mpLinkedListLast As Long    'pointer to the last node in the linked list
Private miTotal As Long             'total number of bytes received

Private Const NODE_NextPtr = ZeroL  'data offsets for each node in the linked list
Private Const NODE_Len = 4&
Private Const NODE_Data = 8&

Public Sub Proc(ByVal lPtrPbBuff As Long, ByVal iBytesRequested As Long, ByRef iBytesProcessed As Long, ByRef lReturn As Long)
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Allocate a node in the linked list and store the given chunk of data in it.
'---------------------------------------------------------------------------------------
    'iBytesProcessed = ZeroL
    'Exit Sub
    If iBytesRequested > ZeroL Then
        
        Dim lpNode As Long
        lpNode = MemAlloc(iBytesRequested + NODE_Data)
        
        Debug.Assert lpNode
        
        If lpNode Then
            
            MemOffset32(lpNode, NODE_NextPtr) = ZeroL
            MemOffset32(lpNode, NODE_Len) = iBytesRequested
            CopyMemory ByVal UnsignedAdd(lpNode, NODE_Data), ByVal lPtrPbBuff, iBytesRequested
            
            If mpLinkedListLast Then MemOffset32(mpLinkedListLast, NODE_NextPtr) = lpNode
            mpLinkedListLast = lpNode
            iBytesProcessed = iBytesRequested
            If mpLinkedList = ZeroL Then mpLinkedList = mpLinkedListLast
            miTotal = miTotal + iBytesRequested
            moOwner.fStream_OutProgress miTotal
            lReturn = ZeroL
            
        Else
            
            iBytesProcessed = ZeroL
            lReturn = vbccOutOfMemory
            
        End If

    Else
        iBytesProcessed = ZeroL
        
    End If
End Sub

Public Function Init(ByVal oOwner As ucRichEdit) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Store a reference to the callback object.
'             Store a pointer to the bstr into which the stream is copied.
'---------------------------------------------------------------------------------------
    Debug.Assert mpLinkedList = ZeroL
    
    Set moOwner = oOwner
    miTotal = ZeroL
    mpLinkedList = ZeroL
    mpLinkedListLast = ZeroL
    Init = True
    
End Function

Public Sub GetStreamResult(ByRef sOut As String)
'---------------------------------------------------------------------------------------
' Date      : 5/1/05
' Purpose   : Build the return string and release the memory allocated for the linked list.
'---------------------------------------------------------------------------------------
    Dim liTotalLen As Long
    Dim lpNode As Long
    
    lpNode = mpLinkedList
    Do While lpNode
        liTotalLen = liTotalLen + MemOffset32(lpNode, NODE_Len)
        lpNode = MemOffset32(lpNode, NODE_NextPtr)
    Loop
    
    sOut = vbNullString
    
    If liTotalLen > ZeroL Then
        Dim lpString As Long
        lpString = MemAllocFromString(ZeroL, liTotalLen)
        
        Debug.Assert lpString
                
        If lpString Then
            
            Dim liLen As Long
            Dim lpStringPos As Long
            lpStringPos = lpString
            
            lpNode = mpLinkedList
            Do While lpNode
                liLen = MemOffset32(lpNode, NODE_Len)
                CopyMemory ByVal lpStringPos, ByVal UnsignedAdd(lpNode, NODE_Data), liLen
                lpStringPos = UnsignedAdd(lpStringPos, liLen)
                lpNode = MemOffset32(lpNode, NODE_NextPtr)
            Loop
            
            mComCtlShared.lstrToStringA lpString, sOut, liTotalLen
            MemFree lpString
            pReleaseMemory
            
        Else
            
            pReleaseMemory
            gErr vbccOutOfMemory, "ucRichEdit"
            
        End If
    End If
End Sub

Private Sub pReleaseMemory()
    Dim lpNode As Long
    Dim lpNodeNext As Long
    
    lpNode = mpLinkedList
    Do While lpNode
        lpNodeNext = MemOffset32(lpNode, NODE_NextPtr)
        MemFree lpNode
        lpNode = lpNodeNext
    Loop
    
    mpLinkedList = ZeroL
    
End Sub

Private Sub Class_Terminate()
    Debug.Assert mpLinkedList = ZeroL
    pReleaseMemory
End Sub
