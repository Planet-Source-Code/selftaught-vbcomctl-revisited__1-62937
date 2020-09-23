%define _patch_pPropString_                  01BCCAABh          ;pointer to null-terminated ANSI string for the window property
%define _patch_GetProp_                      02BCCAABh          ;Relative address to User32.GetPropA
%define _patch_CallWindowProcA_              03BCCAABh          ;Relative address to User32.CallWindowProcA
%define _patch_SetWindowLong_                04BCCAABh          ;Relative address to User32.GetWindowLongA
%define _patch_EbMode_                       05BCCAABh          ;Relative address to vba6.EbMode

%define _lParam_                             [ebp+20]           ;windowproc lParam
%define _wParam_                             [ebp+16]           ;windowproc wParam
%define _uMsg_                               [ebp+12]           ;windowproc uMsg
%define _hWnd_                               [ebp+8]            ;windowproc hWnd
%define _RetAddr_                            [ebp+4]            ;return address
%define _ebp_storage_                        [ebp+0]            ;ebp stored here
%define _ecx_storage_                        [ebp-4]            ;ecx stored here
%define _edi_storage_                        [ebp-8]            ;edi stored here
%define _edx_storage_                        [ebp-12]           ;edx stored here
%define _esi_storage_                        [ebp-16]           ;esi stored here
%define _lReturn_                            [ebp-20]           ;local lReturn
%define _bHandled_                           [ebp-24]           ;local bHandled
%define _pSubClientNext_                     [ebp-28]           ;local pSubClientNext
%define _pWndProcPrev_                       [ebp-32]           ;local pWndProcPrev

%define GWL_WNDPROC                          -4                 ;passed to User32.SetWindowLongA to unsubclass

%define _Subclass_pWndProcPrev_Offset_       0                  ;offset to the previous windowproc in the subclass data structure
%define _Subclass_pSubClientNext_Offset_     4                  ;offset to the next subclass data structure
%define _Subclass_pObject_Offset_            8                  ;offset to the iSubclass interface pointer
%define _Subclass_pMsgTableA_Offset_         12                 ;offset to the before message table
%define _Subclass_pMsgTableB_Offset_         16                 ;offset to the after message table

%define _ALL_MESSAGES_                       -1                 ;special value for the message table pointer indicating that all messages call back

[bits 32]
    push        ebp                                             ;store ebp
    mov         ebp, esp                                        ;setup stack frame
    push        ecx                                             ;store ecx
    push        edi                                             ;store edi
    push        edx                                             ;store edx
    push        esi                                             ;store esi
    xor         esi, esi                                        ;esi is now 0
    push        esi                                             ;allocate/initialize lReturn
    push        esi                                             ;allocate/initialize bHandled
    push        esi                                             ;allocate/initialize pSubClientNext
    push        esi                                             ;allocate/initialize pWndProcPrev

    call        _GetProp_                                       ;try to get the pointer to subclass data
    jz          _Return_                                        ;if this failed, we're SOL, return
    
    mov         edx, [eax+_Subclass_pWndProcPrev_Offset_]       ;get the previous wndproc address
    mov         _pWndProcPrev_, edx                             ;store the previous wndproc address
    mov         edx, eax                                        ;get the pointer to the first subclass client data
    
    mov         esi, _Subclass_pMsgTableB_Offset_               ;set the message table offset to use in this callback
    call        _Callback_                                      ;make the before callbacks
    
    mov         eax, _bHandled_                                 ;store bHandled
    test        eax, eax                                        ;test bHandled
    jnz         _Return_                                        ;if bHandled == True then return
    
    call        _Original_                                      ;call the original window procedure
    
    call        _GetProp_                                       ;try to get the pointer to subclass data
    jz          _Return_                                        ;if this failed, subclass was removed during a previous callback
    
    mov         edx, eax                                        ;get the pointer to the first subclass client data
    mov         esi, _Subclass_pMsgTableA_Offset_               ;set the message table offset to use in this callback
    call        _Callback_                                      ;make the after callbacks
    
_Return_:
    mov         eax, _lReturn_                                  ;store the return value
    add         esp, 16                                         ;deallocate pWndProcPrev, pSubClientNext, bHandled and lReturn
    pop         esi                                             ;restore esi
    pop         edx                                             ;restore edx
    pop         edi                                             ;restore edi
    pop         ecx                                             ;restore ecx
    pop         ebp                                             ;restore ebp
    ret         16                                              ;return and deallocate lParam, wParam, uMsg and hWnd
    
_GetProp_:
    push        dword _patch_pPropString_                       ;push the pointer to the property string
    push        dword _hWnd_                                    ;push the hwnd
    db          0e8h                                            ;far call opcode
    dd          _patch_GetProp_                                 ;relative address to User32.GetPropA - Patched at runtime
    test        eax, eax                                        ;test the return value
    ret                                                         ;return to caller
    
_Find_Message_:
    test        edi, edi                                        ;test the message table pointer
    jz          _Find_Message_Not_Found_                        ;if 0, no messages call back
    cmp         edi, _ALL_MESSAGES_                             ;test the message table pointer for -1
    je          _Find_Message_Found_                            ;if -1, all messages call back
    
    mov         ecx, [edi]                                      ;store the message table count (the first DWORD)
    add         edi, 4                                          ;increment the message table pointer past the count
    mov         eax, _uMsg_                                     ;store the message we will look for
    repne       scasd                                           ;scan the table
    jne         _Find_Message_Not_Found_                        ;if message was not found, jump
    
_Find_Message_Found_:
    xor         eax, eax                                        ;clear eax
    dec         eax                                             ;eax is now -1
    test        eax, eax                                        ;set the flags
    ret                                                         ;return to caller
    
_Find_Message_Not_Found_:
    xor         eax, eax                                        ;clear eax
    test        eax, eax                                        ;set the flags
    ret                                                         ;return to caller
    
_Original_:
    push        dword _lParam_                                  ;push byval lParam
    push        dword _wParam_                                  ;push byval wParam
    push        dword _uMsg_                                    ;push byval uMsg
    push        dword _hWnd_                                    ;push byval hWnd
    mov         eax,  _pWndProcPrev_                            ;get the original wndproc address
    push        eax                                             ;push the original wndproc address
    db          0e8h                                            ;far call opcode
    dd          _patch_CallWindowProcA_                         ;relative address to User32.CallWindowProcA - patched at runtime
    mov         _lReturn_, eax                                  ;store the return value
    ret
    
_Ide_Stop_:
    push        dword _pWndProcPrev_                            ;push the original wndproc address
    push        dword GWL_WNDPROC                               ;push GWL_WNDPROC
    push        dword _hWnd_                                    ;push the hwnd
    db          0e8h                                            ;far call opcode
    dd          _patch_SetWindowLong_                           ;relative address to User32.SetWindowLongA - patched at runtime
    
_Ide_Break_:
    pop         eax                                             ;pop the local return address
    cmp         esi, _Subclass_pMsgTableB_Offset_               ;find out if the original procedure has already been called
    jne         _Return_                                        ;if it has, return
    call        _Original_                                      ;if it has not, call it
    jmp short   _Return_                                        ;then return
    
_Callback_:
    jmp short   _No_Ide_Check_                                  ;skip the ide check - patched with NOOPs if in the ide
    
    db          0E8h                                            ;far call opcode
    dd          _patch_EbMode_                                  ;relative address to vba6.EbMode - patched at runtime
    cmp         al, 2                                           ;test for return value 2
    je          _Ide_Break_                                     ;if return value is 2, call the original procedure
    test        eax, eax                                        ;test for return value 0
    je          _Ide_Stop_                                      ;if return value is 0, call the original procedure and unsubclass
    
_No_Ide_Check_:
    mov         edi, [edx+_Subclass_pSubClientNext_Offset_]     ;get the pointer to the next subclass client data
    mov         _pSubClientNext_, edi                           ;store the pointer to the next subclass client data
    mov         edi, [edx+esi]                                  ;get the pointer to the current subclass client message table
    call        _Find_Message_                                  ;try to find the message in the table
    jz          _Callback_Skip_                                 ;if the message was not found, skip the callback
    
    push        dword _lParam_                                  ;push byval lParam
    push        dword _wParam_                                  ;push byval wParam
    push        dword _uMsg_                                    ;push byval uMsg
    push        dword _hWnd_                                    ;push byval hWnd
    lea         eax, _lReturn_                                  ;get addressof lReturn
    push        eax                                             ;push byref lReturn
    
    cmp         esi, _Subclass_pMsgTableA_Offset_               ;are we making an after or before callback?
    je          _Callback_After_                                ;if after, jump
    
_Callback_Before_:
    lea         eax, _bHandled_                                 ;get addressof bHandled
    push        eax                                             ;push byref bHandled
    
    mov         eax, [edx+_Subclass_pObject_Offset_]            ;get the pointer to the iSubclass interface
    push        eax                                             ;push the pointer to the iSubclass interface
    mov         eax, [eax]                                      ;get a pointer to the vtable
    call        dword [eax+1Ch]                                 ;make the before call, vtable offset 1Ch
    
    mov         eax, _bHandled_                                 ;store bHandled
    test        eax, eax                                        ;test bHandled
    jnz         _Callback_Return_                               ;if bHandled == True, return
    jmp short   _Callback_Skip_                                 ;skip the after callback for now
    
_Callback_After_:
    mov         eax, [edx+_Subclass_pObject_Offset_]            ;get the pointer to the iSubclass interface
    push        eax                                             ;push the pointer to the iSubclass interface
    mov         eax, [eax]                                      ;get a pointer to the vtable
    call        dword [eax+20h]                                 ;make the call, vtable offset 20h
    
_Callback_Skip_:
    mov         edx, _pSubClientNext_                           ;store the pointer to the next subclass client data
    test        edx, edx                                        ;set the flags
    jnz         _Callback_                                      ;if there is another subclass client, start over
    
_Callback_Return_:
    ret                                                         ;return to caller