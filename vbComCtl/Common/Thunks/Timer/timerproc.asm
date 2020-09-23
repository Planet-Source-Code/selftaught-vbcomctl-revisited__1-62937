%define _patch_EbMode_            01BCCAABh             ;relative address to EbMode
%define _patch_KillTimer_         02BCCAABh             ;relative address to KillTimer

%define iTickCount                [esp+16]              ;milliseconds elapsed since timer start
%define iTimerID                  [esp+12]              ;id returned by SetTimer
%define iOwnerID                  [esp+8]               ;owner id
%define oOwner                    [esp+4]               ;iTimer interface pointer
%define iReturnAddress            [esp+0]               ;caller return address


[bits 32]
    jmp short   _callback                               ;patched with NOOPs if in the IDE    

_idecheck:                                              ;Check to see if the IDE is stopped or on a breakpoint
    db          0E8h                                    ;Far call op-code
    dd          _patch_EbMode_                          ;Call EbMode, the relative address to EbMode is patched at runtime
    cmp         al, 2                                   ;If EbMode returns 2
    je short    _return                                 ;   The IDE is on a breakpoint
    test        eax, eax                                ;If EbMode returns 0
    je short    _killtimer                              ;   The IDE has stopped

_callback:
    mov         eax, iOwnerID                           ;store the owner id
    mov         iTimerID, eax                           ;replace the timer id with the owner id
    mov         eax, oOwner                             ;store the owner object pointer
    mov         iOwnerID, eax                           ;replace the owner id with the owner object pointer
    pop         eax                                     ;pop the return address off the stack
    mov         [esp+0], eax                            ;move the return address back onto the stack
    mov         eax, [esp+4]                            ;get the owner object pointer
    mov         eax, [eax]                              ;dereference pointer
    jmp         [eax+1Ch]                               ;pass control to iTimer::Proc

_killtimer:                                             ;The IDE has stopped, kill the timer and return
    mov         eax, iTimerID                           ;get the timer ID
    push        eax                                     ;push the timer ID
    xor         eax, eax                                ;clear eax
    push        eax                                     ;push 0 as the timer hwnd
    db          0E8h                                    ;Far call
    dd          _patch_KillTimer_                       ;call KillTimer, patched at runtime

_return:                                                ;return control to calling function
    xor         eax, eax                                ;clear eax
    ret         16                                      ;return