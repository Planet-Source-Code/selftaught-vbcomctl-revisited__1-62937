%define _patch_InitialTick_       01BCCAABh             ;initial tickcount
%define _patch_iOwnerId_          02BCCAABh             ;owner data
%define _patch_oOwner_            03BCCAABh             ;iTimer object interface pointer
%define _patch_Callback_          04BCCAABh             ;address of the callback procedure
%define _patch_iTimerId_          05BCCAABh             ;timer identifier
%define _patch_pNextTimerProc_    06BCCAABh             ;next timer proc in the linked list

%define dwTime                    [esp+16]
%define idEvent                   [esp+12]
%define uMsg                      [esp+8]
%define hWnd                      [esp+4]

[bits 32]
    sub         dword dwTime, _patch_InitialTick_       ;subtract the initial tickcount from the current tickcount    
    mov         dword uMsg, _patch_iOwnerId_            ;copy the owner id to the stack
    mov         dword hWnd, _patch_oOwner_              ;copy the owner object pointer to the stack
    mov         eax, _patch_Callback_                   ;store the callback address
    jmp         eax                                     ;pass control to the callback procedure
    dd          _patch_iTimerId_                        ;timer identifier
    dd          _patch_pNextTimerProc_                  ;next timer proc in the linked list