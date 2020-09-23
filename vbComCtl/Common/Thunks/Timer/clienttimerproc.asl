     1                                  %define _patch_InitialTick_       01BCCAABh             ;initial tickcount
     2                                  %define _patch_iOwnerId_          02BCCAABh             ;owner data
     3                                  %define _patch_oOwner_            03BCCAABh             ;iTimer object interface pointer
     4                                  %define _patch_Callback_          04BCCAABh             ;address of the callback procedure
     5                                  %define _patch_iTimerId_          05BCCAABh             ;timer identifier
     6                                  %define _patch_pNextTimerProc_    06BCCAABh             ;next timer proc in the linked list
     7                                  
     8                                  %define dwTime                    [esp+16]
     9                                  %define idEvent                   [esp+12]
    10                                  %define uMsg                      [esp+8]
    11                                  %define hWnd                      [esp+4]
    12                                  
    13                                  [bits 32]
    14 00000000 816C2410ABCABC01            sub         dword dwTime, _patch_InitialTick_       ;subtract the initial tickcount from the current tickcount    
    15 00000008 C7442408ABCABC02            mov         dword uMsg, _patch_iOwnerId_            ;copy the owner id to the stack
    16 00000010 C7442404ABCABC03            mov         dword hWnd, _patch_oOwner_              ;copy the owner object pointer to the stack
    17 00000018 B8ABCABC04                  mov         eax, _patch_Callback_                   ;store the callback address
    18 0000001D FFE0                        jmp         eax                                     ;pass control to the callback procedure
    19 0000001F ABCABC05                    dd          _patch_iTimerId_                        ;timer identifier
    20 00000023 ABCABC06                    dd          _patch_pNextTimerProc_                  ;next timer proc in the linked list
