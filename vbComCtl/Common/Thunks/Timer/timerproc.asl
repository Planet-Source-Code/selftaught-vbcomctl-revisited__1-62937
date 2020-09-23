     1                                  %define _patch_EbMode_            01BCCAABh             ;relative address to EbMode
     2                                  %define _patch_KillTimer_         02BCCAABh             ;relative address to KillTimer
     3                                  
     4                                  %define iTickCount                [esp+16]              ;milliseconds elapsed since timer start
     5                                  %define iTimerID                  [esp+12]              ;id returned by SetTimer
     6                                  %define iOwnerID                  [esp+8]               ;owner id
     7                                  %define oOwner                    [esp+4]               ;iTimer interface pointer
     8                                  %define iReturnAddress            [esp+0]               ;caller return address
     9                                  
    10                                  
    11                                  [bits 32]
    12 00000000 EB0D                        jmp short   _callback                               ;patched with NOOPs if in the IDE    
    13                                  
    14                                  _idecheck:                                              ;Check to see if the IDE is stopped or on a breakpoint
    15 00000002 E8                          db          0E8h                                    ;Far call op-code
    16 00000003 ABCABC01                    dd          _patch_EbMode_                          ;Call EbMode, the relative address to EbMode is patched at runtime
    17 00000007 3C02                        cmp         al, 2                                   ;If EbMode returns 2
    18 00000009 742E                        je short    _return                                 ;   The IDE is on a breakpoint
    19 0000000B 85C0                        test        eax, eax                                ;If EbMode returns 0
    20 0000000D 741D                        je short    _killtimer                              ;   The IDE has stopped
    21                                  
    22                                  _callback:
    23 0000000F 8B442408                    mov         eax, iOwnerID                           ;store the owner id
    24 00000013 8944240C                    mov         iTimerID, eax                           ;replace the timer id with the owner id
    25 00000017 8B442404                    mov         eax, oOwner                             ;store the owner object pointer
    26 0000001B 89442408                    mov         iOwnerID, eax                           ;replace the owner id with the owner object pointer
    27 0000001F 58                          pop         eax                                     ;pop the return address off the stack
    28 00000020 890424                      mov         [esp+0], eax                            ;move the return address back onto the stack
    29 00000023 8B442404                    mov         eax, [esp+4]                            ;get the owner object pointer
    30 00000027 8B00                        mov         eax, [eax]                              ;dereference pointer
    31 00000029 FF601C                      jmp         [eax+1Ch]                               ;pass control to iTimer::Proc
    32                                  
    33                                  _killtimer:                                             ;The IDE has stopped, kill the timer and return
    34 0000002C 8B44240C                    mov         eax, iTimerID                           ;get the timer ID
    35 00000030 50                          push        eax                                     ;push the timer ID
    36 00000031 31C0                        xor         eax, eax                                ;clear eax
    37 00000033 50                          push        eax                                     ;push 0 as the timer hwnd
    38 00000034 E8                          db          0E8h                                    ;Far call
    39 00000035 ABCABC02                    dd          _patch_KillTimer_                       ;call KillTimer, patched at runtime
    40                                  
    41                                  _return:                                                ;return control to calling function
    42 00000039 31C0                        xor         eax, eax                                ;clear eax
    43 0000003B C21000                      ret         16                                      ;return
