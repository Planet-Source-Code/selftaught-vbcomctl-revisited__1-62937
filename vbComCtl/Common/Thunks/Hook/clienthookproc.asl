     1                                  %define _patch_iHookType_               01BCCAABh           ;Relative address of the EbMode function
     2                                  %define _patch_hHook_                   02BCCAABh           ;Hook handle for UnhookWindowsHookEx
     3                                  %define _patch_oOwner_                  03BCCAABh           ;Hook Type  
     4                                  %define _patch_Callback_                04BCCAABh           ;Addressof HookProc
     5                                  %define _patch_pClientHookProcNext_     05BCCAABh           ;Next proc in the linked list
     6                                  
     7                                  [bits 32]
     8 00000000 58                          pop  eax                                                ;pop the return address off the stack
     9 00000001 68ABCABC01                  push dword _patch_iHookType_                            ;push the hook type
    10 00000006 68ABCABC02                  push dword _patch_hHook_                                ;push the hook handle
    11 0000000B 68ABCABC03                  push dword _patch_oOwner_                               ;push the owner iHook interface
    12 00000010 50                          push eax                                                ;push the return address back on the stack
    13 00000011 B8ABCABC04                  mov  eax, _patch_Callback_                              ;store the return address back on the stack
    14 00000016 FFE0                        jmp  eax                                                ;pass control to the callback procedure
    15 00000018 ABCABC05                    dd   _patch_pClientHookProcNext_                        ;Next proc in the linked list
