%define _patch_iHookType_               01BCCAABh           ;Relative address of the EbMode function
%define _patch_hHook_                   02BCCAABh           ;Hook handle for UnhookWindowsHookEx
%define _patch_oOwner_                  03BCCAABh           ;Hook Type  
%define _patch_Callback_                04BCCAABh           ;Addressof HookProc
%define _patch_pClientHookProcNext_     05BCCAABh           ;Next proc in the linked list

[bits 32]
    pop  eax                                                ;pop the return address off the stack
    push dword _patch_iHookType_                            ;push the hook type
    push dword _patch_hHook_                                ;push the hook handle
    push dword _patch_oOwner_                               ;push the owner iHook interface
    push eax                                                ;push the return address back on the stack
    mov  eax, _patch_Callback_                              ;store the return address back on the stack
    jmp  eax                                                ;pass control to the callback procedure
    dd   _patch_pClientHookProcNext_                        ;Next proc in the linked list