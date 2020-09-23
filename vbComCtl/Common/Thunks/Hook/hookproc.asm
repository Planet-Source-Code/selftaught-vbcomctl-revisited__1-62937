%define _patch_EbMode_              01BCCAABh   ;relative address to vba6.EbMode
%define _patch_CallNextHookEx_      02BCCAABh   ;relative address to User32.CallNextHookEx
%define _patch_UnhookWindowsHookEx_ 03BCCAABh   ;relative address to User32.UnHookWindowsHookEx

%define lParam                      [ebp+40]    ;hookproc lparam
%define wParam                      [ebp+36]    ;hookproc wparam
%define nCode                       [ebp+32]    ;hookproc nCode
%define iHookType                   [ebp+28]    ;hook type ie WH_MOUSE, etc.
%define hHook                       [ebp+24]    ;hook handle
%define oOwner                      [ebp+20]    ;iHook interface
%define pAddrReturn                 [ebp+16]    ;return address
%define ebpStorage                  [ebp+12]    ;ebp stored here
%define esiStorage                  [ebp+8]     ;esi stored here
%define lReturn                     [ebp+4]     ;local lReturn
%define bHandled                    [ebp+0]     ;local bHandled

%define iHook_Before                1Ch         ;vtable offset to iHook::Before
%define iHook_After                 20h         ;vtable offset to iHook::After

[bits 32]
    
    push        ebp                             ;store ebp
    push        esi                             ;store esi
    xor         ebp, ebp                        ;clear ebp
    push        ebp                             ;allocate lReturn
    push        ebp                             ;allocate bHandled
    mov         ebp, esp                        ;setup stack frame

    jmp short   _noide_                         ;patched with NOOPs if in the ide
    
    db          0E8h                            ;Far call
    dd          _patch_EbMode_                  ;relative address to vba6.EbMode - patched at runtime
    cmp         al, 2                           ;test return value for 2
    je          _Ide_Break_                     ;If 2, IDE is on a breakpoint, just call the next hook
    test        eax, eax                        ;test return value for 0
    je          _Ide_Stop_                      ;If 0, IDE has stopped, unhook

_noide_:    
    mov         esi, iHook_Before               ;first call iHookProc_Before
    call        _Callback_                      ;make the before call
    
    mov         esi, bHandled                   ;store bHandled
    test        esi, esi                        ;set flags
    jnz         _Return_                        ;if bHandled == True, return
    
    call        _NextHook_                      ;call the next hook
    
    mov         esi, iHook_After                ;next call iHookProc_After
    call        _Callback_                      ;make the after call
    
_Return_:
    mov         eax, lReturn                    ;store the result
    pop         esi                             ;deallocate bHandled
    pop         esi                             ;deallocate lReturn
    pop         esi                             ;restore esi
    pop         ebp;                            ;restore ebp
    ret         24                              ;return and adjust the stack
    
_Callback_:
    push        dword lParam                    ;push byval lParam
    push        dword wParam                    ;push byval wParam
    push        dword nCode                     ;push byval nCode
    push        dword iHookType                 ;push byval HookType
    lea         eax, lReturn                    ;get addressof lReturn  
    push        eax                             ;push byref lReturn
    
    cmp         esi, iHook_Before               ;set flags
    jne         _Skip_bHandled_                 ;if making the after call, skip the bHandled argument
    
    lea         eax, bHandled                   ;get addressof bHandled
    push        eax                             ;push byref bHandled
    
_Skip_bHandled_:
    mov         eax, oOwner                     ;get iHook interface pointer
    push        eax                             ;push iHook interface pointer
    mov         eax, [eax]                      ;get iHook vTable pointer
    call        dword [eax+esi]                 ;make the call
    ret

_Ide_Break_:
    call        _NextHook_                      ;call the next hook
    jmp short   _Return_                        ;return

_Ide_Stop_:
    call        _NextHook_                      ;call the next hook
    call        _Unhook_                        ;unhook
    jmp short   _Return_                        ;return
    
_NextHook_:
    push        dword lParam                    ;push ByVal lParam
    push        dword wParam                    ;push ByVal wParam
    push        dword nCode                     ;push ByVal nCode
    push        dword hHook                     ;push ByVal hHook
    db          0e8h                            ;Far call opcode
    dd          _patch_CallNextHookEx_          ;Relative address of User32.CallNextHookEx - patched at runtime
    mov         lReturn, eax                    ;Preserve the return value
    ret

_Unhook_:
    push        dword hHook                     ;push byval hHook
    db          0e8h                            ;far call opcode
    dd          _patch_UnhookWindowsHookEx_     ;relative address to User32.UnhookWindowHookEx - patched at runtime
    ret