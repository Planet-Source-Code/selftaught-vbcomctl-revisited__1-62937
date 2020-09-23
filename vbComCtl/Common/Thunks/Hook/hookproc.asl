     1                                  %define _patch_EbMode_              01BCCAABh   ;relative address to vba6.EbMode
     2                                  %define _patch_CallNextHookEx_      02BCCAABh   ;relative address to User32.CallNextHookEx
     3                                  %define _patch_UnhookWindowsHookEx_ 03BCCAABh   ;relative address to User32.UnHookWindowsHookEx
     4                                  
     5                                  %define lParam                      [ebp+40]    ;hookproc lparam
     6                                  %define wParam                      [ebp+36]    ;hookproc wparam
     7                                  %define nCode                       [ebp+32]    ;hookproc nCode
     8                                  %define iHookType                   [ebp+28]    ;hook type ie WH_MOUSE, etc.
     9                                  %define hHook                       [ebp+24]    ;hook handle
    10                                  %define oOwner                      [ebp+20]    ;iHook interface
    11                                  %define pAddrReturn                 [ebp+16]    ;return address
    12                                  %define ebpStorage                  [ebp+12]    ;ebp stored here
    13                                  %define esiStorage                  [ebp+8]     ;esi stored here
    14                                  %define lReturn                     [ebp+4]     ;local lReturn
    15                                  %define bHandled                    [ebp+0]     ;local bHandled
    16                                  
    17                                  %define iHook_Before                1Ch         ;vtable offset to iHook::Before
    18                                  %define iHook_After                 20h         ;vtable offset to iHook::After
    19                                  
    20                                  [bits 32]
    21                                      
    22 00000000 55                          push        ebp                             ;store ebp
    23 00000001 56                          push        esi                             ;store esi
    24 00000002 31ED                        xor         ebp, ebp                        ;clear ebp
    25 00000004 55                          push        ebp                             ;allocate lReturn
    26 00000005 55                          push        ebp                             ;allocate bHandled
    27 00000006 89E5                        mov         ebp, esp                        ;setup stack frame
    28                                  
    29 00000008 EB0D                        jmp short   _noide_                         ;patched with NOOPs if in the ide
    30                                      
    31 0000000A E8                          db          0E8h                            ;Far call
    32 0000000B ABCABC01                    dd          _patch_EbMode_                  ;relative address to vba6.EbMode - patched at runtime
    33 0000000F 3C02                        cmp         al, 2                           ;test return value for 2
    34 00000011 7454                        je          _Ide_Break_                     ;If 2, IDE is on a breakpoint, just call the next hook
    35 00000013 85C0                        test        eax, eax                        ;test return value for 0
    36 00000015 7457                        je          _Ide_Stop_                      ;If 0, IDE has stopped, unhook
    37                                  
    38                                  _noide_:    
    39 00000017 BE1C000000                  mov         esi, iHook_Before               ;first call iHookProc_Before
    40 0000001C E820000000                  call        _Callback_                      ;make the before call
    41                                      
    42 00000021 8B7500                      mov         esi, bHandled                   ;store bHandled
    43 00000024 85F6                        test        esi, esi                        ;set flags
    44 00000026 750F                        jnz         _Return_                        ;if bHandled == True, return
    45                                      
    46 00000028 E84D000000                  call        _NextHook_                      ;call the next hook
    47                                      
    48 0000002D BE20000000                  mov         esi, iHook_After                ;next call iHookProc_After
    49 00000032 E80A000000                  call        _Callback_                      ;make the after call
    50                                      
    51                                  _Return_:
    52 00000037 8B4504                      mov         eax, lReturn                    ;store the result
    53 0000003A 5E                          pop         esi                             ;deallocate bHandled
    54 0000003B 5E                          pop         esi                             ;deallocate lReturn
    55 0000003C 5E                          pop         esi                             ;restore esi
    56 0000003D 5D                          pop         ebp;                            ;restore ebp
    57 0000003E C21800                      ret         24                              ;return and adjust the stack
    58                                      
    59                                  _Callback_:
    60 00000041 FF7528                      push        dword lParam                    ;push byval lParam
    61 00000044 FF7524                      push        dword wParam                    ;push byval wParam
    62 00000047 FF7520                      push        dword nCode                     ;push byval nCode
    63 0000004A FF751C                      push        dword iHookType                 ;push byval HookType
    64 0000004D 8D4504                      lea         eax, lReturn                    ;get addressof lReturn  
    65 00000050 50                          push        eax                             ;push byref lReturn
    66                                      
    67 00000051 81FE1C000000                cmp         esi, iHook_Before               ;set flags
    68 00000057 7504                        jne         _Skip_bHandled_                 ;if making the after call, skip the bHandled argument
    69                                      
    70 00000059 8D4500                      lea         eax, bHandled                   ;get addressof bHandled
    71 0000005C 50                          push        eax                             ;push byref bHandled
    72                                      
    73                                  _Skip_bHandled_:
    74 0000005D 8B4514                      mov         eax, oOwner                     ;get iHook interface pointer
    75 00000060 50                          push        eax                             ;push iHook interface pointer
    76 00000061 8B00                        mov         eax, [eax]                      ;get iHook vTable pointer
    77 00000063 FF1430                      call        dword [eax+esi]                 ;make the call
    78 00000066 C3                          ret
    79                                  
    80                                  _Ide_Break_:
    81 00000067 E80E000000                  call        _NextHook_                      ;call the next hook
    82 0000006C EBC9                        jmp short   _Return_                        ;return
    83                                  
    84                                  _Ide_Stop_:
    85 0000006E E807000000                  call        _NextHook_                      ;call the next hook
    86 00000073 E817000000                  call        _Unhook_                        ;unhook
    87 00000078 EBBD                        jmp short   _Return_                        ;return
    88                                      
    89                                  _NextHook_:
    90 0000007A FF7528                      push        dword lParam                    ;push ByVal lParam
    91 0000007D FF7524                      push        dword wParam                    ;push ByVal wParam
    92 00000080 FF7520                      push        dword nCode                     ;push ByVal nCode
    93 00000083 FF7518                      push        dword hHook                     ;push ByVal hHook
    94 00000086 E8                          db          0e8h                            ;Far call opcode
    95 00000087 ABCABC02                    dd          _patch_CallNextHookEx_          ;Relative address of User32.CallNextHookEx - patched at runtime
    96 0000008B 894504                      mov         lReturn, eax                    ;Preserve the return value
    97 0000008E C3                          ret
    98                                  
    99                                  _Unhook_:
   100 0000008F FF7518                      push        dword hHook                     ;push byval hHook
   101 00000092 E8                          db          0e8h                            ;far call opcode
   102 00000093 ABCABC03                    dd          _patch_UnhookWindowsHookEx_     ;relative address to User32.UnhookWindowHookEx - patched at runtime
   103 00000097 C3                          ret
