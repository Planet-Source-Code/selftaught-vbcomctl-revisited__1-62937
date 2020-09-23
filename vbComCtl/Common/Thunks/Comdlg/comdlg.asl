     1                                  %define _patch_EbMode_      01BCCAABh   ;relative address to vba6.EbMode
     2                                  %define _patch_dlg_type_    02BCCAABh   ;value that the caller uses to identify the type of dialog
     3                                  %define _patch_owner_       03BCCAABh   ;address of owner object
     4                                  
     5                                  %define lParam              [ebp+20]    ;wndproc lParam
     6                                  %define wParam              [ebp+16]    ;wndproc wParam
     7                                  %define iMsg                [ebp+12]    ;wndproc msg
     8                                  %define hDlg                [ebp+8]     ;handle to dialog window
     9                                  %define lRet                [ebp-4]     ;wndproc return value
    10                                  
    11                                  [bits 32]
    12                                  
    13 00000000 55                      push        ebp                         ;preserve ebp
    14 00000001 89E5                    mov         ebp,esp                     ;store stack pointer in ebp
    15 00000003 31C0                    xor         eax, eax                    ;eax is now 0
    16 00000005 50                      push        eax                         ;allocate lRet
    17                                  
    18 00000006 EB0D                    jmp short   _callback_                  ;patched with NOOP's if in the ide
    19                                  
    20 00000008 E8                      db          0E8h                        ;far call opcode
    21 00000009 ABCABC01                dd          _patch_EbMode_              ;relative address to vba6.EbMode - patched at runtime
    22 0000000D 3C02                    cmp         al, 2                       ;test for return value 2
    23 0000000F 7425                    je          _return_                    ;if return value is 2, ide is on a breakpoint
    24 00000011 85C0                    test        eax, eax                    ;test for return value 0
    25 00000013 7421                    je          _return_                    ;if return value is 0, ide is stopped
    26                                  
    27                                  _callback_:
    28 00000015 FF7514                  push        dword lParam                ;push byval lParam
    29 00000018 FF7510                  push        dword wParam                ;push byval lParam
    30 0000001B FF750C                  push        dword iMsg                  ;push byval iMsg
    31 0000001E FF7508                  push        dword hDlg                  ;push byval hDlg
    32                                  
    33 00000021 8D45FC                  lea         eax, lRet                   ;get address of lRet
    34 00000024 50                      push        eax                         ;push byref lRet
    35 00000025 B8ABCABC02              mov         eax, _patch_dlg_type_       ;store DialogType (patched at runtime)
    36 0000002A 50                      push        eax                         ;push byval DialogType
    37 0000002B B8ABCABC03              mov         eax, _patch_owner_          ;store objptr of the owner object (patched at runtime)
    38 00000030 50                      push        eax                         ;Push objptr of the owner object
    39 00000031 8B00                    mov         eax,  [eax]                 ;Get the address of the vTable
    40 00000033 FF501C                  call        dword [eax+1Ch]             ;Call iComDlgHook::Proc, vTable offset 1Ch
    41                                  
    42                                  _return_:
    43 00000036 8B45FC                  mov         eax, lRet                   ;store the function return value
    44 00000039 5D                      pop         ebp                         ;deallocate lRet
    45 0000003A 5D                      pop         ebp                         ;restore ebp
    46 0000003B C21000                  ret         16                          ;done!
