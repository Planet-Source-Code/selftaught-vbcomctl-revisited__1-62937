%define _patch_EbMode_      01BCCAABh   ;relative address to vba6.EbMode
%define _patch_dlg_type_    02BCCAABh   ;value that the caller uses to identify the type of dialog
%define _patch_owner_       03BCCAABh   ;address of owner object

%define lParam              [ebp+20]    ;wndproc lParam
%define wParam              [ebp+16]    ;wndproc wParam
%define iMsg                [ebp+12]    ;wndproc msg
%define hDlg                [ebp+8]     ;handle to dialog window
%define lRet                [ebp-4]     ;wndproc return value

[bits 32]

push        ebp                         ;preserve ebp
mov         ebp,esp                     ;store stack pointer in ebp
xor         eax, eax                    ;eax is now 0
push        eax                         ;allocate lRet

jmp short   _callback_                  ;patched with NOOP's if in the ide

db          0E8h                        ;far call opcode
dd          _patch_EbMode_              ;relative address to vba6.EbMode - patched at runtime
cmp         al, 2                       ;test for return value 2
je          _return_                    ;if return value is 2, ide is on a breakpoint
test        eax, eax                    ;test for return value 0
je          _return_                    ;if return value is 0, ide is stopped

_callback_:
push        dword lParam                ;push byval lParam
push        dword wParam                ;push byval lParam
push        dword iMsg                  ;push byval iMsg
push        dword hDlg                  ;push byval hDlg

lea         eax, lRet                   ;get address of lRet
push        eax                         ;push byref lRet
mov         eax, _patch_dlg_type_       ;store DialogType (patched at runtime)
push        eax                         ;push byval DialogType
mov         eax, _patch_owner_          ;store objptr of the owner object (patched at runtime)
push        eax                         ;Push objptr of the owner object
mov         eax,  [eax]                 ;Get the address of the vTable
call        dword [eax+1Ch]             ;Call iComDlgHook::Proc, vTable offset 1Ch

_return_:
mov         eax, lRet                   ;store the function return value
pop         ebp                         ;deallocate lRet
pop         ebp                         ;restore ebp
ret         16                          ;done!