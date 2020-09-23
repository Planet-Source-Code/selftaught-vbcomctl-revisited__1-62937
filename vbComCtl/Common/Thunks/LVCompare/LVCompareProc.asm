%define _patch_vTableOffset_        1Ch         ;vtable offset for callback function

%define _ObjPtr_                    [ebp+16]    ;iLVCompare interface pointer
%define _lParam2_                   [ebp+12]    ;lParam from second listitem to compare
%define _lParam1_                   [ebp+8]     ;lParam from first listitem to compare
%define _ReturnAddr_                [ebp+4]     ;Return address to sorting function
%define _ebp_Storage_               [ebp+0]     ;ebp stored here
%define _lReturn_                   [ebp-4]     ;function return value stored here

[bits 32]
    
    push        ebp                             ;store ebp
    mov         ebp, esp                        ;setup stack frame

    xor         eax, eax                        ;clear eax
    push        eax                             ;allocate lReturn
	
    lea         eax, _lReturn_                  ;get addressof lReturn
    push        eax                             ;push byref lReturn
    push        dword _lParam2_                 ;push byval lParam2
    push        dword _lParam1_                 ;push byval lParam1
    mov         eax, _ObjPtr_                   ;get ObjPtr
    push        eax                             ;push ObjPtr
    mov         eax, [eax]                      ;get vTable Ptr
    
    call        [eax+_patch_vTableOffset_]      ;call the function
    mov         eax, _lReturn_                  ;store the return value
    mov         esp, ebp                        ;restore stack pointer
    pop         ebp                             ;restore ebp
    ret         12