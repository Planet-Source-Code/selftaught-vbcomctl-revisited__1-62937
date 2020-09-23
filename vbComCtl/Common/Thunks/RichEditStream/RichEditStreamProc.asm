%define _lpBytesProcessed_           [ebp+20]           ;bytes processed
%define _BytesRequested_             [ebp+16]           ;byres requested
%define _lPtrPbBuff_                 [ebp+12]           ;in/out buffer
%define _lpObject_                   [ebp+8]            ;callback objptr
%define _ReturnAddress_              [ebp+4]            ;return address
%define _ebp_storage_                [ebp+0]            ;ebp stored here
%define _lReturn_                    [ebp-4]            ;return value stored here

[bits 32]
    push        ebp                                     ;store ebp
    mov         ebp, esp                                ;setup stack frame
    
    xor         eax, eax                                ;clear eax
    push        eax                                     ;initialize lReturn
    lea         eax, _lReturn_                          ;get addressof lReturn
    push        eax                                     ;push byref lReturn
    push        dword _lpBytesProcessed_                ;push byref BytesProcessed
    push        dword _BytesRequested_                  ;push byval BytesRequested
    push        dword _lPtrPbBuff_                      ;push byval PtrPbBuff
    mov         eax, _lpObject_                         ;get the callback objptr
    push        eax                                     ;push byval objptr
    mov         eax, [eax]                              ;get pointer to vtable
    call        [eax+1Ch]                               ;call RichEditStream::Proc
    
    mov         eax, _lReturn_                          ;return value given
    mov         esp, ebp                                ;restore stack pointer
    pop         ebp                                     ;restore ebp
    ret         16                                      ;return