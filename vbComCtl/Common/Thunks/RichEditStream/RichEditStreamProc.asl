     1                                  %define _lpBytesProcessed_           [ebp+20]           ;bytes processed
     2                                  %define _BytesRequested_             [ebp+16]           ;byres requested
     3                                  %define _lPtrPbBuff_                 [ebp+12]           ;in/out buffer
     4                                  %define _lpObject_                   [ebp+8]            ;callback objptr
     5                                  %define _ReturnAddress_              [ebp+4]            ;return address
     6                                  %define _ebp_storage_                [ebp+0]            ;ebp stored here
     7                                  %define _lReturn_                    [ebp-4]            ;return value stored here
     8                                  
     9                                  [bits 32]
    10 00000000 55                          push        ebp                                     ;store ebp
    11 00000001 89E5                        mov         ebp, esp                                ;setup stack frame
    12                                      
    13 00000003 31C0                        xor         eax, eax                                ;clear eax
    14 00000005 50                          push        eax                                     ;initialize lReturn
    15 00000006 8D45FC                      lea         eax, _lReturn_                          ;get addressof lReturn
    16 00000009 50                          push        eax                                     ;push byref lReturn
    17 0000000A FF7514                      push        dword _lpBytesProcessed_                ;push byref BytesProcessed
    18 0000000D FF7510                      push        dword _BytesRequested_                  ;push byval BytesRequested
    19 00000010 FF750C                      push        dword _lPtrPbBuff_                      ;push byval PtrPbBuff
    20 00000013 8B4508                      mov         eax, _lpObject_                         ;get the callback objptr
    21 00000016 50                          push        eax                                     ;push byval objptr
    22 00000017 8B00                        mov         eax, [eax]                              ;get pointer to vtable
    23 00000019 FF501C                      call        [eax+1Ch]                               ;call RichEditStream::Proc
    24                                      
    25 0000001C 8B45FC                      mov         eax, _lReturn_                          ;return value given
    26 0000001F 89EC                        mov         esp, ebp                                ;restore stack pointer
    27 00000021 5D                          pop         ebp                                     ;restore ebp
    28 00000022 C21000                      ret         16                                      ;return
