     1                                  %define _patch_vTableOffset_        1Ch         ;vtable offset for callback function
     2                                  
     3                                  %define _ObjPtr_                    [ebp+16]    ;iLVCompare interface pointer
     4                                  %define _lParam2_                   [ebp+12]    ;lParam from second listitem to compare
     5                                  %define _lParam1_                   [ebp+8]     ;lParam from first listitem to compare
     6                                  %define _ReturnAddr_                [ebp+4]     ;Return address to sorting function
     7                                  %define _ebp_Storage_               [ebp+0]     ;ebp stored here
     8                                  %define _lReturn_                   [ebp-4]     ;function return value stored here
     9                                  
    10                                  [bits 32]
    11                                      
    12 00000000 55                          push        ebp                             ;store ebp
    13 00000001 89E5                        mov         ebp, esp                        ;setup stack frame
    14                                  
    15 00000003 31C0                        xor         eax, eax                        ;clear eax
    16 00000005 50                          push        eax                             ;allocate lReturn
    17                                  	
    18 00000006 8D45FC                      lea         eax, _lReturn_                  ;get addressof lReturn
    19 00000009 50                          push        eax                             ;push byref lReturn
    20 0000000A FF750C                      push        dword _lParam2_                 ;push byval lParam2
    21 0000000D FF7508                      push        dword _lParam1_                 ;push byval lParam1
    22 00000010 8B4510                      mov         eax, _ObjPtr_                   ;get ObjPtr
    23 00000013 50                          push        eax                             ;push ObjPtr
    24 00000014 8B00                        mov         eax, [eax]                      ;get vTable Ptr
    25                                      
    26 00000016 FF501C                      call        [eax+_patch_vTableOffset_]      ;call the function
    27 00000019 8B45FC                      mov         eax, _lReturn_                  ;store the return value
    28 0000001C 89EC                        mov         esp, ebp                        ;restore stack pointer
    29 0000001E 5D                          pop         ebp                             ;restore ebp
    30 0000001F C20C00                      ret         12
