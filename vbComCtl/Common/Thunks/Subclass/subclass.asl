     1                                  %define _patch_pPropString_                  01BCCAABh          ;pointer to null-terminated ANSI string for the window property
     2                                  %define _patch_GetProp_                      02BCCAABh          ;Relative address to User32.GetPropA
     3                                  %define _patch_CallWindowProcA_              03BCCAABh          ;Relative address to User32.CallWindowProcA
     4                                  %define _patch_SetWindowLong_                04BCCAABh          ;Relative address to User32.GetWindowLongA
     5                                  %define _patch_EbMode_                       05BCCAABh          ;Relative address to vba6.EbMode
     6                                  
     7                                  %define _lParam_                             [ebp+20]           ;windowproc lParam
     8                                  %define _wParam_                             [ebp+16]           ;windowproc wParam
     9                                  %define _uMsg_                               [ebp+12]           ;windowproc uMsg
    10                                  %define _hWnd_                               [ebp+8]            ;windowproc hWnd
    11                                  %define _RetAddr_                            [ebp+4]            ;return address
    12                                  %define _ebp_storage_                        [ebp+0]            ;ebp stored here
    13                                  %define _ecx_storage_                        [ebp-4]            ;ecx stored here
    14                                  %define _edi_storage_                        [ebp-8]            ;edi stored here
    15                                  %define _edx_storage_                        [ebp-12]           ;edx stored here
    16                                  %define _esi_storage_                        [ebp-16]           ;esi stored here
    17                                  %define _lReturn_                            [ebp-20]           ;local lReturn
    18                                  %define _bHandled_                           [ebp-24]           ;local bHandled
    19                                  %define _pSubClientNext_                     [ebp-28]           ;local pSubClientNext
    20                                  %define _pWndProcPrev_                       [ebp-32]           ;local pWndProcPrev
    21                                  
    22                                  %define GWL_WNDPROC                          -4                 ;passed to User32.SetWindowLongA to unsubclass
    23                                  
    24                                  %define _Subclass_pWndProcPrev_Offset_       0                  ;offset to the previous windowproc in the subclass data structure
    25                                  %define _Subclass_pSubClientNext_Offset_     4                  ;offset to the next subclass data structure
    26                                  %define _Subclass_pObject_Offset_            8                  ;offset to the iSubclass interface pointer
    27                                  %define _Subclass_pMsgTableA_Offset_         12                 ;offset to the before message table
    28                                  %define _Subclass_pMsgTableB_Offset_         16                 ;offset to the after message table
    29                                  
    30                                  %define _ALL_MESSAGES_                       -1                 ;special value for the message table pointer indicating that all messages call back
    31                                  
    32                                  [bits 32]
    33 00000000 55                          push        ebp                                             ;store ebp
    34 00000001 89E5                        mov         ebp, esp                                        ;setup stack frame
    35 00000003 51                          push        ecx                                             ;store ecx
    36 00000004 57                          push        edi                                             ;store edi
    37 00000005 52                          push        edx                                             ;store edx
    38 00000006 56                          push        esi                                             ;store esi
    39 00000007 31F6                        xor         esi, esi                                        ;esi is now 0
    40 00000009 56                          push        esi                                             ;allocate/initialize lReturn
    41 0000000A 56                          push        esi                                             ;allocate/initialize bHandled
    42 0000000B 56                          push        esi                                             ;allocate/initialize pSubClientNext
    43 0000000C 56                          push        esi                                             ;allocate/initialize pWndProcPrev
    44                                  
    45 0000000D E843000000                  call        _GetProp_                                       ;try to get the pointer to subclass data
    46 00000012 7430                        jz          _Return_                                        ;if this failed, we're SOL, return
    47                                      
    48 00000014 8B10                        mov         edx, [eax+_Subclass_pWndProcPrev_Offset_]       ;get the previous wndproc address
    49 00000016 8955E0                      mov         _pWndProcPrev_, edx                             ;store the previous wndproc address
    50 00000019 89C2                        mov         edx, eax                                        ;get the pointer to the first subclass client data
    51                                      
    52 0000001B BE10000000                  mov         esi, _Subclass_pMsgTableB_Offset_               ;set the message table offset to use in this callback
    53 00000020 E89F000000                  call        _Callback_                                      ;make the before callbacks
    54                                      
    55 00000025 8B45E8                      mov         eax, _bHandled_                                 ;store bHandled
    56 00000028 85C0                        test        eax, eax                                        ;test bHandled
    57 0000002A 7518                        jnz         _Return_                                        ;if bHandled == True then return
    58                                      
    59 0000002C E85A000000                  call        _Original_                                      ;call the original window procedure
    60                                      
    61 00000031 E81F000000                  call        _GetProp_                                       ;try to get the pointer to subclass data
    62 00000036 740C                        jz          _Return_                                        ;if this failed, subclass was removed during a previous callback
    63                                      
    64 00000038 89C2                        mov         edx, eax                                        ;get the pointer to the first subclass client data
    65 0000003A BE0C000000                  mov         esi, _Subclass_pMsgTableA_Offset_               ;set the message table offset to use in this callback
    66 0000003F E880000000                  call        _Callback_                                      ;make the after callbacks
    67                                      
    68                                  _Return_:
    69 00000044 8B45EC                      mov         eax, _lReturn_                                  ;store the return value
    70 00000047 81C410000000                add         esp, 16                                         ;deallocate pWndProcPrev, pSubClientNext, bHandled and lReturn
    71 0000004D 5E                          pop         esi                                             ;restore esi
    72 0000004E 5A                          pop         edx                                             ;restore edx
    73 0000004F 5F                          pop         edi                                             ;restore edi
    74 00000050 59                          pop         ecx                                             ;restore ecx
    75 00000051 5D                          pop         ebp                                             ;restore ebp
    76 00000052 C21000                      ret         16                                              ;return and deallocate lParam, wParam, uMsg and hWnd
    77                                      
    78                                  _GetProp_:
    79 00000055 68ABCABC01                  push        dword _patch_pPropString_                       ;push the pointer to the property string
    80 0000005A FF7508                      push        dword _hWnd_                                    ;push the hwnd
    81 0000005D E8                          db          0e8h                                            ;far call opcode
    82 0000005E ABCABC02                    dd          _patch_GetProp_                                 ;relative address to User32.GetPropA - Patched at runtime
    83 00000062 85C0                        test        eax, eax                                        ;test the return value
    84 00000064 C3                          ret                                                         ;return to caller
    85                                      
    86                                  _Find_Message_:
    87 00000065 85FF                        test        edi, edi                                        ;test the message table pointer
    88 00000067 741D                        jz          _Find_Message_Not_Found_                        ;if 0, no messages call back
    89 00000069 81FFFFFFFFFF                cmp         edi, _ALL_MESSAGES_                             ;test the message table pointer for -1
    90 0000006F 740F                        je          _Find_Message_Found_                            ;if -1, all messages call back
    91                                      
    92 00000071 8B0F                        mov         ecx, [edi]                                      ;store the message table count (the first DWORD)
    93 00000073 81C704000000                add         edi, 4                                          ;increment the message table pointer past the count
    94 00000079 8B450C                      mov         eax, _uMsg_                                     ;store the message we will look for
    95 0000007C F2AF                        repne       scasd                                           ;scan the table
    96 0000007E 7506                        jne         _Find_Message_Not_Found_                        ;if message was not found, jump
    97                                      
    98                                  _Find_Message_Found_:
    99 00000080 31C0                        xor         eax, eax                                        ;clear eax
   100 00000082 48                          dec         eax                                             ;eax is now -1
   101 00000083 85C0                        test        eax, eax                                        ;set the flags
   102 00000085 C3                          ret                                                         ;return to caller
   103                                      
   104                                  _Find_Message_Not_Found_:
   105 00000086 31C0                        xor         eax, eax                                        ;clear eax
   106 00000088 85C0                        test        eax, eax                                        ;set the flags
   107 0000008A C3                          ret                                                         ;return to caller
   108                                      
   109                                  _Original_:
   110 0000008B FF7514                      push        dword _lParam_                                  ;push byval lParam
   111 0000008E FF7510                      push        dword _wParam_                                  ;push byval wParam
   112 00000091 FF750C                      push        dword _uMsg_                                    ;push byval uMsg
   113 00000094 FF7508                      push        dword _hWnd_                                    ;push byval hWnd
   114 00000097 8B45E0                      mov         eax,  _pWndProcPrev_                            ;get the original wndproc address
   115 0000009A 50                          push        eax                                             ;push the original wndproc address
   116 0000009B E8                          db          0e8h                                            ;far call opcode
   117 0000009C ABCABC03                    dd          _patch_CallWindowProcA_                         ;relative address to User32.CallWindowProcA - patched at runtime
   118 000000A0 8945EC                      mov         _lReturn_, eax                                  ;store the return value
   119 000000A3 C3                          ret
   120                                      
   121                                  _Ide_Stop_:
   122 000000A4 FF75E0                      push        dword _pWndProcPrev_                            ;push the original wndproc address
   123 000000A7 68FCFFFFFF                  push        dword GWL_WNDPROC                               ;push GWL_WNDPROC
   124 000000AC FF7508                      push        dword _hWnd_                                    ;push the hwnd
   125 000000AF E8                          db          0e8h                                            ;far call opcode
   126 000000B0 ABCABC04                    dd          _patch_SetWindowLong_                           ;relative address to User32.SetWindowLongA - patched at runtime
   127                                      
   128                                  _Ide_Break_:
   129 000000B4 58                          pop         eax                                             ;pop the local return address
   130 000000B5 81FE10000000                cmp         esi, _Subclass_pMsgTableB_Offset_               ;find out if the original procedure has already been called
   131 000000BB 7587                        jne         _Return_                                        ;if it has, return
   132 000000BD E8C9FFFFFF                  call        _Original_                                      ;if it has not, call it
   133 000000C2 EB80                        jmp short   _Return_                                        ;then return
   134                                      
   135                                  _Callback_:
   136 000000C4 EB0D                        jmp short   _No_Ide_Check_                                  ;skip the ide check - patched with NOOPs if in the ide
   137                                      
   138 000000C6 E8                          db          0E8h                                            ;far call opcode
   139 000000C7 ABCABC05                    dd          _patch_EbMode_                                  ;relative address to vba6.EbMode - patched at runtime
   140 000000CB 3C02                        cmp         al, 2                                           ;test for return value 2
   141 000000CD 74E5                        je          _Ide_Break_                                     ;if return value is 2, call the original procedure
   142 000000CF 85C0                        test        eax, eax                                        ;test for return value 0
   143 000000D1 74D1                        je          _Ide_Stop_                                      ;if return value is 0, call the original procedure and unsubclass
   144                                      
   145                                  _No_Ide_Check_:
   146 000000D3 8B7A04                      mov         edi, [edx+_Subclass_pSubClientNext_Offset_]     ;get the pointer to the next subclass client data
   147 000000D6 897DE4                      mov         _pSubClientNext_, edi                           ;store the pointer to the next subclass client data
   148 000000D9 8B3C32                      mov         edi, [edx+esi]                                  ;get the pointer to the current subclass client message table
   149 000000DC E884FFFFFF                  call        _Find_Message_                                  ;try to find the message in the table
   150 000000E1 7437                        jz          _Callback_Skip_                                 ;if the message was not found, skip the callback
   151                                      
   152 000000E3 FF7514                      push        dword _lParam_                                  ;push byval lParam
   153 000000E6 FF7510                      push        dword _wParam_                                  ;push byval wParam
   154 000000E9 FF750C                      push        dword _uMsg_                                    ;push byval uMsg
   155 000000EC FF7508                      push        dword _hWnd_                                    ;push byval hWnd
   156 000000EF 8D45EC                      lea         eax, _lReturn_                                  ;get addressof lReturn
   157 000000F2 50                          push        eax                                             ;push byref lReturn
   158                                      
   159 000000F3 81FE0C000000                cmp         esi, _Subclass_pMsgTableA_Offset_               ;are we making an after or before callback?
   160 000000F9 7416                        je          _Callback_After_                                ;if after, jump
   161                                      
   162                                  _Callback_Before_:
   163 000000FB 8D45E8                      lea         eax, _bHandled_                                 ;get addressof bHandled
   164 000000FE 50                          push        eax                                             ;push byref bHandled
   165                                      
   166 000000FF 8B4208                      mov         eax, [edx+_Subclass_pObject_Offset_]            ;get the pointer to the iSubclass interface
   167 00000102 50                          push        eax                                             ;push the pointer to the iSubclass interface
   168 00000103 8B00                        mov         eax, [eax]                                      ;get a pointer to the vtable
   169 00000105 FF501C                      call        dword [eax+1Ch]                                 ;make the before call, vtable offset 1Ch
   170                                      
   171 00000108 8B45E8                      mov         eax, _bHandled_                                 ;store bHandled
   172 0000010B 85C0                        test        eax, eax                                        ;test bHandled
   173 0000010D 7512                        jnz         _Callback_Return_                               ;if bHandled == True, return
   174 0000010F EB09                        jmp short   _Callback_Skip_                                 ;skip the after callback for now
   175                                      
   176                                  _Callback_After_:
   177 00000111 8B4208                      mov         eax, [edx+_Subclass_pObject_Offset_]            ;get the pointer to the iSubclass interface
   178 00000114 50                          push        eax                                             ;push the pointer to the iSubclass interface
   179 00000115 8B00                        mov         eax, [eax]                                      ;get a pointer to the vtable
   180 00000117 FF5020                      call        dword [eax+20h]                                 ;make the call, vtable offset 20h
   181                                      
   182                                  _Callback_Skip_:
   183 0000011A 8B55E4                      mov         edx, _pSubClientNext_                           ;store the pointer to the next subclass client data
   184 0000011D 85D2                        test        edx, edx                                        ;set the flags
   185 0000011F 75A3                        jnz         _Callback_                                      ;if there is another subclass client, start over
   186                                      
   187                                  _Callback_Return_:
   188 00000121 C3                          ret                                                         ;return to caller
