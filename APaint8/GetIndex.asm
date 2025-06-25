; GetIndex.asm
; res=CallWindowProc(ptrMC, LongDerived, ptStanPal, 3&, 4&)
;                             8          12         16  20
%define LongDerived    [ebp+8]
%define ptStanPal      [ebp+12]
%define MinD           [ebp-4]
%define LongVal        [ebp-8]
%define Index          [ebp-12]

[bits 32]
    push ebp
    mov ebp,esp
    sub esp,12
    push edi
    push esi
    push ebx
    
    xor eax,eax
    mov Index,eax
    mov eax,LongDerived
    movd mm3,eax
    mov eax,1000
    mov MinD,eax
    mov edi,ptStanPal
    mov ecx,255
Fork:
   
   movq mm0,mm3     ; mm0 =RGBA LongDerived
   movd mm1,[edi]   ; mm1 =RGBA Standard
   movq mm2,mm0     ; mm2 =RGBA LongDerived
   psubusb mm0,mm1 ; eq mm0-mm1  Derived-Standard
   psubusb mm1,mm2 ; eq mm1-mm0  Standard-Derived
   por mm0,mm1     ; ABS(mm0-mm1) mm0 =| | | |ABGR|

   pxor mm1,mm1
   punpcklbw mm0,mm1  ; mm0 =|A|B|G|R|
   
   movq mm1,mm0
   movq mm2,mm0
   psrlq mm1,16       ; mm1 =| |A|B|G|
   psrlq mm2,32       ; mm2 =| | |A|B|
   paddsw mm0,mm1
   paddsw mm0,mm2     ; mm0 =|A|B+A|G+B+A|R+G+B|
   movd eax,mm0
   and eax,00000FFFFh ; eax =R+G+B =Sum[ABS(Diffs)]

   cmp eax,MinD    
   jg nextk
      mov MinD,eax
      mov Index,ecx
nextk:
   mov eax,4
   add edi,eax
   dec ecx
   jnz Fork

   mov eax,255
   sub eax,Index
 
GETOUT:
    emms
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16
;#########################################
