; AutoHotKey�ݒ�l
#HotkeyInterval, 1000
#MaxHotkeysPerInterval, 200
excel_hscroll_speed := 2

; �E�B���h�E�����
!q::Send, !{F4}

; Chrome���A�i�ށE�߂�p�ݒ�
vk1D & ,::Send, !{Left}
vk1D & .::Send, !{Right}

; �A���_�[�X�R�A
vkE2:: +vkE2

; �E�B���h�E�ړ�
vk1C & Tab::Send, #+{Right}
; �E�B���h�E����
vk1C & i::Send, #{Up}           ; �ő剻
vk1C & j::Send, #{Left}         ; ����
vk1C & l::Send, #{Right}        ; �E��
vk1C & k::Send, #{Down}         ; �ŏ���

; ���z�f�X�N�g�b�v�؂�ւ�
vk1D & r::Send, #^{Right}
vk1D & e::Send, #^{Left}

;���ϊ�+jkil = �㉺���E
;���ϊ�+shift+�㉺���E = shift+�㉺���E

;���ϊ�+j����
vk1D & j::
    if GetKeyState("shift", "P") && GetKeyState("ctrl", "P"){
        Send, +^{Left}
    }else if GetKeyState("shift", "P"){
        Send, +{Left}
    }else if GetKeyState("ctrl", "P"){
        Send, ^{Left}
    }else{
        Send, {Left}
    }
    return
;���ϊ�+i����
vk1D & i::
    if GetKeyState("shift", "P") && GetKeyState("ctrl", "P"){
        Send, +^{Up}
    }else if GetKeyState("shift", "P"){
        Send, +{Up}
    }else if GetKeyState("ctrl", "P"){
        Send, ^{Up}
    }else{
        Send, {Up}
    }
    return
;���ϊ�+k����
vk1D & k::
    if GetKeyState("shift", "P") && GetKeyState("ctrl", "P"){
        Send, +^{Down}
    }else if GetKeyState("shift", "P"){
        Send, +{Down}
    }else if GetKeyState("ctrl", "P"){
        Send, ^{Down}
    }else if GetKeyState("alt", "P"){
        Send, !{Down}
    }else{
        Send, {Down}
    }
    return
;���ϊ�+l���E
vk1D & l::
    if GetKeyState("shift", "P") && GetKeyState("ctrl", "P"){
        Send, +^{Right}
    }else if GetKeyState("shift", "P"){
        Send, +{Right}
    }else if GetKeyState("ctrl", "P"){
        Send, ^{Right}
    }else{
        Send, {Right}
    }
    return

; PageUp
vk1D & u::
    if GetKeyState("ctrl", "P"){
        Send,^{PgUp}
    }else{
        Send,{PgUp}
    }
    return

; PageDown
vk1D & o::
    if GetKeyState("ctrl", "P"){
        Send,^{PgDn}
    }else{
        Send,{PgDn}
    }
    return

; �t�@���N�V�����L�[�u������
vk1C & 1::
    if GetKeyState("ctrl", "P"){
        Send, ^{F1}
    }else{
    }
    return
vk1C & 2::Send, {F2}
vk1C & 3::Send, {}
vk1C & 4::Send, {F4}
vk1C & 5::Send, {F5}
vk1C & 7::Send, {F7}    ; �S�p�J�i
vk1C & 8::Send, {F8}    ; ���p�J�i
vk1C & 0::Send, {F10}   ; ���p�p��
vk1C & -::Send, {F11}
vk1C & ^::Send, {F12}

; ����pFunction�L�[
vk1D & f::Send, {F11}
vk1D & d::
    if GetKeyState("shift", "P"){
        Send, +{F12}
    }else{
        Send, {F12}
    }
    return

; �E�N���b�N
vk1D & 0::Send, +{F10}

; Excel ���X�N���[��
#IfWinActive, ahk_exe EXCEL.EXE
~LShift & WheelUp:: ; Scroll left.
SetScrollLockState, On
Loop, %excel_hscroll_speed%
{
    SendInput {Left}
}
SetScrollLockState, Off
return
~LShift & WheelDown:: ; Scroll right.
SetScrollLockState, On
Loop, %excel_hscroll_speed%
{
    SendInput {Right}
}
SetScrollLockState, Off
