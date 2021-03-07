
; Initialize ensuring Google Drive shell context are disabled.
setRegistry(false)

; Also clean up when exiting.
OnExit("exitFunc")

; Don't bother executing any hooks below at all unless actually in Explorer.
#IfWinActive, ahk_exe HmFilerClassic.exe
{
    ; Register shift key hooks.
    ~LShift::
        setRegistry(true)
    return

    ~LShift Up::
        setRegistry(false)
    return


    setRegistry(enable) {
        prefix:="-"
        if (enable) {
            prefix:=""
        }
        
        ; The extra comma in parameter list is an empty value for the "(Default)" value for the key (as you'd see it in regedit).
        ; See: https://autohotkey.com/docs/commands/RegWrite.htm
        
        ; Directories
        RegWrite, REG_SZ, HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers\GDContextMenu, , %prefix%{BB02B294-8425-42E5-983F-41A1FA970CD6}
        RegWrite, REG_SZ, HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers\DriveFS 28 or later, , %prefix%{EE15C2BD-CECB-49F8-A113-CA1BFC528F5B}
        
        ; Files
        RegWrite, REG_SZ, HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\GDContextMenu, , %prefix%{BB02B294-8425-42E5-983F-41A1FA970CD6}
        RegWrite, REG_SZ, HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\DriveFS 28 or later, , %prefix%{EE15C2BD-CECB-49F8-A113-CA1BFC528F5B}
    }
}
#IfWinActive

; Reverts back to enabling drive on exit.
exitFunc() {
    setRegistry(true)
    ExitApp
}


; AutoHotKey設定値
#HotkeyInterval, 1000
#MaxHotkeysPerInterval, 200
excel_hscroll_speed := 2

; ウィンドウを閉じる
!q::Send, !{F4}

; Chrome等、進む・戻る用設定
vk1D & ,::Send, !{Left}
vk1D & .::Send, !{Right}

; アンダースコア
vkE2:: +vkE2

; ウィンドウ移動
vk1C & Tab::Send, #+{Right}
; ウィンドウ操作
vk1C & i::Send, #{Up}           ; 最大化
vk1C & j::Send, #{Left}         ; 左寄せ
vk1C & l::Send, #{Right}        ; 右寄せ
vk1C & k::Send, #{Down}         ; 最小化

; 仮想デスクトップ切り替え
vk1D & r::Send, #^{Right}
vk1D & e::Send, #^{Left}

;無変換+jkil = 上下左右
;無変換+shift+上下左右 = shift+上下左右

;無変換+j→左
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
;無変換+i→上
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
;無変換+k→下
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
;無変換+l→右
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

; ファンクションキー置き換え
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
vk1C & 7::Send, {F7}    ; 全角カナ
vk1C & 8::Send, {F8}    ; 半角カナ
vk1C & 0::Send, {F10}   ; 半角英数
vk1C & -::Send, {F11}
vk1C & ^::Send, {F12}

; 左手用Functionキー
vk1D & f::Send, {F11}
vk1D & d::
    if GetKeyState("shift", "P"){
        Send, +{F12}
    }else{
        Send, {F12}
    }
    return

; 右クリック
vk1D & 0::Send, +{F10}