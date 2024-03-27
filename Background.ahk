; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;; 0 Settings
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
#Requires AutoHotkey <2 64-bit
; msgbox, % A_OSVersion


VarSetCapacity(APPBARDATA, A_PtrSize=4 ? 36:48)

SendMode, Input

; SetKeyDelay, -1
; SetBatchLines -1

Process, Priority, , High

#MaxHotkeysPerInterval,120

SetTitleMatchMode, 2 ; 2:contain anywhere inside 1:startwith 3:exact

#InputLevel, 2 ; prevent send trigger other hotkey

; maxthread??? for toggle script to work
#MaxThreadsPerHotkey 4

#InstallKeybdHook
#InstallMouseHook

SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
; TODO: include game.ahk

SysGet RemoteControl, 4096 ; SM_REMOTESESSION ; 8193 SM_REMOTECONTROL: This system metric is used in a Terminal Services environment. Its value is nonzero if the current session is remotely controlled; zero otherwise.

; playtime tracker
SetTimer, PlaytimeTracker, 10000 ; 10s

SetTimer, FlaskChecker, 7000

; proxy checker
; SetTimer,ProxyChecker,15000

; hotkey,~WheelLeft,label1
; hotkey,~WheelRight,label2

HideMagifier:="T"
bak:="C:\"
X2FuncSwitch:="1"

hotkey,WheelUp,off
hotkey,WheelDown,off
; hotkey,RButton,off
hotkey,LButton,off

if GetKeyState("CapsLock", "T")
    SetCapsLockState, Off
; SetCapsLockState, AlwaysOff


if (A_Args[1]="nocapslock") ; if rdp still wtf, detect mstsc.exe and just relaunch ahk in client
{
    Hotkey, If, not WinActive("- 远程桌面连接 ahk_exe mstsc.exe")
    ; hotkey,*~capslock,cnm1
    ; hotkey,*~capslock up,cnm2
    Hotkey,CapsLock & w,toggle
    Hotkey,CapsLock & a,toggle
    Hotkey,CapsLock & s,toggle
    Hotkey,CapsLock & d,toggle
    Hotkey,CapsLock & q,toggle
    Hotkey,CapsLock & e,toggle
    Hotkey,CapsLock & r,toggle
    Hotkey,CapsLock & f,toggle
    Hotkey,CapsLock & z,toggle
    Hotkey,CapsLock & x,toggle
    ; Hotkey,CapsLock & c,toggle
    ; Hotkey,CapsLock & v,toggle
    Hotkey,CapsLock & 1,toggle
    Hotkey,CapsLock & 2,toggle
    Hotkey,CapsLock & 3,toggle
    Hotkey,CapsLock & 4,toggle
    Hotkey,CapsLock & 5,toggle
    Hotkey,CapsLock & F1,toggle
    Hotkey,CapsLock & F2,toggle
    Hotkey,CapsLock & F3,toggle
    Hotkey,CapsLock & Space,toggle
    Hotkey,CapsLock & ``,toggle
    Hotkey,CapsLock & Tab,toggle
    Hotkey,CapsLock & Alt,toggle
    Hotkey,CapsLock & Shift,toggle
    Hotkey, If
}
Else
{
    SetCapsLockState, AlwaysOff
    Hotkey, If, not WinActive("- 远程桌面连接 ahk_exe mstsc.exe") ; just for consistency?
    hotkey,*capslock,cnm1
    hotkey,*capslock up,cnm2
    Hotkey, If
}

if (A_Args[2]="wtf")
    SetTimer, wtf, -10
if (A_Args[2]="wtf2")
    SetTimer, wtf2, -10

; Hotkey, If, not WinActive("- 远程桌面连接 ahk_exe mstsc.exe") ; just for consistency?
; hotkey,RButton,off
; hotkey,RButton Up,off
; Hotkey, If


toggleAcrobatShift:=0
toggleBlockInput:=0
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;;   0.1 Functions
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------

; ReloadAdmin()
; {
;     full_command_line := DllCall("GetCommandLine", "str")
;     try
;     {
;         if A_IsCompiled
;             Run *RunAs "%A_ScriptFullPath%" /restart
;         else
;             Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
;     }
;     ; ExitApp
;     return
; }
; Menu,Tray,Insert,1&,Reload with Admin,ReloadAdmin
; Menu,Tray,NoStandard
; Menu,Tray,Standard

MoveCursor()
{
    Sleep,50
    CoordMode, Mouse, Screen
    MouseGetPos, MouseX, MouseY
    SysGet, Mon1, MonitorWorkArea, 1
    SysGet, Mon2, MonitorWorkArea, 2

    ; msgbox,%MouseX%, %Mon1Left% %Mon1Right% %Mon1Top% %Mon1Bottom%
    ; why cannot handle negative correctly???
    if MouseX between %Mon1Left% and %Mon1Right% and MouseY between Mon1Top and Mon1Bottom
    {
        x := (Mon2Left + Mon2Right) / 2
        y := (Mon2Top + Mon2Bottom) / 2
    }
    else
    {
        x := (Mon1Left + Mon1Right) / 2
        y := (Mon1Top + Mon1Bottom) / 2
    }
    DllCall("SetCursorPos", int, x, int, y)
    return
}

SendViaClip(Var) {
    ClipSaved := Clipboard
    Clipboard := Var
    Send,^v
    sleep,50 ; enough?!
    Clipboard := ClipSaved
    return
}

isJLabActive() {
    return WinActive(" - JupyterLab 和另外 ahk_exe msedge.exe") or WinActive("ahk_exe JupyterLab.exe") or WinActive("JupyterLab and ahk_exe msedge.exe") or 
        WinActive("JupyterLab - Google Chrome ahk_exe chrome.exe") or
        WinActive("RStudio Server ahk_exe msedge.exe")
}


; LastKeyEvents(n)
; {
;    Global Keys
;    KeyHistory
;    ControlGetText Keys, Edit1, ahk_class AutoHotkey
;    Send !{F4}
;    StringGetPos pos, Keys, ----------`r`n
;    StringTrimLeft Keys, Keys, pos+12
;    StringTrimRight Keys, Keys, 23
;    StringGetPos pos, Keys, `n, R%n%
;    StringTrimLeft Keys, Keys, pos+1
;    StringSplit Line, Keys, `n
;    return
; }
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;; 1 (To Sort)
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
#If WinActive("ahk_exe ONENOTE.exe") ; or WinActive("ahk_exe ApplicationFrameHost.exe") ; uwp, onenote etc
    Capslock & w::ControlSend, OneNote::DocumentCanvas1,{up},ahk_exe ONENOTE.EXE
    Capslock & s::ControlSend, OneNote::DocumentCanvas1,{down},ahk_exe ONENOTE.EXE
    ; Capslock & w::Sendplay, {Up} ; fuck!
    ; F1::Send, {Up} ; fuck!
#If 


; #If not WinActive("- 远程桌面连接 ahk_exe mstsc.exe")

^!F11:: ; emm
!F11::
^F11::
    WinGet, TempWindowID, ID, A
    If (WindowID != TempWindowID)
    {
        WindowID:=TempWindowID
        WindowState:=0
    }
    If (WindowState != 1)
    {
        WinGetPos, WinPosX, WinPosY, WindowWidth, WindowHeight, ahk_id %WindowID%
        WinSet, Style, -0xC40000, ahk_id %WindowID%
        ; WinMove, ahk_id %WindowID%, , 0, 0, A_ScreenWidth, A_ScreenHeight
        WinMove, ahk_id %WindowID%, , 0, 100, A_ScreenWidth, A_ScreenWidth/1920*1080
        ; MouseGetPos, , , hwnd
        ; Gui Cursor:+Owner%hwnd%
        ; DllCall("ShowCursor", Int,0)
        ; HideCursor("Off")
        ; MouseMove, 1920, 540 ; ke yi
    }
    Else
    {
        WinSet, Style, +0xC40000, ahk_id %WindowID%
        WinMove, ahk_id %WindowID%, , WinPosX, WinPosY, WindowWidth, WindowHeight
        ; DllCall("ShowCursor", Int,1)
        ; HideCursor("On")
    }
    WindowState:=!WindowState
    return

; ^!+#s::suspend ; shit the ^!+# triggers office

; CapsLock::goto,cnm
; cnm:
;     Send,{Enter}
;     return

; CapsLock::Enter
; *~CapsLock::goto, cnm1
; *~CapsLock up::goto, cnm2


; --------------------------------------------------------------------------------------------------------------------------------
;;   1.0 Capslock
; --------------------------------------------------------------------------------------------------------------------------------
#If not WinActive("- 远程桌面连接 ahk_exe mstsc.exe")
        
    ^#CapsLock:: ; fuck capslock still lag ; reload?!
    ^#Enter:: ; fuck capslock still lag ; reload?!
        ; SetTimer, wtf, -10
        Run, "%A_AhkPath%" /restart "%A_ScriptFullPath%" "lag" "wtf"
        return


    ^+#CapsLock::
    ^+#Enter::
        ; msgbox, % A_ComputerName " no capsloack lag " ; and WheelLeft/Right hotkey" ; wtf client ahk still lag???
        Run, "%A_AhkPath%" /restart "%A_ScriptFullPath%" "nocapslock" "wtf2"
        return

    ; CapsLock::Enter
    CapsLock & w::Up
    CapsLock & a::Left
    CapsLock & s::Down
    CapsLock & d::Right
    CapsLock & q::Home
    CapsLock & e::End
    CapsLock & r::PgUp
    CapsLock & f::PgDn
    CapsLock & z::BackSpace
    CapsLock & x::Delete
    ; CapsLock & c::^c  ; now available...
    ; CapsLock & v::Delete

    ; CapsLock & 1::Media_Prev
    ; CapsLock & 2::Media_Play_Pause
    ; CapsLock & 3::Media_Next
    CapsLock & 1::^!+[
    CapsLock & 2::^!+\
    CapsLock & 3::^!+]
    CapsLock & 4::^!+-
    CapsLock & 5::^!+=
    CapsLock & `::^!+` ; dupe
    CapsLock & esc::^!+`
    ; CapsLock & `::`

    CapsLock & F1::!F10
    CapsLock & F2::F11
    CapsLock & F3::F12

    CapsLock & Space::click

    CapsLock & Tab::^Home

    CapsLock & Alt::Alt ; finally...?

    CapsLock & Shift::Send,^{End} ; finally...?

    ; +Space::click,right ; hai xing ; bu xing le...acrobat li
    ; !Space::click,right ; hai xing ba...

    ; wow... even no need fn any more ; save it for no fn keyboard
    ; CapsLock & Up::PgUp
    ; CapsLock & Down::PgDn
    ; CapsLock & Left::Home
    ; CapsLock & Right::End
#If ; not WinActive("- 远程桌面连接 ahk_exe mstsc.exe")


; tmp: ACU
; CapsLock::Send,{shift down}{w down}

; TODO: surface only
; RAlt::RCtrl
; RAlt::Run,D:\miniconda3\envs\dcs\python.exe E:\lab\dcs\api.py
; RAlt::
;     ; StartTime := A_TickCount
;     DllCall("QueryPerformanceFrequency", "Int64*", freq)
;     DllCall("QueryPerformanceCounter", "Int64*", CounterBefore)

;     whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
;     whr.Option(4) := 13056
;     whr.Open("GET", "https://127.0.0.1:50000?test=test", true)
;     whr.Send()
;     a := whr.WaitForResponse(2)
;     ; ElapsedTime := A_TickCount - StartTime
;     ; MsgBox % a "  " ElapsedTime
;     DllCall("QueryPerformanceCounter", "Int64*", CounterAfter)
;     MsgBox % "Elapsed QPC time is " . (CounterAfter - CounterBefore) / freq * 1000 " ms" "`na=" a
;     return

F22::Send,{AppsKey} ; screen record key...
^+F22::return ; used by screen record key...
; $#^+F22::Send,{Ctrl down}{Alt down}{Shift down}{Shift up}{Alt up}{Ctrl up} ; used by touchpad three finger click...
; $#^+F24:: ; used by touchpad four finger click
;    sleep,50
 ;   Send,^;
  ;  return

F10:: ; free f10 in debug
    IfWinActive,ahk_exe code.exe
        send,{F10}
    else
        send,!{F10}
    return

; --------------------------------------------------------------------------------------------------------------------------------
;;   1.1 Mouse Button
; --------------------------------------------------------------------------------------------------------------------------------


; mouse side button
; MButton::send,{Alt down}
#If not WinActive("ahk_exe javaw.exe") and not WinActive("ahk_exe WindowsTerminal.exe") and not WinActive("ahk_exe yuanshen.exe")
    RButton::
        ; else if (A_PriorKey=="XButton2" and A_TimeSincePriorHotkey < 250) 
        if (X12State=="X2down" or GetKeyState("XButton2","P")) {
            IfWinActive, ahk_class CabinetWClass
                ExplorerReopenTab()
            else
                Send,^+t
        }
        else if (X12State=="X1down" or GetKeyState("XButton1","P"))
            send,{Ctrl Up}{Alt Down}{Tab}
        ; msgbox,% GetKeyState("XButton2","P")
        else
            ; Send,{RButton down}
            Send,{RButton}
        return
    ; RButton Up::
    ;     Send,{Rbutton up}
    ;     return
#If

LButton::
    ; else if (A_PriorKey=="XButton2" and A_TimeSincePriorHotkey < 250) 
    if (X12State=="X2down" or GetKeyState("XButton2","P")) {
        IfWinActive, ahk_class CabinetWClass
            ExplorerCloseTab()
        else
            Send,^w
    }
    ; else if (X12State=="X1down" or GetKeyState("XButton1","P"))
        ; send,{Ctrl Up}{Alt Down}{Tab}
    ; msgbox,% GetKeyState("XButton2","P")
    else
        ; Send,{RButton down}
        Send,{LButton}
    return
; RButton Up::
;     Send,{Rbutton up}
;     return

XButton1:: ; can also zoom!
    X12State:="X1down"
    ; hotkey,RButton,on
    ; hotkey,RButton Up,on
    ; if (A_PriorHotkey <> A_ThisHotkey or A_TimeSincePriorHotkey > 400) {
    ;     ; Too much time between presses, so this isn't a double-press.
    ; Send,^; ; hahaha finally
    ; no need keywait anymore!
    send,{ctrl down}
    ; hotkey,MButton,on
    if (A_PriorHotkey == "XButton1 Up" and A_TimeSincePriorHotkey < 250) { ; double click not work well
        send,{Alt Down}
        ; sleep,80
        ; sendevent,^;
        ; sleep,80
    }
    return

XButton1 Up::
    X12State:="1234"

    ; hotkey,RButton,off
    ; hotkey,RButton up,off
    Send,{Ctrl Up}{Alt Up}
    ; hotkey,MButton,off

    ; if (A_PriorHotkey != "XButton1")
    if (A_PriorKey != "XButton1")
        return
    if (A_TimeSincePriorHotkey < 250) {
        send,{Ctrl down}{Alt down}{Shift down}
        send,{MButton}
        sleep,20
        send,{Ctrl Up}{Alt Up}{Shift up}
    }
    else  if (A_TimeSincePriorHotkey > 250 and A_TimeSincePriorHotkey < 750) {
        sleep,80
        sendevent,^;
        sleep,80
        ; msgbox,2
    }
    ; msgbox,%A_TimeSincePriorHotkey%
    return

; XButton2::Shift
; XButton2 & WheelUp::Send,+{WheelUp}
; XButton2 & WheelDown::Send,+{WheelDown}
; Xbutton2 & MButton::AltTab
; Xbutton2 & WheelLeft::Send,^#{Left}
; Xbutton2 & WheelRight::Send,^#{Right}

WheelUp::
    ; Send,{WheelLeft}
    IfWinActive, ahk_exe excel.exe
        Send,{WheelLeft}
    ; else IfWinActive, ahk_exe FoxitPDFEditor.exe
    ;     send,{WheelUp 3}
    ; else if (X2FuncSwitch=="1" || A_ComputerName=="LAB")
    else if (X2FuncSwitch=="1")
        Send,{Shift Down}{WheelUp}
    ; else if (X2DoubleClick=="1")
        ; Send,!{Left}
    else
        Send,^+{Tab}
    ; hotkey,WheelUp,off
    ; hotkey,WheelDown,off
    return
WheelDown::
    ; Send,{WheelRight}
    IfWinActive, ahk_exe excel.exe
        Send,{WheelRight}
    ; else if (X2FuncSwitch=="1" || A_ComputerName=="LAB")
    else if (X2FuncSwitch=="1")
        Send,{Shift Down}{WheelDown}
    ; else if (X2DoubleClick=="1")
        ; Send,!{Right}
    else
        Send,^{Tab}
    ; hotkey,WheelUp,off
    ; hotkey,WheelDown,off
    return

XButton2::
    ; hotkey,RButton,on
    ; msgbox,wtf
    X12State:="X2down"
    hotkey,LButton,on
    hotkey,WheelUp,on
    hotkey,WheelDown,on
    ; hotkey,RButton Up,on
    if (A_PriorHotkey == "XButton2 Up" and A_TimeSincePriorHotkey < 200) {
        ; send,^+t ; 6
        ; send,{Ctrl Down}{Alt Down}
        ; send,{shift up}
        ; msgbox,wtf
        ; X2DoubleClick:="1"
        Send,^t
        ; if (X2FuncSwitch=="1")
        ;     X2FuncSwitch:="0"
        ; else
        ;     X2FuncSwitch:="1"
    }
    ; else
        ; send,{shift down}
    ; TODO: use hotkey to achieve mbutton as alttab, and use inputlevel to achieve wheelleft/right as desktop switcher
    ; msgbox,% A_PriorHotkey
    return

XButton2 Up::
    ; hotkey,RButton,off
    ; hotkey,RButton up,off
    X12State:="2234"
    ; X2DoubleClick:="0"

    hotkey,WheelUp,off
    hotkey,WheelDown,off
    hotkey,LButton,off

    Send,{shift up}{Ctrl Up}{Alt Up}
    return


; XButton and Click!

; TODO: touchpad - friendly
; WheelLeft::
;     If GetKeyState("XButton2", "P")
;         Send,{Shift Up}^1{Shift Down} ; wahaha i made it
;     Else if GetKeyState("XButton1", "P")
;         ; Send,#{Left} ; wahaha perfect ; not practical...
;         Send,{Ctrl Up}!{Left}{Ctrl Down}
;     Else
;         Send,^{PgUp}
;     return

; WheelRight::
;     If GetKeyState("XButton2", "P")
;         Send,{Shift Up}^9{Shift Down}
;     Else if GetKeyState("XButton1", "P")
;         ; Send,#{Right}
;         Send,{Ctrl Up}!{Right}{Ctrl Down}
;     Else
;         Send,^{PgDn}
;     return


#if not WinActive("ahk_exe Acrobat.exe")
MButton::
    ; Send,{MButton}
    ; hotkey,RButton,on
    ; hotkey,RButton Up,on
    if (A_PriorHotkey == "MButton" and A_TimeSincePriorHotkey < 250)
        ; msgbox,% A_PriorHotkey
        send,^t ; 6
    ; else if (A_PriorKey=="XButton2" and A_TimeSincePriorHotkey < 250) 
    else if (X12State=="X2down")
        Send,{shift up}^w
    else If GetKeyState("XButton2", "P") 
        ; and not WinActive("JupyterLab ahk_exe msedge.exe")
        Send,{Shift Up}^w ; {Shift Down}
    else If GetKeyState("XButton1", "P")
        ; Send,{Ctrl Up}^!{Tab}{Ctrl Down} ; TODO: improve ; not that useful...
        ; Send,{Ctrl Up}^r ; no need down again {Ctrl Down} ; why f5 not work???
        Send,{Ctrl Up}^h
    else if (X12State=="X1down")
        Send,{Ctrl Up}^h
        ; msgbox,wtf
    Else
        Send,{MButton} ; no need
    ; If GetKeyState("RButton","P") ; emm rbutton still function
        ; Send,
        ; msgbox,keyi
    ; If GetKeyState("LButton","P") ; tai bie niu le
        ; msgbox,bucuo
        ; Send,{F5}
    return
#if 
; ~MButton Up::
;     hotkey,RButton,off
;     hotkey,RButton Up,off
;     return

; XButton1 & XButton2::Send,wtf ; can no longer zoom...


; --------------------------------------------------------------------------------------------------------------------------------
;;   1.2 win key
; --------------------------------------------------------------------------------------------------------------------------------

#esc::send,{BackSpace} ;keyikeyi
#f1::send,{Delete} ;keyikeyi

F4::
    ; send,{alt up} ;???
    ; IfWinExist, XMR v2.1
    ;     return
    Suspend
    ; msgbox,block on
    toggleBlockInput:=1
    Suspend,On
    BlockInput, On ; not bad ; but may need change to another hotkey
    BlockInput, MouseMove ; not bad ; but may need change to another hotkey
    ; send,{alt up} ;???
    return

#If (toggleBlockInput) ; why so six...
    ; Space:: ; no need~ disable click twice fullscreen now, no lag!
    ;     Suspend,off
    ;     Send,{Space}
    ;     suspend,on
    ;     return
    F4::
        Suspend
        ; msgbox,block off
        toggleBlockInput:=0
        Suspend,Off
        BlockInput, Off ; not bad ; but may need change to another hotkey
        BlockInput, MouseMoveOff ; not bad ; but may need change to another hotkey
        return
#If

; insert number seq
#F2::
    clipboard:="1`n2"
    send,^v
    return
#F3::
    clipboard:="1`n2`n3"
    send,^v
    return
#F4::
    clipboard:="1`n2`n3`n4"
    send,^v
    return
#F5::
    clipboard:="1`n2`n3`n4`n5"
    send,^v
    return
#F6::
    clipboard:="1`n2`n3`n4`n5`n6"
    send,^v
    return
#F7::
    clipboard:="1`n2`n3`n4`n5`n6`n7"
    send,^v
    return
#F8::
    clipboard:="1`n2`n3`n4`n5`n6`n7`n8"
    send,^v
    return
#F9::
    clipboard:="1`n2`n3`n4`n5`n6`n7`n8`n9"
    send,^v
    return

#F12::
    InputBox, n ; , n, Please enter n.
    a:=""
    Loop, %n%
        a:=a A_Index "`n"
    clipboard:=a
    send,^v
    return


; #If WinActive("- 远程桌面连接 ahk_exe mstsc.exe")
    ; CapsLock::CapsLock
; #If


; stick active window on top
#t::WinSet,AlwaysOnTop,On,A

; turn of stick
+#t::WinSet,AlwaysOnTop,Off,A

; wtf
; #4::
    ; DetectHiddenWindows, On
    ; IfWinExist, ahk_exe foobar2000.exe
    ;     send,{f3}
    ; else
    ;     send,#4
    ; DetectHiddenWindows, Off
    ; return

; open Acrobat
#q::#5
; #q:: ; #5
;     ; WinGet, State, MinMax, ahk_exe Acrobat.exe
;     ; ; msgbox, % state
;     ; IfWinNotExist, ahk_exe Acrobat.exe
;     ;     State := -1
;     ; if (State = -1) { ; min
;     ;     ; MsgBox, Window is minimized
;     ;     ; if (toggleAcrobatShift==0)
;     ;     ;     send,#+5 ; jing ran ting kuai ...
;     ;     ; else
;     ;         send,^#5 ; jing ran ting kuai ...
;     ; }
;     ; else if (State = 1) { ; max
;     ;     IfWinActive, ahk_exe acrobat.exe
;     ;         WinMinimize, A
;     ;     else
;     ;         WinActivate, ahk_exe acrobat.exe
;     ; }
;     ; else ; background
;     ;     WinActivate, ahk_exe acrobat.exe
;     IfWinActive, ahk_exe Acrobat.exe
;         WinMinimize, A
;     else 
;         Send, ^#5 ; emm 666...?
;     ; WinSet, AlwaysOnTop,On, ahk_exe Acrobat.exe
;     ; WinSet, AlwaysOnTop,Off, ahk_exe Acrobat.exe
;     return

#!q::#!5
; #+q:: ; TODO auto detect acrobat tab number...
;     if (toggleAcrobatShift==1) {
;         toggleAcrobatShift:=0
;         msgbox,shit ; emm
;     }
;     else {
;         msgbox,no shit ; emm
;         toggleAcrobatShift:=1
;     }
    return

; open Word
#w::#6
#+w::#+6
#!w::#!6

; open VSCode
#c::#7
#+c::#+7
#!c::#!7
^#c::^#c ;???

; open moba ; first unpinned window
#f::#8 ; not send! not activate bug??? why recently show up????
#+f::#+8
#!f::#!8

; open onenote ; second unpinned window
#v::#9
#+v::#+9
#!v::#!9

; open todo ; third unpinned window
; #b::#0
; #+b::#+0
; #!b::#!0
#z::#0
    ; SetKeyDelay, 50
    ; Send,#b{Enter}{up}
    ; Return

; ^#z::
;     SetKeyDelay, 50
;     send,#b{left 2}{Enter}
;     return


; open Onenote
; +#n::
;     IfWinNotExist,ahk_exe ApplicationFrameHost.exe,OneNote
;         ; Run,onenote-cmd://
;         Run,onenote://
;     else
;         WinActivate,ahk_exe ApplicationFrameHost.exe,OneNote
;     return

; new quick notes
; #n::Run, onenote-cmd://quicknote?onOpen=typing
; #n::Run, onenote://quicknote?onOpen=typing

#s::
    clipboard = ;
    Send, ^c
    ClipWait, 0.5
    ; Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "(^\s+|\s+$)")
    Clipboard := RegExReplace(Clipboard, "(^\s+|\s+$)")
    ; If SubStr(Clipboard,1,7) = "http://" or SubStr(Clipboard,1,8) = "https://" or SubStr(Clipboard,1,4) = "www." 
    If RegExMatch(Clipboard,"(^http(s?)://)|(^www\.)|(\.(com|net|cn|io|ai)$)") ; open url
        Loop, parse, Clipboard, `n, `r
            Run, % "https://" RegExReplace(A_LoopField,"^http(s?)://","")
    Else if (RegExMatch(Clipboard,"^(?i)(PMID)?(\d{7,8})$")) ; open pmid
        Loop, parse, Clipboard, `n, `r
            Run, % "https://pubmed.ncbi.nlm.nih.gov/" RegExReplace(A_LoopField, "PMID")
    Else if (RegExMatch(Clipboard,"^(?i)(GSE|GPL|GSM)(\d{3,7})$")) ; open geo
        Loop, parse, Clipboard, `n, `r
            Run, % "https://www.ncbi.nlm.nih.gov/geo/query/acc.cgi?acc=" A_LoopField
    Else if (RegExMatch(Clipboard,"^(?i)(PRJNA)(\d{3,6})$")) ; open geo
        Loop, parse, Clipboard, `n, `r
            Run, % "https://www.ncbi.nlm.nih.gov/bioproject/" A_LoopField
    Else if (RegExMatch(Clipboard,"^10.\d\d\d\d/")) ; open doi
        Loop, parse, Clipboard, `n, `r
            Run, % "https://doi.org/" A_LoopField
    Else if (RegExMatch(Clipboard,"^tt\d{6,10}")) ; open imdb
        Loop, parse, Clipboard, `n, `r
            Run, % "https://www.imdb.com/title/" A_LoopField
    Else if (RegExMatch(Clipboard,"^BV.*")) ; open imdb
        Loop, parse, Clipboard, `n, `r
            Run, % "https://www.bilibili.com/video/" A_LoopField
    Else if (RegExMatch(Clipboard,"^/"))  ; open folder
        Loop, parse, Clipboard, `n, `r 
            {
                hwnd:=WinExist("ahk_class CabinetWClass")
                if (hwnd)
                    Explorer_Navigate_New_Tab(RegExReplace(A_LoopField,"^/","X:/"), hwnd)
            }
    Else if (RegExMatch(Clipboard,"(^.:\\)|(^\\\\wsl)"))  ; open folder
        Loop, parse, Clipboard, `n, `r 
        {
            ; Run, % "file:///" A_LoopField
            hwnd:=WinExist("ahk_class CabinetWClass")
            if (hwnd)
                ; if (RegExMatch(A_LoopField,"(^.:\\)|(^\\\\wsl)"))
                Explorer_Navigate_New_Tab(A_LoopField, hwnd)
        }
    Else { ; search google
        ; Loop, parse, Clipboard, `n, `r
        Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "(^\s+|\s+$)")
        Run, % "https://www.google.com/search?q=" Clipboard
    }
    return

^+#s:: ; search multiple line ; thesaurus
    clipboard = ;
    Send, ^c
    ClipWait, 0.5
    Clipboard := RegExReplace(Clipboard, "(^\s+|\s+$)")
    Loop, parse, Clipboard, `n, `r
        ; Run, % A_LoopField
        Run, % "https://www.google.com/search?q=gastric+cancer+" A_LoopField
    return


#!s:: ; search gene
    clipboard = ;
    Send, ^c
    ClipWait, 0.5
    ; Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "(^\s+|\s+$)")
    Clipboard := RegExReplace(Clipboard, "(^\s+|\s+$)")
    Loop, parse, Clipboard, `n, `r
        ; Run, % "https://www.ncbi.nlm.nih.gov/geo/query/acc.cgi?acc=" A_LoopField
        ; Run, % "https://www.genecards.org/Search/Keyword?queryString=" A_LoopField
        ; Run, % "https://www.genecards.org/cgi-bin/carddisp.pl?gene=" A_LoopField
        Run, % "https://www.proteinatlas.org/search/" A_LoopField
        ; Run, % "https://www.thesaurus.com/browse/" A_LoopField
        ; Run, % "https://pubmed.ncbi.nlm.nih.gov/" A_LoopField
    return

^#!s:: ; search google scholar
    clipboard = ;
    Send, ^c
    ClipWait, 0.5
    Clipboard := RegExReplace(RegExReplace(Clipboard, "\r?\n"," "), "(^\s+|\s+$)")
    ; Run, % "https://sci-hub.se/" Clipboard ; ru ; shop ; st ; cat ; ee ; ren ; wf ; ...
    ; Run, % "https://sc.panda321.com/scholar?q=" Clipboard ; has
    ; Run, % "https://scholar.lanfanshu.cn/scholar?q=" Clipboard ; has limit
    ; Run, % "https://www.semanticscholar.org/search?q=" Clipboard ; too slow?

    ; Run, % "https://deepsearch.cc/scholar?q=" Clipboard
    ; Run, % "https://x.sci-hub.org.cn/scholar?q=" Clipboard
    ; Run, % "https://xs.lsqwl.org/scholar?as_sdt=0%2C5&q=" Clipboard ; bu cuo zhi chi lookup ; finally...
    Run, % "https://scholar.google.com/scholar?q=" Clipboard ; bu cuo zhi chi lookup ; finally...
    ; Run, % "https://xs2.dailyheadlines.cc/scholar?hl=zh-CN&as_sdt=0%2C5&q=" Clipboard ; bu cuo zhi chi lookup
    ; Run, % "https://xs2.zidianzhan.net/scholar?hl=zh-CN&as_sdt=0%2C5&q=" Clipboard ; bu cuo zhi chi lookup
    ; Run, % "https://www.scidown.cn/go/google.php?hl=zh-CN&q=" Clipboard ; bu cuo zhi chi lookup
    return


; show/hide magnify
#h::
    If WinExist("ahk_exe Magnify.exe") {
        WinHide, ahk_exe Magnify.exe
        HideMagifier:="T"
        ; msgbox,unhide
    }
    else if (HideMagifier=="T") {
        HideMagifier:="F"
        WinShow, ahk_exe Magnify.exe
        ; WinHide, ahk_exe Magnify.exe
    }

    ; now available...
    return


; show/hide taskbar
#j::
    if DllCall("Shell32\SHAppBarMessage", "UInt", 4, "Ptr", &APPBARDATA, "Int")
    {
        NumPut(2, APPBARDATA, 40)
        DllCall("Shell32\SHAppBarMessage", "UInt", 10, "Ptr", &APPBARDATA)
    }
    else
    {
        NumPut(1, APPBARDATA, 40)
        DllCall("Shell32\SHAppBarMessage", "UInt", 10, "Ptr", &APPBARDATA)
    }
    return

#+e:: ; TODO: decrease lag
    ; Runwait,%AppData%\Microsoft\Windows\Recent\
    path:="C:\Users\" A_UserName "\AppData\Roaming\Microsoft\Windows\Recent"
    ; WinWaitActive, 最近使用的项目 ahk_exe explorer.exec
    ; sleep,500
    ; send,{tab}
    WinActivate, ahk_class CabinetWClass
    WinWaitActive, ahk_class CabinetWClass
    sleep,100
    Explorer_Navigate_New_Tab(path)
    return

; let's use tab!
#e::
    ; let's merge!
    WinGet, n, Count, ahk_class CabinetWClass
    if (n=1) {
        If WinActive("ahk_class CabinetWClass")
            Send,^t
        else {
            WinActivate, ahk_class CabinetWClass
            WinWaitActive, ahk_class CabinetWClass
            Send,+{F6}^t ; ???
        }
        return
    }
    ; else merge
    r:=""
    dict:={}
    for w in ComObjCreate("Shell.Application").Windows {
        if (w.hwnd != hwnd)
        if not (dict[w.hwnd])
            dict[w.hwnd]:=[]
        dict[w.hwnd].Push(w.Document.Folder.Self.Path)
    }
    maxkey:=0
    maxnum:=0
    For key, v in dict {
        if (v.Length()>maxnum) {
            maxkey:=key
            maxnum:=v.Length()
        }
    }
    WinActivate,ahk_id %maxkey%
    WinWaitActive,ahk_id %maxkey%

    For key, v in dict {
        if (key==maxkey) {
            For key2, v2 in dict {
                if (key2!=maxkey) {
                    For index, value in v2 {
                        Explorer_Navigate_New_Tab(value,key)
                    }
                }
            }
        }
        else {
            WinClose,ahk_id %key%
        }
    }
    Return


; wtf
#a::
    if (A_ComputerName="DESKTOP-92789C7")
        Send,#a
    Else
        Send,#n
    return
^#a::Send,#a
; #`::Send,#a

#`::
    if (A_ComputerName!="LAB") {
        IfWinNotExist,Profile 1 - Microsoft​ Edge Beta ahk_exe msedge.exe
        {
            Run, "C:\Program Files (x86)\Microsoft\Edge Beta\Application\msedge.exe"
            ; WinWaitActive, Profile 1 - Microsoft​ Edge Beta ahk_exe msedge.exe
            ; send,!+2
        }
        else
        {
            IfWinActive, Profile 1 - Microsoft​ Edge Beta ahk_exe msedge.exe
                WinMinimize, Profile 1 - Microsoft​ Edge Beta ahk_exe msedge.exe
            else
                WinActivate,Profile 1 - Microsoft​ Edge Beta ahk_exe msedge.exe ; TODO: choose the latest used window
        }
        return
    }
    IfWinNotExist,Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe
    {
        Run, "C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\ImplicitAppShortcuts\e35895725797c7e3\Microsoft Edge Dev.lnk"
        ; WinWaitActive, Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe
        ; send,!+2
    }
    else
    {
        IfWinActive, Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe
            WinMinimize, Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe
        else
            WinActivate,Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe ; TODO: choose the latest used window
    }
    return


^#z::
    IfWinNotExist,ahk_exe zotero.exe
    {
        ; Run, wt.exe ; -d %USERPROFILE%
        ; WinWaitActive, ahk_exe zotero.exe
        ; send,!+2
    }
    else
    {
        IfWinActive, ahk_exe zotero.exe
            WinMinimize, ahk_exe zotero.exe
        else
            WinActivate,ahk_exe zotero.exe ; TODO: choose the latest used window
    }
    return

; #Shift:: ; stupid!!!
;     send,{ctrl down}{lwin down}{space}
;     ; KeyWait, lwin
;     ; send,{ctrl up}{lwin up}
;     return
; #Space::Send,^#{Space}
; ^#Space::Send,#{Space}
#CapsLock::Send,^#{Space} ;MsgBox, 666

; #F3::Send,#=
; #F2::Send,#{esc} ; use ctrl+alt!!!
#WheelUp:: ;Send,^!{WheelUp}
    DetectHiddenWindows, On
    IfWinExist, ahk_exe Magnify.exe
        send,^!{WheelUp}
    else {
        send,#=
    }
    DetectHiddenWindows, Off
    If WinExist("ahk_exe Magnify.exe")
        WinHide, ahk_exe Magnify.exe
    return

; #WheelDown::Send,#- ; zhong yu shu fu le
#WheelDown::
    Send,^!{WheelDown}
    If WinExist("ahk_exe Magnify.exe")
        WinHide, ahk_exe Magnify.exe
    return
; #MButton::Send,#{esc} ; zhong yu shu fu le

; #WheelLeft::send,^!{left} ; ???
; #WheelRight::send,^!{Right}
#WheelLeft::Send,^#{Left}
#WheelRight::Send,^#{Right}
; #+WheelUp::send,^!{left}
; #+WheelDown::send,^!{Right}



; --------------------------------------------------------------------------------------------------------------------------------
;;   1.3 ctrl/alt/shift key
; --------------------------------------------------------------------------------------------------------------------------------



; open Task Manager
+^!d::
    sleep,150 ;wtf??? it was 50 but when admin must 150??? to avoid reset layout
    Run,taskmgr ;"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\System Tools\Task Manager.lnk"
    return

; open powershell
^!s:: ; bu cuo bu cuo
+^!s::
    IfWinNotExist,ahk_exe WindowsTerminal.exe
    {
        Run, wt.exe ; -d %USERPROFILE%
        ; WinWaitActive, ahk_exe WindowsTerminal.exe
        ; send,!+2
    }
    else
    {
        IfWinActive, ahk_exe WindowsTerminal.exe
            WinMinimize, ahk_exe WindowsTerminal.exe
        else
            WinActivate,ahk_exe WindowsTerminal.exe ; TODO: choose the latest used window
    }
    return

; open cmd
; use less cmd!!!

; +^!c::Run, notepad++.exe E:\test\clipboard.txt
+^!c::Run, "D:\scoop\apps\vscode\current\Code.exe" E:\test\clipboard.txt

; open Notepad++
; +^!n::
;     IfWinNotExist,ahk_exe notepad++.exe
;     {
;         Run,notepad++.exe
;         WinWaitActive, ahk_exe notepad++.exe
;         send,!+2
;     }
;     else
;         WinActivate,ahk_exe notepad++.exe
;     return

; reconnect Tsinghua-5G
+^!w::
    ; ClipSaved := ClipboardAll
    ; Runwait,%comspec% /c netsh wlan show interface | clip,,hide
    ; wifiName = % RegExReplace(clipboard, "s).*?\R\s+SSID\s+:(\V+).*", "$1")
    ; if StrLen(wifiName) > 400
    ;     wifiName = Tsinghua-5G
    ; Runwait,%comspec% /c netsh wlan connect name=%wifiName% ssid=%wifiName% & pause
    ; Clipboard := ClipSaved
    ; ClipSaved := ""
    return

+^!q::
    Run *RunAs powershell.exe -Command "Disable-NetAdapterBinding -Name '以太网' -ComponentID 'ms_tcpip6'; Enable-NetAdapterBinding -Name '以太网' -ComponentID 'ms_tcpip6'"
    return

    ; Run *RunAs %comspec% /c netsh int isatap set state disable & netsh int isatap set state enable & exit
    ; Run *RunAs %comspec% /c netsh int ip set dns "        * 14" static 8.8.8.8 & netsh int ip set dns "        * 15" static 8.8.8.8 & pause & exit ; wtf
    ; Run *RunAs %comspec% /c netsh int ip set dns 62 static 8.8.8.8 & netsh int ip set dns 63 static 8.8.8.8 & pause & exit
    ; Run *RunAs %comspec% /c netsh int ip set dns 12 static 8.8.8.8 & netsh int ip set dns 20 static 8.8.8.8 & pause & exit
    ; Run *RunAs %comspec% /c netsh int ip set dns 12 static 114.114.114.114 & netsh int ip set dns 20 static 114.114.114.114 & pause & exit



; 6 number date
+^!r::
    FormatTime, date,, yyMMdd
    SendEvent,%date%
    return

+^!t::
    FormatTime, date,, HHmmss
    SendEvent,%date%
    return

; open windows spy
+^!p::
    IfWinExist, Window Spy ahk_exe autohotkey.exe
        WinActivate
    else
        Run, "C:\Program Files\AutoHotkey\WindowSpy.ahk"
    return
; +^!p::Run, "C:\Program Files\scoop\apps\ahk\current\WindowSpy.ahk"

; open vcxsrv
; +^!m::Run, E:\test\etc\multiple.xlaunch

+^!x::
    ; Run, "E:\test\ahk\X-MacroRecorder.ahk"
    IfWinExist, XMR v2.1 ahk_exe AutoHotKey.exe
        WinActivate
    else
        Run, "E:\test\ahk\X-MacroRecorder.ahk"
    return



^+!a::
    IfWinNotExist,ahk_exe MobaXterm.exe
    {
        Run,D:\MobaXterm\MobaXterm.exe
        WinWaitActive, ahk_exe MobaXterm.exe
        sleep,800
        WinActivate, ahk_exe MobaXterm.exe
        sleep,800
        send,!+2
    }
    else
        WinActivate,ahk_exe MobaXterm.exe
    return

#If not WinActive("- 远程桌面连接 ahk_exe mstsc.exe") and not WinActive("远程桌面 ahk_exe ApplicationFrameHost.exe")

; minimize active window except desktop
!x::
    IfWinActive,ahk_class WorkerW
        return
    IfWinActive,ahk_class CWebviewControlHostWnd
        WinActivate, ahk_class WeChatMainWndForPC
    WinMinimize,A
    return

!^x::
    IfWinActive,ahk_class WorkerW
        return
    WinRestore,A
    return

!+x::
    IfWinActive,ahk_class WorkerW
        return
    WinMaximize, A
    return
#If ; not WinActive("- 远程桌面连接 ahk_exe mstsc.exe")

; #If WinActive("- 远程桌面连接 ahk_exe mstsc.exe")
;     #!x::
;         WinMinimize, A
;         return
; #If ; not WinActive("- 远程桌面连接 ahk_exe mstsc.exe")


; up down left right
^#w::Send,{Up}
^#s::Send,{Down} ; emm ye bu shi mei yong


; ^#+w::Send,+{Up}
; ^#+s::Send,+{Down}
; ^#+a::Send,+{Left}
; ^#+d::Send,+{Right} ; farewell

; PgUp etc?

; open virtual keyboard with surface pen single click
$#F20::
    sleep,200
    SetKeyDelay, 20
    ; Send,#b{right 4}{enter} ; without USB
    Send,#b{right 5}{enter}
    return

; ^.:: ;AppsKey
^esc::Send,{AppsKey}

^`::send,^/

; for editing
$+-::Send,{Text}_
/::send,{NumpadDiv}

^+9::Send,+9+0{left}
^+'::Send,+'+'{left}


#+a::send,#{PrintScreen}
; for steam ; tu ran jiu not work le... zhi hao zhi jie gai xbox game bar she zhi le:
#!a::Sendevent,#!{PrintScreen}
F13 up::Sendevent,#!{PrintScreen}
F14 up::Sendevent,#!g

; ^!f10::toggle get word from clipboard

^+z::send,^y ; bu xing zai shuo

; multi window manager
; !`::
;     MoveCursor()
;     return

; #`::
;     Send,#+{Left}
;     ; MoveCursor()
;     return

!;::Sendevent,^; ;wei sha zao dian mei xiang dao
; ^!q::Sendevent,^; ;wei sha zao dian mei xiang dao
!^q::send,^!w
; ^!s::SendEvent, ^; ;wei sha send bu xing???

; ^space::click
; ^+space::click,Right

; if no touchpad
; wheelleft/right as alttab(shift)


; ^!f::send,^!{F12} ; fangdajing magnifier wcnm!
; ^!d::msgbox,magnify cnm
^!d::
    IfWinNotExist,ahk_exe Taskmgr.exe
    {
        Run, Taskmgr ; -d %USERPROFILE%
        ; WinWaitActive, ahk_exe Taskmgr.exe
        ; send,!+2
    }
    else
    {
        IfWinActive, ahk_exe Taskmgr.exe
            WinMinimize, ahk_exe Taskmgr.exe
        else
            WinActivate,ahk_exe Taskmgr.exe ; TODO: choose the latest used window
    }
    return
^!l::msgbox,magnify cnm
^!m::msgbox,magnify cnm

^!g::
    IfWinNotExist,ahk_exe GitHubDesktop.exe
    {
        Run, "C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Scoop Apps\GitHub Desktop.lnk"
        ; WinWaitActive, ahk_exe GitHubDesktop.exe
        ; send,!+2
    }
    else
    {
        IfWinActive, ahk_exe GitHubDesktop.exe
            WinMinimize, ahk_exe GitHubDesktop.exe
        else
            WinActivate,ahk_exe GitHubDesktop.exe ; TODO: choose the latest used window
    }
    return

    
; tomorrow: try mouse!

Insert::return ; msgbox,wtf

; `::Send,0 ; not useful...
; 0::Send,``
; #`::Send,``

; F1::
;     if (A_ComputerName="LAPTOP-SLNGK7IH")
;         msgbox, % A_ComputerName " wtf!!!"
;     else
;         msgbox, % A_ComputerName " wtf2!!!"
;     return

^#x::
    ; msgbox,wtf
    send,^!{home} ; not work...
    return
; x::home

#!v::
    Clipboard := RegExReplace(Clipboard, "\\", "/")
    Clipboard := RegExReplace(Clipboard, "X:")
    Clipboard := RegExReplace(Clipboard, """")
    send,^v
    return


; #If
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;; 2 Per Window
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------



; GroupAdd,AltWForClose,ahk_exe notepad++.exe
; GroupAdd,AltWForClose,ahk_exe Explorer.EXE
; GroupAdd,AltWForClose,ahk_exe rstudio.exe

; #IfWinActive,ahk_group AltWForClose
;     !w::Send,^w
; #IfWinActive

#IfWinActive,ahk_exe notepad++.exe
    !w::Send,^w
    !z::Send,^+t
    ^t::send,^n
    !t::send,^n
    !2::send,^+{tab}
    !3::send,^{tab}

    ^!/::Send,^{/}!+{Down}^{/} ;copy to next line and comment original line
    ^+/::Send,!{Up}^/{Down}^{/} ;switch comment with prev line

    ; emm
    ^d::^+l
    
#IfWinActive

#IfWinActive,ahk_exe ditto.exe
    WheelRight::WheelLeft
    WheelLeft::WheelRight
    +WheelUp::send,+{WheelDown}
    +WheelDown::send,+{WheelUp}
#IfWinActive


; F1::
;     WinGet windows, List, ahk_class CabinetWClass
;     r:=""
;     Loop %windows%
;     {
;         id := windows%A_Index%
;         ; WinGetTitle wt, ahk_id %id%
;         WinGetText, wt, ahk_id %id%
;         r .= StrLen(wt) . "`n"
;         if (StrLen(wt)<200) {
;             WinGetTitle, a, ahk_id %id%
;         }
;     }
;     MsgBox %r%
;     Return

GetActiveExplorerPath() {
    hwnd := WinActive("ahk_class CabinetWClass")
    activeTab := 0
    try ControlGet, activeTab, Hwnd,, % "ShellTabWindowClass1", % "ahk_id" hwnd
    for w in ComObjCreate("Shell.Application").Windows {
        if (w.hwnd != hwnd)
            continue
        if activeTab {
            static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
            shellBrowser := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
            DllCall(NumGet(numGet(shellBrowser+0)+3*A_PtrSize), "Ptr", shellBrowser, "UInt*", thisTab)
            if (thisTab != activeTab)
                continue
            ObjRelease(shellBrowser)
        }
        return w.Document.Folder.Self.Path
    }
}


Explorer_Navigate_New_Tab(FullPath, hwnd="") {

	if (hwnd) { ; make sure this Explorer window is active
		WinActivate, ahk_id %hwnd%
		WinWaitActive, ahk_id %hwnd%
	}
	else
		hwnd := WinExist("A") ; if omitted, use active window
    WinGet, ProcessName, ProcessName, % "ahk_id " hwnd
    if (ProcessName != "explorer.exe")  ; not Windows explorer
        return
	
	; Windows Explorer is the active window
    a:=ComObjCreate("Shell.Application").Windows.Count
	Send, ^t ; add a new tab
    Loop, 40 ; 1s
    {
        b:=ComObjCreate("Shell.Application").Windows.Count
        if (a!=b) {
            break
        }
        sleep,25
    }
	Sleep, 50 ; for safety
    Explorer_Navigate_Tab(FullPath, hwnd)
}

Explorer_Navigate_Tab(FullPath, hwnd="", warning=True) {
    ; see https://www.autohotkey.com/boards/viewtopic.php?p=489575#p489575
    hwnd := (hwnd="") ? WinExist("A") : hwnd ; if omitted, use active window
    ; WinGet, ProcessName, ProcessName, % "ahk_id " hwnd
    ; if (ProcessName != "explorer.exe")  ; not Windows explorer
        ; return
    For pExp in ComObjCreate("Shell.Application").Windows {
;         a:=pExp.Document.Folder.Self.Path
;         if (a="`:`:{20D04FE0-3AEA-1069-A2D8-08002B30309D}") {
; 			pExp.Navigate("file:///" FullPath)
;             return
;         }
        if (pExp.hwnd = hwnd) { ; matching window found
			activeTab := 0
			try ControlGet, activeTab, Hwnd,, % "ShellTabWindowClass1", % "ahk_id" hwnd
			if activeTab {
				static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
				shellBrowser := ComObjQuery(pExp, IID_IShellBrowser, IID_IShellBrowser)
				DllCall(NumGet(numGet(shellBrowser+0)+3*A_PtrSize), "Ptr", shellBrowser, "UInt*", thisTab)
				if (thisTab != activeTab) ; matching active tab
					continue
				ObjRelease(shellBrowser)
			}
            if (warning and pExp.Document.Folder.Self.Path!="`:`:{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                msgbox,wtf! ; interesting
			pExp.Navigate("file:///" FullPath)
			return
		}
    }
}


stack(command, value = 0) {
    static
    if !pointer
        pointer = 10000
    if (command = "push") {
        _p%pointer% := value
        pointer -= 1
        return value
    }
    if (command = "pop") {
        pointer += 1
        return _p%pointer%
    }
    if (command = "peek") {
        next := pointer + 1
        return _p%next%
    }
    if (command = "empty") {
        if (pointer == 10000)
            return "empty"
        else
            return 0
    }
}

ExplorerCloseTab(){
    wtf:=GetActiveExplorerPath()
    if not (wtf="`:`:{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
        stack("push",wtf)
    send,^w
    ; sleep,100
    ; WinActivate,ahk_class CabinetWClass
}

ExplorerReopenTab(){
    wtf:=stack("pop")
    if (wtf="")
        msgbox,emptyed
        ; Send,^t
    else
        Explorer_Navigate_New_Tab(wtf)
}

#If WinActive("ahk_class CabinetWClass") and A_ComputerName!="DESKTOP-EIJP5FP"
    ^w::
    !w::
        ExplorerCloseTab()
        Return
    !z::
        ExplorerReopenTab()
        return
    !s:: ; save status?
    ^s::
        r:=""
        ; WinGet windows, List, ahk_class CabinetWClass
        ; Loop %windows%
        ; {
        ;     id := windows%A_Index%
        ;     ; WinGetTitle wt, ahk_id %id%
        ;     WinGetText, wt, ahk_id %id%
        ;     r .= StrLen(wt) . "`n"
        ;     if (StrLen(wt)<200) {
        ;         WinGetTitle, a, ahk_id %id%
        ;     }
        ; }
        ; MsgBox %r%
        ; Return
        for w in ComObjCreate("Shell.Application").Windows {
            r .= w.Document.Folder.Self.Path . "`n"
        }
        msgbox, % r
        filename:="E:\test\log\otherlogs\openedfolers.log"
        FileRead, a, %filename%
        FileDelete, %filename%
        b:="`n"
        FileAppend, %r%%b%%a%, %filename%
        return

    !q::
        ; Send,!d^c^t!d^v{Enter}
        Explorer_Navigate_New_Tab(GetActiveExplorerPath())
        return


    !F1::Send,^1
    !F2::Send,^2
    !F3::Send,^3
    !F4::Send,^4

#IfWinActive


#If WinActive("ahk_class CabinetWClass") ; and A_ComputerName=="DESKTOP-EIJP5FP"
    !w::Send,^w
    
    ; !w::Send,^w
    !g::Send,!{Right}
    !a::Send,!{Left}
    ; !e::Send,!{Up} ;wtf??? sendinput may open another window???
    ; !e::sendevent,!{Up} ;still occasional
    !e::sendevent,+{F6}!{Up} ; come on
    
    ; !r::send,{f5}
    !r::click,right

    ; 22h2 tab, nice!
    !1::Send,^1
    !2::Send,^+{tab}
    !3::Send,^{tab}
    !4::Send,^1^+{tab}
    !t::Send,^t

#If


#IfWinActive, ahk_exe rstudio.exe
    !w::Send,^w
    !f::Send,!r ; for the damn menubar mnemonics

    ^!/::Send,^{/}!+{Down}^{/} ;copy to next line and comment original line
    ^+/::Send,!{Up}^/{Down}^{/} ;switch comment with prev line

    ; workaround for the run current line cannot run help example bug
    !Enter::Send,^{Enter}
    !CapsLock::Send,^{Enter}

    ^Enter::
    ^CapsLock::Send,{End}{Enter}

    ^+Enter::
    ^+CapsLock::Send,{Home}{Enter}{Up}

    ; ::libr::library(Seurat)`nlibrary(glue)`nlibrary(ggplot2)`nlibrary(cowplot)`nlibrary(reticulate)`n
#IfWinActive


#IfWinActive, ahk_exe Acrobat.exe
    !w::Send,^w
    !a::Send,!{Left}
    !g::Send,!{Right}
    ; !q::Send,!w ; Window

    $^+z::Send,^+z

    !1::send,!{w}{1}
    !2::send,^+{tab}
    !3::send,^{tab}
    !4::send,!{w}{1}^+{tab}

    ; WheelLeft::Left
    ; WheelRight::Right
    ; WheelLeft::Send,{WheelLeft}
    ; WheelRight::Send,{WheelRight}
    MButton::Mbutton

    !WheelUp::Send,{PgUp}
    !WheelDown::Send,{PgDn}

    !z::Send,!+t
    !e::Send,^k
#IfWinActive


#IfWinActive,ahk_exe FoxitPDFEditor.exe
    WheelUp::
        if GetKeyState("XButton2", "P")
            Send,+{WheelUp 3}
        else
            Send,{WheelUp 3}
        return
    WheelDown::
        if GetKeyState("XButton2", "P")
            Send,+{WheelDown 3}
        else
            Send,{WheelDown 3}
        return

    #F2::send,^+{tab}
    #F3::send,^{tab}

    
    ; ; Alt::return
    ; 2::
    ;     if GetKeyState("alt", "P")
    ;         Sendevent,^+{PgUp}
    ;     return
    ; 3::
    ;     if GetKeyState("alt", "P")
    ;         Sendevent,^+{PgDn}
    ;     return
    
    ; tab::
    ;     if GetKeyState("alt", "P")
    ;         Send,!{tab}
    ;     return

    ; SendLevel, 1
    ; #InputLevel, 1
    ; Alt::F15
    ; #InputLevel, 0
    alt::
        ; Sendevent, {Blind}{F15 DownR}  ; DownR is like Down except that other Send commands in the script won't assume "b" should stay down during their Send.
        wcnm:=1
        return
    alt up::
        ; Sendevent, {Blind}{F15 up}
        ; if (A_PriorHotkey=="alt")
        ;     send,{alt down}
        if (A_Priorkey=="LAlt")
            send,{alt down}
        send,{alt up}
        wcnm:=0
        return
    ; SendLevel, 0
    ; #InputLevel, 0
    ; SendLevel, 0

    ; F15 up::send,{alt up}
    ; SendLevel, 0
    ; #InputLevel, 2

    ; !2::Send,^+{PgUp}
    ; !3::Send,^+{PgDn}
    ; !4::send,!{w}{1}^+{tab}
#IfWinActive

#if wcnm=1 ; wow basically made it!
    ; SendLevel,3
    1::send,{alt down}1{alt up}
    2::send,^+{tab}
    3::send,^{tab}
    w:: ; actually reopen is feasiable... detect the darker color in titlebar, then right click, B, save to stack
        send,^w
        return
    ^q::send,^!w
    ^w::send,^!w
    e::Send,^k
    f::send,!f
    v::send,!v
    r::Send,^k!r{Enter}
    a::Send,!{Left}
    g::Send,!{Right}
    tab::send,{alt down}{tab}
    alt up::
        ; Sendevent, {Blind}{F15 up}
        send,{alt up}
        wcnm:=0
        return
    x:: ; send,!x
        WinMinimize,A
        wcnm:=0
        return
#if

; alt up::send,{alt up}
; #InputLevel, 0
; F15 & tab::send,{alt down}{tab}{alt up}
; F15 & tab::msgbox,wtf
; F15 up::msgbox,cnm

#IfWinActive, ahk_exe zotero.exe
    !a::Send,!{Left}
    !g::Send,!{Right}
    !2::Send,^+{Tab}
    !3::Send,^{Tab}
#IfWinActive




#IfWinActive, ahk_exe powershell.exe
    RButton::return ;msgbox,fuck
    ; time to introduce admin?
#IfWinActive

#IfWinActive, ahk_exe WindowsTerminal.exe
    ; RButton::
    ;     return
    !g::Send,{Right}{Enter}
#IfWinActive

#If WinActive("ahk_exe ONENOTE.exe") or WinActive("ahk_exe ApplicationFrameHost.exe") ; uwp, onenote etc
    !a::Send,!{Left}
    !g::Send,!{Right}
    ; !e::
    ; return
    $^.::send,^. ; ???

    +WheelUp::WheelLeft
    +WheelDown::WheelRight
    XButton2 & WheelDown::WheelRight
    XButton2 & WheelUp::WheelLeft

    ; WheelLeft::return ; Send,{WheelLeft}
    ; WheelRight::return ; Send,{WheelRight}

    ^+9::
        send,+9
        sleep,50
        send,+0
        sleep,50
        send,{left}
        return

    /::send,{NumpadDiv} ; keyi

    ^`::send,^.

    !s::Send,^!0

    !1::send,^+{tab}
    !4::send,^{tab}
    !2::send,^{PgUp}
    !3::send,^{PgDn}
    ; !5::send,^5

    ; ??
    CapsLock & w::msgbox,wtf
    ; CapsLock & s::msgbox,wtf ; Send,{Down}
    ; CapsLock & a::msgbox,wtf ; Send,{Down}
    ; CapsLock & d::msgbox,wtf ; Send,{Down}
#IfWinActive


#IfWinActive,ahk_exe MobaXterm.exe
    !x::
        CoordMode, Mouse, Screen
        MouseGetPos, xpos, ypos
        CoordMode, Mouse, Window
        WinGetPos, , , Width, , ahk_exe MobaXterm.exe
        a:=Width-250
        click,%a%,30
        CoordMode, Mouse, Screen
        DllCall("SetCursorPos", int, xpos, int, ypos)
        return
    !1::Send,^+{Tab}
    !4::Send,^{Tab}
    
    !WheelUp::send,!{Up}
    !WheelDown::send,!{Down}

    +WheelUp::WheelLeft ; ???
    +WheelDown::WheelRight
    XButton2 & WheelDown::WheelRight
    XButton2 & WheelUp::WheelLeft

    ; ^+v::^v ; emm

    ; not elegant...
    !Left::Send,^{Left}
    !Right::Send,^{Right}
    ^BackSpace::Send,!{BackSpace}
    ^Delete::Send,!d
    !Delete::Send,!d

#IfWinActive

#IfWinActive,ahk_exe MATLAB.exe
    ^d::send,{End}{Home 2}+{Down}{Del}
    ; !+Down::send,{End}{Home 2}+{End}^c{Right}+{Enter}^v

    !CapsLock::send,!1
    !Enter::send,!1
#IfWinActive

#IfWinActive,ahk_exe WINWORD.exe
    +WheelUp::Send,{WheelLeft}
    +WheelDown::Send,{WheelRight}

    XButton2 & WheelUp::Send,{WheelLeft}
    XButton2 & WheelDown::Send,{WheelRight}

    ^d::Send,{Esc}{Home}+{End}{Del}

    ; !v::Send,{AppsKey} ; ???

#IfWinActive

#IfWinActive,ahk_exe POWERPNT.EXE
    +WheelUp::Send,{WheelLeft}
    +WheelDown::Send,{WheelRight}
    XButton2 & WheelUp::Send,{WheelLeft}
    XButton2 & WheelDown::Send,{WheelRight}
#IfWinActive

#IfWinActive, devenv.exe ;vs
    ; ^!/::Send,^/^+d^/ ;copy to next line and comment original line
#IfWinActive


#If A_ComputerName!="MBE" and not WinActive("- 远程桌面连接 ahk_exe mstsc.exe") and not WinActive("ahk_exe ApplicationFrameHost.exe") and not WinActive("ahk_exe winword.exe") and (not WinActive("ahk_exe acrobat.exe") or A_ComputerName=="LAB" )  and RemoteControl=0 ;  and A_ScreenWidth=3840 and A_ScreenHeight=2160 

    ; VK9C::
    ; WheelLeft::Send,{F13}
    ; WheelRight::Send,{F14}
    WheelLeft::
        ; msgbox,% GetKeySC("WheelRight")
        ; If GetKeyState("SC001")
            ; return
        If GetKeyState("XButton2", "P")
            Send,{Shift Up}^1 ; {Shift Down} ; wahaha i made it
        Else if GetKeyState("XButton1", "P")
            ; Send,#{Left} ; wahaha perfect ; not practical...
            Send,{Ctrl Up}!{Left}{Ctrl Down}
        Else
            ; Send,^{PgUp}
            Send,^+{Tab}
        return

    ; VK9D::
    WheelRight::
        ; If GetKeyState("SC001")
        ;     return
        If GetKeyState("XButton2", "P")
            Send,{Shift Up}^9 ; {Shift Down}
        Else if GetKeyState("XButton1", "P")
            ; Send,#{Right}
            Send,{Ctrl Up}!{Right}{Ctrl Down}
        Else
            ; Send,^{PgDn}
            Send,^{Tab}
        return
#If

; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;;   2.1 msedge and Jupyterlab
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; #If WinActive("ahk_exe JupyterLab.exe") or WinActive("JupyterLab - Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe")
;     !2::Send,^+[
;     !3::Send,^+]
; #If

#If WinActive("ahk_exe msedge.exe") or WinActive("ahk_exe chrome.exe")
    !1::send,^1
    !2::send,^{PgUp}
    !3::send,^{PgDn}
    !4::send,^9
    !5::send,^5

    ; !+1::send,^1
    ; !+2::send,^2
    ; !+3::send,^3
    ; !+4::send,^4
    ; !+5::send,^5

    ; !e::send,^+a
    ; ^+a::Send,^+{F12}
    ; 
    !t::send,^t
    !z::send,^+t
    !h::send,^h
    F2::send,{F6} ; now can rename in jupyterlab 3.4.2~ ; emm undo hao le zai shuo
    ; F1::send,{F6}

    ; WheelLeft::Send,^{PgUp}
    ; WheelRight::Send,^{PgDn}
#IfWinActive


#If WinActive("-微信读书 ahk_exe msedge.exe") ; finally
; ~LButton::
;  if (A_PriorHotkey == "LButton" and A_TimeSincePriorHotkey < 200)
; msgbox,% A_PriorHotkey
;     Send,{Ctrl down}{Alt down}{Shift down}{Ctrl Up}{Alt Up}{Shift up}
~RButton::Send,{Ctrl down}{Alt down}{Shift down}{Ctrl Up}{Alt Up}{Shift up}
#If

#If WinActive("Human ahk_exe msedge.exe") ; finally
    !e::
        ; msgbox,111
        x:=3200
        y:=800
        while (True) {
            PixelGetColor, a, x, y
            if (a==0x50971E) {
                click
                break
            }
            sleep,10
        }
     return
#If

; #If WinActive("ahk_exe msedge.exe") and not WinActive("JupyterLab 和另外 ahk_exe msedge.exe",,":")
#If (WinActive("ahk_exe msedge.exe") or WinActive("ahk_exe chrome.exe")) and not isJLabActive()
    !a::send,!{left}
    !g::send,!{right}
    !q::send,^+k
    !r::send,{F5}
    !w::send,^w
#If

#If WinActive("RStudio Server ahk_exe msedge.exe")
    ^+tab::send,^[
    ^tab::send,^]
    
    ^!/::Send,^{/}!+{Down}^{/} ;copy to next line and comment original line
    ^+/::Send,!{Up}^/{Down}^{/} ;switch comment with prev line

    ; workaround for the run current line cannot run help example bug
    !Enter::Send,^{Enter}
    !CapsLock::Send,^{Enter}

    ^Enter::
    ^CapsLock::Send,{End}{Enter}

    ^+Enter::
    ^+CapsLock::Send,{Home}{Enter}{Up}
#If

; lab? (3) - JupyterLab and 1 more page - Profile 0 - Microsoft​ Edge Dev
; #If WinActive("JupyterLab 和另外 ahk_exe msedge.exe",,":") ; finally
; #If WinActive(" - JupyterLab 和另外 ahk_exe msedge.exe") or WinActive("ahk_exe JupyterLab.exe") or WinActive("JupyterLab - Profile 0 - Microsoft​ Edge Dev ahk_exe msedge.exe") or WinActive("RStudio Server ahk_exe msedge.exe")
#If isJLabActive()
; finally

; --------------------------------------------------------------------------------------------------------------------------------
;;     Jupyterlab key remap
; --------------------------------------------------------------------------------------------------------------------------------
    ; remap default keybindings
    ^w::send,!w ; zhe
    ^+tab::send,^+[
    ^tab::send,^+]
    ^+z::send,^+z ; bu xing zai shuo ; jing ran zi dai le

    ; macro
    ~!+up::send,!+{down}{Up} ; duplicate line and cursor up
    ^!/::Send,^/!+{down}^/ ;copy to next line and comment original line
    ^+/::Send,!{Up}^/{Down}^{/} ;switch comment with prev line

    ; magnify cnm
    ^!f::Send,^!{f12} ; also backup before format...
    ^!up::send,^!{f6}
    ^!down::send,^!{f7} ; sigh wtf

    ; disable alt?
    ; Alt::return
        ; msgbox,wtf
    ;     return
        
    ; Alt up::
    ;     send,{alt up}
    ;     return
    ;     if (GetKeyState("CapsLock", "P"))
    ;         Send,{Alt Down}
        ; return ; but capslock + alt did not work then, also f3 to search conflict with sublime key ; 6
    ; Alt Up::
    ;     Send,{Alt Up}
    ;     return

    ; pipe
    ^+.::send,{space}`%>`%{space}
    ^.::send,{space}`%>`%{space}
    !.::send,{space}`%>`%{space}
    ^+,::send,{space}`%<>`%{space}
    ^!,::send,{space}`%<>`%{space}
    !,::send,{space}`%<>`%{space} ; in case need %>% but accidently input %<>% ... abandon! ; used too frequent...
    ^!.::send,{space}`%T>`%{space}

    !=::send,`%{+}`%
    !-::send,`%{c}`%
    !i::send,`%in`%
    ^+i::send,`%nin`%


    ^+e::Send,^+f
    ^+x::
    !e::
        ; getPixelColor!
        ; x:=120 ;
        x:=300/3840*A_ScreenWidth
        y:=780/2160*A_ScreenHeight
        ; y:=712
        if (A_ComputerName=="MBE") {
            x:=120
            y:=660
        }
        if (WinActive("JupyterLab and ahk_exe msedge.exe")) {
            x:=120/3840*A_ScreenWidth
            y:=600/2160*A_ScreenHeight
        }
        PixelGetColor, a, x, y
        ; msgbox, % a
        if (a=0x424242) {
            CoordMode, Mouse, Screen
            MouseGetPos, xpos, ypos
            CoordMode, Mouse, Window
            ; click,140,860 ; with fav bar
            ; click,40,600 ; full screen
            ; click,140,680 ; without fav bar
            ; click,300,780 ; without fav bar
            ; click,150,390 ; without fav bar 1080p
            click,%x%,%y%
            CoordMode, Mouse, Screen
            DllCall("SetCursorPos", int, xpos, int, ypos)
        }
        else ; if a=0x212121
            send,^+f
        return

; --------------------------------------------------------------------------------------------------------------------------------
;;     Jupyterlab hotstrings long
; --------------------------------------------------------------------------------------------------------------------------------

    :Co:imp::import sys, os, time, math, random, re, datetime
    :Co:imp1::
    Var =
    ( %
import sys, os, time, math, random, re
import matplotlib.pyplot as plt
from numpy import r_ as r
import numpy as np
import pandas as pd
    )
; import torch
    SendViaClip(Var)
    return


    :Co:ate::
        Var = 
        ( %
source('~/lab/test/libr.R')
# library('Signac')
        )
        SendViaClip(Var)
        Send,{Esc}q^{Enter}ba1+{Enter}aa ; initiate -> ate...
        return

;     :Co:ph::
;     Var =
;     ( %
; exec("from IPython.display import display, HTML; display(HTML('<style>.jp-CodeCell.jp-mod-outputsScrolled .jp-Cell-outputArea { max-height: 10em; }</style>'))")
; 1 %>% add(2)
;     )
;     SendViaClip(Var)
;     return

    ; :Co:wtf::
    ; Var = 
    ; ( %
    ; )
    ; SendViaClip(Var)
    ; return


; --------------------------------------------------------------------------------------------------------------------------------
;;     Jupyterlab hotstrings short
; --------------------------------------------------------------------------------------------------------------------------------


    :Co:hg1::hg19 = fread("~/lab/data.sync/features.hg19.tsv", header = F) %>% mutate_at("V2", make.unique)
    :Co:hg3::hg38 = fread("~/lab/data.sync/features.hg38.tsv", header = F) %>% mutate_at("V2", make.unique)
    :Co:asdfv::as.data.frame.vector
    :Co:asdft::as.data.frame.table
    ; :Co:asdfa::as.data.frame.array
    :Co:asdfr::as.data.frame.raw
    :C*:asdff::as.data.frame
    :Co:acd::active.ident

    :Co:pla::plan(multicore, workers = 4)
    ; :Co:pla2::plan(list(tweak(multicore, workers = 4), tweak(multicore, workers = 2)))

    :Co:nF::nFeature_RNA
    :Co:nC::nCount_RNA
    :Co:pe::percent.mt
    :Co:rsr::RNA_snn_res.0.8
    :Co:ssr::SCT_snn_res.0.8
    :Co:isr::integrated_snn_res.0.8
    :Co:psr::peaks_snn_res.0.8
    :Co:gb::group.by

    :Co:tt::{PgUp}tictoc`::tic()`n{PgDn}`ntictoc`::toc()
    :Co:ttt::{PgUp}start = time.time()`n{PgDn}`nprint(time.time() - start)
    :Co:rr::{PgDn}`nrp(){PgUp}rp(10,5)`n{Left 2}
    :Co:rr1::{PgDn}`nrmm(){PgUp}rmm(nr(df))`n{Left 2}
    :Co:rr2::{PgDn}`nrmm(){PgUp}rmm(dim(df))`n{Left 2}
    :Co:rr3::{PgDn}`nrpp(){PgUp}rpp(10000)`n{Left 2}

    :Co:futur::{Home}e=future({End},gc=T)
    :Co:valu::if (resolved(e)) a=value(e) else cat('NO'){Home}{Ctrl Down}{Right 7}{Ctrl Up} ; {Right} ; No and stop

    ; :Co:sv::saveRDS(){Left} ; no longer needed! use parallel implement!
    ; :Co:rd::readRDS(){Left}
    :Co:ii::%in%{Space}
    ; :Co:bu::box`::use(dplyr[...],purrr[...])

    :Co:ssec::map(a, ~map2_df(a, list(.x), ~len(sec(.x, .y)))) %>% do.call(rbind, .) ; use sec2 instead?

    ; :Co:oxrx::openxlsx`::read.xlsx(){Left}

    :Co:ccc::
        Send,^a!+i`"{end},^{End}{Backspace}^a+9^{Home}c{left}={left}
        return


    ; :Co:rmm::
    ;     SendViaClip()
    ;     return

#If ; WinActive


; --------------------------------------------------------------------------------------------------------------------------------
;;   2.2 More
; --------------------------------------------------------------------------------------------------------------------------------


#IfWinActive,ahk_exe code.exe
    ; ^!up::send,^+!{up}
    ; ^!down::send,^+!{down} ; sigh wtf
    ; magnify cnm
    ^!f::Send,^!{F12} ; also backup before format...
    ^!up::send,^!{f6} ; wtf mstsc not work
    ^!down::send,^!{f7} ; sigh wtf

    ; XButton2 & WheelDown::WheelRight
    ; XButton2 & WheelUp::WheelLeft

#IfWinActive



#IfWinActive,ahk_exe MendeleyDesktop.exe
    !w::Send,^w
    !2::send,^+{tab}
    !3::send,^{tab}

#IfWinActive

; #IfWinActive,ahk_exe mstsc.exe ; emm sha shi hou fang shang lai de
;     CapsLock::CapsLock
; #IfWinActive

#IfWinActive,ahk_exe EXCEL.EXE
    ^+tab::send,^{PgUp}
    ^tab::send,^{PgDn}
    ; WheelLeft::msgbox,wtf
    ; WheelRight::Send,{WheelRight}
    +WheelUp::Send,{WheelLeft}
    +WheelDown::Send,{WheelRight}
    ; XButton2 & WheelUp::Send,{WheelLeft}
    ; XButton2 & WheelDown::Send,{WheelRight}
#IfWinActive

#IfWinActive,ahk_exe PotPlayerMini64.exe
    WheelLeft::Send,{PgUp}
    WheelRight::Send,{PgDn}
#IfWinActive


#IfWinActive,ahk_exe DB Browser for SQLite.exe
    +WheelUp::Send,{WheelLeft}
    +WheelDown::Send,{WheelRight}
    XButton2 & WheelUp::Send,{WheelLeft}
    XButton2 & WheelDown::Send,{WheelRight}
#IfWinActive

; #IfWinActive,ahk_exe RemotePlay.exe
;     rbutton::Send,^!h

; #IfWinActive
#IfWinActive,ahk_exe steam.exe
    !a::Send,!{Left}
    !g::Send,!{Right}
#IfWinActive



#IfWinActive,ahk_exe eudic.exe
    !a::Send,^{Left}
    !g::Send,^{Right}
#IfWinActive
; #If A_ComputerName=="DESKTOP-1I7827D"
;     F1::msgbox,test
; #If


#IfWinActive,ahk_exe taskmgr.exe
    +WheelUp::Send,{WheelLeft 2}
    +WheelDown::Send,{WheelRight 2}
    !1::send,^+{tab}
    !2::send,^+{tab}
    !3::send,^{tab}
    !4::send,^{tab}

#IfWinActive


#IfWinActive,ahk_exe Everything.exe
    !f::Send,!d!{down} ; this (all in one place) is easier...
    !w::Send,^w
#IfWinActive


#IfWinActive,ahk_exe WeChatAppEx.exe
    !a::send,!{left}
    !g::send,!{right}
    !w::send,^w
    
    !1::send,^1
    !2::send,^{PgUp}
    !3::send,^{PgDn}
    !4::send,^9
    !5::send,^5

#IfWinActive


#IfWinActive,ahk_exe WsaClient.exe
    Esc::Send,!{left}
    ; Backspace::Send,!{left}
#IfWinActive


#IfWinActive,ahk_exe Illustrator.exe
    ^+z::send,^+z
#IfWinActive


; --------------------------------------------------------------------------------------------------------------------------------
;;   2.3 prevent conflict
; --------------------------------------------------------------------------------------------------------------------------------
^!w::
    If WinExist("微信 ahk_class com.tencent.mm ahk_exe WsaClient.exe") {
        If WinActive("微信 ahk_class com.tencent.mm ahk_exe WsaClient.exe") {
            WinMinimize,A
        } 
        else {
            WinActivate
        }
    }
    else {
        send,^!w
        ; msgbox,wtf2
        ; msgbox,wtf3
    }
    return


; #IfWinActive,ahk_exe
!r::
    a:=A_ScreenWidth
    b:=A_ScreenHeight
    ra:=a/3840
    rb:=b/2160
    ; msgbox,% A_ScreenHeight
    If WinActive("eusoft_eudic_en_win32 ahk_exe eudic.exe") 
        winmove, A, , , , 1200*ra, 1200*rb
    else If WinActive("ahk_exe eudic.exe")
        winmove,A, , (1920-2400/2)*ra, (1080-1600/2)*rb, 2400*ra, 1600*rb
    else If WinActive("ahk_exe taskmgr.exe")
        winmove,A, , (1920-2100/2)*ra, (1080-1400/2)*rb, 2100*ra, 1400*rb
    else If WinActive("ahk_class CefWebViewWnd")  or WinActive("ahk_exe WeChatAppEx.exe") ; wechat web
        winmove, A, , (1920-900-200)*ra, (1080-900)*rb, 1800*ra, 1800*rb
    else If WinActive("ahk_class H5SubscriptionProfileWnd")
        ; winmove, A, , (1920-900-64)*ra, (1080-900)*rb, 1100*ra, 1600*rb
        winmove, A, , (1920-900-64)*ra, (1080-900)*rb, , 1800*rb
    else If WinActive("ahk_class CWebviewControlHostWnd") or WinActive("ahk_class WeChatMainWndForPC") 
        winmove, ahk_class WeChatMainWndForPC, , (1920-2143/2)*ra, (1080-1400/2)*rb, 2144*ra, 1400*rb
    else If WinActive("ahk_class SubscriptionWnd") or  WinActive("ahk_exe GitHubDesktop.exe") 
        winmove, A, , (1920-2912/2+100)*ra, (1080-1800/2)*rb, 2912*ra, 1800*rb
    else If WinActive("ahk_exe foobar2000.exe")
        winmove, A, , (1920-3000/2)*ra, (1080-1800/2)*rb, 3000*ra, 1800*rb
    else
        send,!r
    return
; #IfWinActive


; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;; 3 Labels
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------


cnm1:
    ; hotkey,If, WinActive("JupyterLab 和另外 ahk_exe msedge.exe") ; finally
    ; hotkey,Alt,off
    ; hotkey,If
    SetKeyDelay -1   ; If the destination key is a mouse button, SetMouseDelay is used instead.
    Send {Blind}{Enter DownR}  ; DownR is like Down except that other Send commands in the script won't assume "b" should stay down during their Send.
    return

cnm2:
    ; hotkey,If, WinActive("JupyterLab 和另外 ahk_exe msedge.exe") ; finally
    ; hotkey,Alt,on
    ; hotkey,If
    SetKeyDelay -1  ; See note below for why press-duration is not specified with either of these SetKeyDelays.
    Send {Blind}{Enter up}

    return

ProxyChecker:
    IESET:="Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    RegRead, OutputVar, HKCU\%IESET%, AutoConfigURL
    if (OutputVar!="") {
        ; MsgBox, %OutputVar%
        RegDelete,HKCU\%IESET%,AutoConfigURL
        RegWrite,REG_DWORD,HKCU\%IESET%,Proxyenable,1
    }
    return

PlaytimeTracker:
    ; IfWinActive, mon Brilliant Diamond (64-bit) | 1.1.1 | NVIDIA
    ; IfWinActive, ahk_exe Minecraft.exe
    ; {
    ;     game := "Minecraft"
    ;     FormatTime, time,, yyMMdd-HHmmss
    ;     text := Format("{1}`n",time)
    ;     filename := Format("E:\test\log\{1}.log", game)
    ;     FileAppend, %text%, %filename%
    ; }
    for exe, game in {"Minecraft ahk_exe javaw.exe":"Minecraft"
                    ,"ahk_exe Doki Doki Literature Club Plus.exe":"DDLC"
                    ; ,"":""
                    ,"ahk_exe Genshin Impact Cloud Game.exe":"YuanShen" ; not that necessary...
                    ,"ahk_exe YuanShen.exe":"YuanShen"} {
        if WinActive(exe) {
            FormatTime, time,, yyMMdd-HHmmss
            text := Format("{1}`n",time)
            filename := Format("E:\test\log\{1}.log", game)
            FileAppend, %text%, %filename%
        }
    }
    return


FlaskChecker:
    ; if WinActive("ahk_exe mstsc.exe") 
        ; WinMinimize, A
    ; if GetKeyState("CapsLock", "T")
    ;     SetCapsLockState, Off
    ; SetCapsLockState, AlwaysOff
        

    global RemoteControl
    SysGet RemoteControl, 4096 ; SM_REMOTESESSION ; 8193 SM_REMOTECONTROL: This system metric is used in a Terminal Services environment. Its value is nonzero if the current session is remotely controlled; zero otherwise.
    ; game := Format("rdp_{1}", A_ComputerName)
    ; FormatTime, time,, yyMMdd-HHmmss
    ; text := Format("{1} {2}`n",time,RemoteControl)
    ; filename := Format("E:\test\log\{1}.log", game)
    ; FileAppend, %text%, %filename%
    
    if (HideMagifier=="T") {
        If WinExist("ahk_exe Magnify.exe")
            WinHide, ahk_exe Magnify.exe
    }
    
    
    ; if (A_ComputerName!="DESKTOP-EAJ7U03")
    ; if (A_ComputerName!="LAB") ; more than one month no attachment...
        return
    ; global prevPid
    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Option(4) := 13056
    whr.Open("GET", "https://127.0.0.1:50000?test=test", true)
    whr.Send()
    try
        a := whr.WaitForResponse(2)
    catch e
        a := "wtf"
    if (a != -1) {
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        whr.Option(4) := 13056
        whr.Open("GET", "https://127.0.0.1:50000?test=test", true)
        whr.Send()
        try
            a := whr.WaitForResponse(3)
        catch e
            a := "wtf"
        if (a != -1) {
            ; msgbox,wtf!!!
            WinGet, idwtf, List, D:\miniconda3\envs\dcs\python.exe
            Loop, %idwtf%
            {
                this_id := idwtf%A_Index%
                ; WinClose, ahk_id %this_id%
                WinKill, ahk_id %this_id%
            }
            Run,D:\miniconda3\envs\dcs\python.exe E:\lab\dcs\api.py ; , , Min
            game := "wtf"
            FormatTime, time,, yyMMdd-HHmmss
            text := Format("{1}`n",time)
            filename := Format("E:\lab\log\{1}.log", game)
            FileAppend, %text%, %filename%
        }
    }
    return

wtf:
    msgbox, %A_ComputerName% %RemoteControl% 
    return

wtf2:
    msgbox, %A_ComputerName% %RemoteControl% no capsloack ; and WheelLeft/Right hotkey" ; wtf client ahk still lag???
    return

; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
;; 4 auto-correction
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; --------------------------------------------------------------------------------------------------------------------------------
; C:case sensitive
; *: An ending character (e.g. Space, ., or Enter) is not required to trigger the hotstring
:*:foramt::format
:C*:mamab::mamba
:*:pyhton::python
:C*:codna::conda
:C*:repalce::replace
:C*:fodler::folder
:C*:laod::load
:*:~?::~/


; F1::
;     MouseMove,1920,1480
;     send,{LButton down}
;     MouseMove,1280,900
;     MouseMove,1280,800
;     MouseMove,1290,800
;     MouseMove,1290,900
;     send,{LButton up}
;     return