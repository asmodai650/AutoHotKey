loop
{

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 100
ClipBoard =

Send, {CTRLDOWN}c{CTRLUP}
ClipWait
StringReplace, ClipBoard, ClipBoard, `r`n, ,All

If ClipBoard =
{
Msgbox Macro is complete.
ExitApp
;Pause
}



Sleep, 500
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Sleep, 500
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search,

Sleep, 500
Send, {CTRLDOWN}v{CTRLUP}{ENTER}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Sleep, 500
Send, {TAB}{TAB}{ENTER}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Sleep, 5000

Send, {ALTDOWN}{ALTUP}

Sleep, 1000
send, v
Sleep, 100
send, D

Sleep, 1000


Send, {TAB}
Send, I
Send, {TAB}

Sleep, 2000


Send, {ALTDOWN}{ALTUP}
send, F
send, V
Sleep, 5000

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 500
Send, {TAB}x{DOWN}{LEFT}{LEFT}
Sleep, 500


}

Esc::pause

Return