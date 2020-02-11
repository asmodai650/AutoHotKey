loop
{

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 200
ClipBoard =

Send, {CTRLDOWN}c{CTRLUP}
ClipWait
StringReplace, ClipBoard, ClipBoard, `r`n, ,All

If ClipBoard =
{
Msgbox Macro is complete.
Pause
}



Sleep, 3000
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search,

Send, {CTRLDOWN}v{CTRLUP}{ENTER}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, {TAB}{TAB}{ENTER}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Sleep, 5000

Send, {ALTDOWN}{ALTUP}

Sleep, 2000
send, v
Sleep, 200
send, D

Sleep, 2000


Send, {TAB}
Send, PPN
Send, {TAB}

Sleep, 2000


Send, {ALTDOWN}{ALTUP}
send, F
send, V
Sleep, 4000

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 200
Send, {TAB}x{DOWN}{LEFT}{LEFT}
Sleep, 200


}
Esc::pause

Return