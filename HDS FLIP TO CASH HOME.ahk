loop
{

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel,
WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 100
ClipBoard =

Send, {CTRLDOWN}c{CTRLUP}
ClipWait
StringReplace, ClipBoard, ClipBoard, `r`n, ,All

If ClipBoard =
{
Msgbox Macro is complete.
Pause
}



Sleep, 500
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

Sleep, 5000
send, v
Sleep, 100
send, D

Sleep, 2000

Send, {TAB}
Send, {TAB}
Send, {TAB}
Send, COLL
Send, {TAB}

Sleep, 2000


Send, {ALTDOWN}{ALTUP}
send, F
send, V
Sleep, 2000

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel,
WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 100
Send, {TAB}x{DOWN}{LEFT}{LEFT}
Sleep, 100


}
Esc::pause

Return