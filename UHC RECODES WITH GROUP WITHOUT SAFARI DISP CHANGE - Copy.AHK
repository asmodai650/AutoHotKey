loop
{

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

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



Sleep, 100
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
Send, {TAB}{DOWN}{DOWN}{ENTER}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Sleep, 5000

Send, {ALTDOWN}{ALTUP}
;Send, {ALTUP}
Sleep, 1000
send, v
Sleep, 100
send, E




Sleep, 100
SEnd, {ALTDOWN}{ALTUP}
Sleep, 100
send, F
Sleep, 100
send, N
Sleep, 100
Send, {TAB}{TAB}
Sleep, 100
Send, 400000208
Sleep, 100
Send, {ENTER}
Sleep, 100
Send, {ENTER}
Sleep, 100
Send, {ENTER}
Sleep, 1000
Send, {TAB}{TAB}{TAB}{TAB}
Sleep, 1000

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

;GROUP#
Sleep, 100
Send, {RIGHT}{CTRLDOWN}c{CTRLUP}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}

Sleep, 100
MouseClick, left,  694,  267
MouseClick, left,  694,  267
Sleep, 100

Sleep, 2000
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

;DIRECTORY
Sleep, 100
Send, {RIGHT}{CTRLDOWN}c{CTRLUP}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}


Sleep, 2000
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,

;FACS#
Sleep, 100
Send, {RIGHT}{CTRLDOWN}c{CTRLUP}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
MouseClick, right,  578,  412
Send, {DOWN}{DOWN}{DOWN}{DOWN}{ENTER}{ENTER}



Sleep, 2000
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,

;PLATFORM
Sleep, 100
Send, {RIGHT}{CTRLDOWN}c{CTRLUP}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}{ENTER}{CTRLDOWN}s
Sleep, 2000

Send, {ALTDOWN}{ALTUP}
;SEnd, {ALTUP}
send, v
send, u

Sleep, 2000
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,

;UID#
Sleep, 100
Send, {RIGHT}{CTRLDOWN}c{CTRLUP}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}
Sleep, 2000

Send, {ALTDOWN}q{ALTUP}
;SEnd, {ALTUP}
;send, Q



Sleep, 2000
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,
 

Sleep, 100
;Send, {TAB}x{DOWN}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}
Send, {TAB}x{DOWN}{HOME}
Sleep, 100


}
Esc::pause

Return