
Loop
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

Sleep, 300
WinWait, Provider AR Summary for Claim , 
IfWinNotActive, Provider AR Summary for Claim , , WinActivate, Provider AR Summary for , 
WinWaitActive, Provider AR Summary for Claim , 
Send, {CTRLDOWN}fv{CTRLUP}{ENTER}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, {TAB}{down}{TAB}{ENTER}
WinWait, Provider AR Summary for Claim , 
IfWinNotActive, Provider AR Summary for Claim , , WinActivate, Provider AR Summary for Claim , 
WinWaitActive, Provider AR Summary for Claim , 


Send, {TAB}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ALTDOWN}{ALTUP}t{DOWN}{DOWN}{DOWN}{ENTER}



;MouseClick, right,  369,  367
;Sleep, 100

Sleep, 1000
WinWait, Correction, 
IfWinNotActive, Correction, , WinActivate, Correction, 
WinWaitActive, Correction, 
;MouseClick, left,  445,  510
MouseClick, left,  57,  65

Sleep, 500

WinWait, Transaction Reasons, 
IfWinNotActive, Transaction Reasons, , WinActivate, Transaction Reasons, 
WinWaitActive, Transaction Reasons, 

Send, {TAB}{TAB}

send, {DOWN}

send, {RIGHT}

send, {DOWN}

Send, {TAB}

send, {ENTER}

WinWait, Correction, 
IfWinNotActive, Correction, , WinActivate, Correction, 
WinWaitActive, Correction, 

Send, {TAB}

send, {ENTER}



Sleep, 600
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

Sleep, 600
Send, {RIGHT}x{ENTER}{LEFT}

}

Esc::PAUSE

Return