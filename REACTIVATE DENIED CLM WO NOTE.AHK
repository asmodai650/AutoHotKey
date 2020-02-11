
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
Send, {TAB}{TAB}{ENTER}
WinWait, Provider AR Summary for Claim , 
IfWinNotActive, Provider AR Summary for Claim , , WinActivate, Provider AR Summary for Claim , 
WinWaitActive, Provider AR Summary for Claim , 


Send, {TAB}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ALTDOWN}{ALTUP}t{DOWN}{DOWN}{DOWN}{ENTER}



;MouseClick, right,  369,  367
;Sleep, 100
;MouseClick, left,  388,  421
;Sleep, 100
WinWait, Correction, 
IfWinNotActive, Correction, , WinActivate, Correction, 
WinWaitActive, Correction, 
Send, {ALTDOWN}p{ALTUP}

Sleep, 600
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

Sleep, 600
Send, {RIGHT}x{ENTER}{LEFT}

}

Esc::PAUSE

Return