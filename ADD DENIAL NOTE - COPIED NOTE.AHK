
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

Sleep, 1000

;Send, {ALTDOWN}{ALTUP}en
Send, {ALTDOWN}
Send, {ALTUP}
Send, en

Sleep, 1000

WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, 

Send, {CTRLDOWN}n{CTRLUP}

Sleep, 10000

Send, {TAB}{TAB}{TAB}
Send, PP
Send, {TAB}{TAB}{TAB}


WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,


Send, {TAB}
Send, {CTRLDOWN}c{CTRLUP}

WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes,

Send, {CTRLDOWN}v{CTRLUP}


MouseClick, left,  539,  14
Sleep, 100
Send, {ENTER}
Sleep, 100

Sleep, 600
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

Sleep, 600
Send, {RIGHT}x{ENTER}{LEFT}

}

Esc::PAUSE

Return