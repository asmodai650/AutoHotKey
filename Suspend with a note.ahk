loop,
{
WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

Sleep, 300
Send, {CTRLDOWN}c{CTRLUP}
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Sleep, 500
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}{TAB}{ENTER}

WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Send, {CTRLDOWN}r{CTRLUP}
WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  , 
Send, {CTRLDOWN}{ALTDOWN}es{CTRLUP}{ALTUP}
WinWait, Suspense Request Tracking - Claim # , 
IfWinNotActive, Suspense Request Tracking - Claim # , , WinActivate, Suspense Request Tracking - Claim # , 
WinWaitActive, Suspense Request Tracking - Claim # , 
Send, op{TAB}mELISSA RUSH{TAB}Ret{TAB}
send,Approved for offset attempt.{CTRLDOWN}s{CTRLUP}

WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  , 
;MouseClick, left,  870,  14
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 600
WinWait, Patient Claim , 
	IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
	WinWaitActive, Patient Claim ,



WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

Sleep, 600
Send, x{ENTER}{LEFT 2}


}

Esc::Pause

Return

