MsgBox, 
(
Last Updated: 	02/23/11

Description: 	This resolves suspense claims with a specific note back to RAM. The reason MUST be Close Pass QA. 

Requirements: 	This macro requires an excel spreadsheet labeled “NO LONGER PPN.xls”. It requires the AIM Claim ID in
		Column A and the denial amount in Column B. 

Starting Inst:	Make sure that Safari is running but all claims are closed. Safari should be maximized on one
		screen and Excel should be maximized on your other screen. The spreadsheet should be called
		"No longer PPN.xls".

Time per claim:	TBD
)






Loop
{

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

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

Sleep, 100
WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Sleep, 200
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}{TAB}{DOWN}{ENTER}



WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Send, {CTRLDOWN}r{CTRLUP}
WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  , 
Sleep, 100

Send, {CTRLDOWN}{END}{CTRLUP}{ENTER}
Sleep, 300

WinWait, Suspense Request #, 
IfWinNotActive, Suspense Request #, , WinActivate, Suspense Request #, 
WinWaitActive, Suspense Request #, 
MouseClick, right,  364,  168
Sleep, 600
Send, {DOWN}{DOWN}{ENTER}
Sleep, 600

WinWait, Suspense Request Tracking - Claim # , 
IfWinNotActive, Suspense Request Tracking - Claim # , , WinActivate, Suspense Request Tracking - Claim # , 
WinWaitActive, Suspense Request Tracking - Claim # , 
Send, This refund has aged over 200 with no retraction. Can a cash payment be obtained?

sleep, 600

send, {ALTDOWN}{ALTUP}
Sleep, 100
send, F
Sleep, 100
Send, V
Sleep, 100
send, {Enter}
sleep, 8000

WinWait, Suspense Request #, 
IfWinNotActive, Suspense Request #, , WinActivate, Suspense Request #, 
WinWaitActive, Suspense Request #, 
Sleep, 600

send, {ALTDOWN}{ALTUP}
Sleep, 100
send, F
Sleep, 100
Send, C
Sleep, 3000

WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  ,

send, {ALTDOWN}{ALTUP}
Sleep, 100
send, F
Sleep, 100
Send, C
Sleep, 3000

WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 
Sleep, 100
Send, {RIGHT}x{DOWN}{LEFT}

}

Esc::Pause

Return