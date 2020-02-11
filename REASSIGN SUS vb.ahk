MsgBox, 
(
Last Updated: 	02/23/11

Description: 	This macro reassigns suspense to RAM.

Requirements: 	This macro requires an excel spreadsheet labeled “NO LONGER PPN.xls”. It requires the AIM Claim ID in
		Column A, RAm name in column B and note in column C. 

Starting Inst:	 The spreadsheet should be called "No longer PPN.xls". Start safari with Suspense Request window open.

Time per claim:	TBD
)





Loop
{

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 

Sleep, 200
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

WinWait, Suspense Request , 
IfWinNotActive, Suspense Request , , WinActivate, Suspense Request , 
WinWaitActive, Suspense Request , 

Sleep, 200
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, {CTRLDOWN}v{CTRLUP}
Send, {ENTER}
Sleep, 200
;Send, {TAB}{END}
;Sleep, 500
MouseMove, 346, 171 
MouseClick, left,  346,  171
Send, {END}
Sleep, 200
Send, {ENTER}



WinWait, Suspense Request , 
IfWinNotActive, Suspense Request , , WinActivate, Suspense Request , 
WinWaitActive, Suspense Request , 


MouseMove, 410, 171 
MouseClick, right,  410,  171
Sleep, 100
Send, {DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}
WinWait, Suspense Request Tracking, 
IfWinNotActive, Suspense Request Tracking, , WinActivate, Suspense Request Tracking, 
WinWaitActive, Suspense Request Tracking, 
Send, o
send, {tab}
sleep, 1000
send, {shift}{home}

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,

Send, {TAB}
Send, {CTRLDOWN}c{CTRLUP}

WinWait, Suspense Request Tracking, 
IfWinNotActive, Suspense Request Tracking, , WinActivate, Suspense Request Tracking, 
WinWaitActive, Suspense Request Tracking, 

Send, {CTRLDOWN}v{CTRLUP}
send, {tab}
send, ram-other
send, {tab}

WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN,

Send, {TAB}
Send, {CTRLDOWN}c{CTRLUP}

WinWait, Suspense Request Tracking, 
IfWinNotActive, Suspense Request Tracking, , WinActivate, Suspense Request Tracking, 
WinWaitActive, Suspense Request Tracking, 

Send, {CTRLDOWN}v{CTRLUP}


Send, {ALTDOWN}{ALTUP}
Sleep, 100
Send, f
Sleep, 100
Send, V
Sleep, 5000

WinWait, Suspense Request, 
IfWinNotActive, Suspense Request, , WinActivate, Suspense Request, 
WinWaitActive, Suspense Request, 



sleep, 600
send, {f5}
sleep, 600
send, {f5}

	

	WinWait, Microsoft Excel - NO LONGER PPN, 
	IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
	WinWaitActive, Microsoft Excel - NO LONGER PPN, 
	Sleep, 100
	Send, x{ENTER}{HOME}
}

Esc::Pause

Return