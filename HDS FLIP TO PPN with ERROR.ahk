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
Pause
}


Sleep, 5000
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

sleep, 2000

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
Send, PPN
Send, {TAB}

;;;;;;;; BEGIN new code to close window if claim cannot be flipped ;;;;;;;;
Sleep, 500

Loop,
	{
		Sleep, 1000
		IfWinExist, Invalid Method Transition
		{
			;msgbox, , , Window Found, 2
			WinActivate, Invalid Method Transition ,
			WinWaitActive, Invalid Method Transition ,
			Sleep, 1000
			Send, {Esc}
			Sleep, 1000
			IfWinNotActive, Patient Claim Detail , , WinActivate, Patient Claim Detail , 
			WinWaitActive, Patient Claim Detail ,
			Sleep, 1000
			Send, {Esc}
			Sleep, 1000 
			Comment = Cannot Flip to PPN
			GoTo, EXCEL
		}
		else
			break
	}
;;;;;;;; END new code to close window if claim cannot be flipped ;;;;;;;;


Sleep, 2000


Send, {ALTDOWN}{ALTUP}
send, F
send, V
Sleep, 2000
Comment = Flipped to PPN
GoTo, EXCEL


EXCEL:
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 100
Send, {TAB}
Sleep, 200
Send, %Comment%
Sleep, 200
Send, {DOWN}{LEFT}{LEFT}
Sleep, 1000

}
Pause::pause

Return