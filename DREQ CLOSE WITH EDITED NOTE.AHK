MsgBox, 
(
Last Updated: 	02/23/11

Description: 	This macro denies pre-invoice active claims. The reason MUST be Contractual > Refund Written Incorrectly.
		The note entered on the denial is "CLOSED IN ODAR-Per TRN file, see FACS for exact TRN reason.".

Requirements: 	This macro requires an excel spreadsheet labeled “NO LONGER PPN.xls”. It requires the AIM Claim ID in
		Column A and the denial amount in Column B. Currently an adjustment would be needed depending on
		the denial reason needed and the note needed.

Starting Inst:	Make sure that Safari is running but all claims are closed. Safari should be maximized on one
		screen and Excel should be maximized on your other screen. The spreadsheet should be called
		"No longer PPN.xls".

Time per claim:	TBD
)



WinWait, Microsoft Excel - NO LONGER PPN, 
IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
WinWaitActive, Microsoft Excel - NO LONGER PPN, 
Sleep, 500
Send, {RIGHT 2}
Sleep, 100
Send, =IF(B1<100,"1",IF(B1<1000,"2",IF(B1<3000,"3",IF(B1<5000,"4",IF(B1>=5000,"5"))))){ENTER}
Sleep, 300
Send, {UP}
Sleep, 100

Send, {CTRLDOWN}c{CTRLUP}{CTRLDOWN}{SHIFTDOWN}{DOWN}{CTRLUP}{SHIFTUP}{CTRLDOWN}v{CTRLUP}

Sleep, 100

Send, {CTRLDOWN}c{CTRLUP}


Sleep, 300
Send, {ALTDOWN}{ALTUP}hv

Send, v


Send, {LEFT 2}


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

WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Sleep, 200
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
;Send, {CTRLDOWN}v{CTRLUP}{ENTER}{TAB}{ENTER}
Send, {CTRLDOWN}v{CTRLUP}{ENTER}{TAB}{DOWN}{TAB}{ENTER}


WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Send, {CTRLDOWN}r{CTRLUP}
WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  , 








		WinWait, Microsoft Excel - NO LONGER PPN, 
		IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
		WinWaitActive, Microsoft Excel - NO LONGER PPN, 


			Send, {Right 2}{CTRLDOWN}c{CTRLUP}
			Sleep, 300
			StringReplace, clipboard, clipboard, `r`n, , All


			WinWait, Claim Requests for Claim ID:  , 
			IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
			WinWaitActive, Claim Requests for Claim ID:  , 


			Sleep, 3000


			If (1) = clipboard 
			{
			WinWait, Claim Requests for Claim ID:  , 
			IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
			WinWaitActive, Claim Requests for Claim ID:  , 
			Sleep, 100
			Send, {F5}
			Sleep, 500
			MouseClick, left,  187,  117
			Sleep, 300
			;Send, {CTRLDOWN}{HOME}{CTRLUP}
			Send, {DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}
			Send, {ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, A/R Level 1 Approval, 
			IfWinNotActive, A/R Level 1 Approval, , WinActivate, A/R Level 1 Approval, 
			WinWaitActive, A/R Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}RES-Small balance under $250.00 not worked by ORR.{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Send, {ALTDOWN}{ALTUP}
			Sleep, 100
			Send, f
			Sleep, 100
			Send, V
			Sleep, 2000
			}






			If (2) = clipboard
			{
			WinWait, Claim Requests for Claim ID:  , 
			IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
			WinWaitActive, Claim Requests for Claim ID:  , 
			Sleep, 100
			Send, {F5}
			Sleep, 500
			MouseClick, left,  187,  117
			Sleep, 300
			Send, {CTRLDOWN}{HOME}{CTRLUP}
			Send, {ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, A/R Level 1 Approval, 
			IfWinNotActive, A/R Level 1 Approval, , WinActivate, A/R Level 1 Approval, 
			WinWaitActive, A/R Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}RES-Small balance under $250.00 not worked by ORR.{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN}
			Sleep, 100
			Send, {ENTER}
			Sleep, 500
			WinWait, A/R Level 2 Approval, 
			IfWinNotActive, A/R Level 2 Approval, , WinActivate, A/R Level 2 Approval, 
			WinWaitActive, A/R Level 2 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Send, {ALTDOWN}{ALTUP}
			Sleep, 100
			Send, f
			Sleep, 100
			Send, V
			Sleep, 2000
			}


			If (3) = clipboard
			{
			WinWait, Claim Requests for Claim ID:  , 
			IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
			WinWaitActive, Claim Requests for Claim ID:  , 
			Sleep, 100
			Send, {F5}
			Sleep, 500
			MouseClick, left,  187,  117
			Sleep, 300
			Send, {CTRLDOWN}{HOME}{CTRLUP}
			Send, {ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, A/R Level 1 Approval, 
			IfWinNotActive, A/R Level 1 Approval, , WinActivate, A/R Level 1 			Approval, 
			WinWaitActive, A/R Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}RES-Small balance under $250.00 not worked by ORR.{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial 			Request #, 
			WinWaitActive, Claim Denial Request #, 
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN}
			Sleep, 100
			Send, {ENTER}
			Sleep, 500
			WinWait, Ops Field Office Level 1 Approval, 
			IfWinNotActive, Ops Field Office Level 1 Approval, , WinActivate, Ops 			Field Office Level 1 Approval, 
			WinWaitActive, Ops Field Office Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial 			Request #, 
			WinWaitActive, Claim Denial Request #, 
			Send, {ALTDOWN}{ALTUP}
			Sleep, 100
			Send, f
			Sleep, 100
			Send, V
			Sleep, 2000
			}





			If (4) = clipboard
			{
			WinWait, Claim Requests for Claim ID:  , 
			IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
			WinWaitActive, Claim Requests for Claim ID:  , 
			Sleep, 100
			Send, {F5}
			Sleep, 500
			MouseClick, left,  187,  117
			Sleep, 300
			Send, {CTRLDOWN}{HOME}{CTRLUP}
			Send, {ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, A/R Level 1 Approval, 
			IfWinNotActive, A/R Level 1 Approval, , WinActivate, A/R Level 1 Approval, 
			WinWaitActive, A/R Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}RES-Small balance under $250.00 not worked by ORR.{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN}
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, Ops Field Office Level 1 Approval, 
			IfWinNotActive, Ops Field Office Level 1 Approval, , WinActivate, Ops Field Office Level 1 Approval, 
			WinWaitActive, Ops Field Office Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN 2}
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, Ops Field Office Level 2 Approval, 
			IfWinNotActive, Ops Field Office Level 2 Approval, , WinActivate, Ops Field Office Level 2 Approval, 
			WinWaitActive, Ops Field Office Level 2 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Send, {ALTDOWN}{ALTUP}
			Sleep, 100
			Send, f
			Sleep, 100
			Send, V
			Sleep, 2000
			}




			If (5) = clipboard
			{
			WinWait, Claim Requests for Claim ID:  , 
			IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
			WinWaitActive, Claim Requests for Claim ID:  , 
			Sleep, 100
			Send, {F5}
			Sleep, 500
			MouseClick, left,  187,  117
			Sleep, 300
			Send, {CTRLDOWN}{HOME}{CTRLUP}
			Send, {ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, A/R Level 1 Approval, 
			IfWinNotActive, A/R Level 1 Approval, , WinActivate, A/R Level 1 Approval, 
			WinWaitActive, A/R Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}RES-Small balance under $250.00 not worked by ORR.{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN}
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, Ops Field Office Level 1 Approval, 
			IfWinNotActive, Ops Field Office Level 1 Approval, , WinActivate, Ops Field Office Level 1 Approval, 
			WinWaitActive, Ops Field Office Level 1 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN 2}
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, Ops Field Office Level 2 Approval, 
			IfWinNotActive, Ops Field Office Level 2 Approval, , WinActivate, Ops Field Office Level 2 Approval, 
			WinWaitActive, Ops Field Office Level 2 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 			
			Sleep, 500
			Send, {CTRLDOWN}s{CTRLUP}
			Sleep, 500
			Send, {F5}
			Sleep, 500
			MouseClick, left,  178,  330
			Sleep, 100
			Send, {DOWN 3}
			Sleep, 100
			Send, {ENTER}
			Sleep, 100
			WinWait, A/R Level 3 Approval, 
			IfWinNotActive, A/R Level 3 Approval, , WinActivate, A/R Level 3 Approval, 
			WinWaitActive, A/R Level 3 Approval, 
			Sleep, 500
			Send, a{TAB}{TAB}{ENTER}
			WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #, 
			Send, {ALTDOWN}{ALTUP}
			Sleep, 100
			Send, f
			Sleep, 100
			Send, V
			Sleep, 2000
			}




	WinWait, Claim Requests for Claim ID:  , 
	IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
	WinWaitActive, Claim Requests for Claim ID:  , 
	Sleep, 300
	Send, {ALTDOWN}{ALTUP}
	Sleep, 100
	Send, f
	Sleep, 100
	Send, C
	Sleep, 2000
	
	WinWait, Patient Claim , 
	IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
	WinWaitActive, Patient Claim , 


	WinWait, Microsoft Excel - NO LONGER PPN, 
	IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
	WinWaitActive, Microsoft Excel - NO LONGER PPN, 
	Sleep, 100
	Send, x{ENTER}{LEFT 2}
}

Esc::pause

Return