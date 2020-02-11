MsgBox, 
(
Last Updated: 	11/07/11

Description: 	This macro disapproves denail requests and reactivates claims back to RAM.		

Requirements: 	This macro requires an excel spreadsheet labeled “NO LONGER PPN.xls”. It requires the AIM Claim ID in Column A.

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

WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Sleep, 200
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}{TAB}{DOWN}{DOWN}{ENTER}

WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 

Send, {CTRLDOWN}r{CTRLUP}

WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  ,

Sleep, 600
Send, {CTRLDOWN}{END}{CTRLUP}{ENTER}

WinWait, Claim Denial Request #, 
			IfWinNotActive, Claim Denial Request #, , WinActivate, Claim Denial Request #, 
			WinWaitActive, Claim Denial Request #,

Sleep, 600

MouseClick, right,  105,  327

Sleep, 600
Send, {ENTER}d{TAB}Recoding.
.{TAB}{ENTER}

Send, {ALTDOWN}{ALTUP}
Sleep, 100
send, E
Sleep, 100
Send, C
Sleep, 100
send, {ENTER}
Sleep, 2000

WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  ,
Sleep, 600
send, {ALTDOWN}{ALTUP}
Sleep, 100
send, F
Sleep, 100
Send, C
Sleep, 2000

WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim ,

;Sleep, 600
;send, {F6}
;Sleep, 200
;send, {F2}
;Sleep, 600
;WinWait, Notes, 
;IfWinNotActive, Notes, , WinActivate, Notes, 
;WinWaitActive, Notes, 
;Sleep, 600


;Send, {CTRLDOWN}n{CTRLUP}
;Sleep, 10000

;Send, {TAB}{TAB}{TAB}
;Send, F
;Send, {TAB}{TAB}{TAB}

;send, RAM to review - this refund is a duplicate claim within itself. This claim was loaded due to the client has retracted a partial payment, the provider has issued a direct pay for less the requested refund, or a partial payment was issued in house by the provider. This refund is valid and remains due. If you feel this should be disputed with the client you will need to suspende the claim in order to submit this claim to the client for possible closure. Please enter a suspense request not a denial request. Thank you.
;Sleep, 2000

;MouseClick, left,  539,  14
;Sleep, 100
;Send, {ENTER}
;Sleep, 100

;Sleep, 600

;send, {ALTDOWN}{ALTUP}
;Sleep, 100
;send, F
;Sleep, 100
;Send, C
;Sleep, 2000


WinWait, Microsoft Excel - NO LONGER PPN, 
	IfWinNotActive, Microsoft Excel - NO LONGER PPN, , WinActivate, Microsoft Excel - NO LONGER PPN, 
	WinWaitActive, Microsoft Excel - NO LONGER PPN, 
	Sleep, 600


 

Sleep, 600
Send, {RIGHT}x{DOWN}{LEFT}

}
Esc::Pause

Return