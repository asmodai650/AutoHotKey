;Safari - Application Error Report
MsgBox,,Active to SUS, This macro takes an active claim and moves it to SUS. The spreadsheet should be set up:`n  Claim Id (Safari ID)`n  Department/LOB (example:Client Services)`n  Assigned To`n  SUS Reason (example: Analyst Research)`n  Note`n`nPlease select the first claim on the spreadsheet to be moved.`n`n             Press "Pause" at anytime to Pause or "ESC" to Exit.

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep, 100
send, {Home}{ShiftDown}^{Down}{ShiftUp}
sleep, 100,
send,^c
sleep, 100
send, ^{right}{Right}

StringReplace, clipboard, clipboard, `r`n, |, All
StringSplit, claim, Clipboard,`r `n `t|
Count=%claim0%
Count-=1
inputbox,c,,how many claims?,,200,125,,,,,%Count%

loop, %c%
{
TrayTip,, %a_index% of %c%.,1

WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 100
	send, {ShiftDown}{Space}{ShiftUp}
	Sleep, 100
	Send, ^c{down}{up}
	ClipWait

StringReplace, clipboard, clipboard, `r`n, , All
StringSplit, cell, Clipboard, %a_tab%

claim=%cell1%
LOB=%cell2%
CAS=%cell3%
SUSr=%cell4%
note=%cell5%

TrayTip,, %a_index% of %c%.`n%lob% %CAS%`n%SUSr%`n%note%.,
IfWinNotExist, Patient Claim
	MsgBox, Open any Patient claim in Safari to start.
WinWait, Patient Claim, 
IfWinNotActive, Patient Claim, , WinActivate, Patient Claim, 
WinWaitActive, Patient Claim, 

Sleep, 100
Send, ^f
sleep, 200
IfWinExist, Save Changes
	Send, y
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
sleep, 100
Send, %claim%
Sleep 300
Send {ENTER}
Sleep 300

;if claim not found
IfWinExist, SafariCorp Search
{
WinWait, SafariCorp Search
	IfWinNotActive, SafariCorp Search, , WinActivate, SafariCorp Search, 
	WinWaitActive, SafariCorp Search, 
	sleep, 300
	send, {enter}
	WinWaitClose, SafariCorp Search
	WinWaitActive, Search, 
	WinClose, Search, 
	WinwaitClose, search
	result=There were no items that match your selection criteria. 
	goto, End
}

;search for multiple lines
send,{TAB}{DOWN}
Sleep 200
PixelSearch, Px, Py, 337, 203, 350, 203, 0xC56A31, 3, Fast
	if ErrorLevel = 0
	{
	Sleep 300
	Send WinClose
	result= Review- Multiple Refunds
	GoTo, End
	}
Send, {tab}{enter}

; claim found, move to sus
WinWaitActive, Patient Claim, 
Sleep, 1000
Send, ^r
TrayTip,, %a_index% of %c%.`nClaim %claim%`nsuspending,



WinWaitActive, Claim Requests for Claim ID, 
Sleep, 100

; checks if in sus or den already
PixelGetColor, colo, 454, 67
;MsgBox The color at the current cursor position is %colo%. the grey one is 0x99A8AC
IfEqual, colo, 0x99A8AC
{
TrayTip,, %a_index% of %c%.`nClaim %claim%`ncan't click sus
	WinClose
	Sleep 300
	result= could not suspend
	GoTo, end
}

;suspends claim
Send,!e
sleep, 100
Send, s

WinWaitActive, Suspense Request Tracking - Claim, 
Send, %lob%{TAB}%CAS%{TAB}%SUSr%{TAB}%note%.^s
ToolTip, saving
sleep, 100
TrayTip,, %a_index% of %c%.`nClaim %claim%`nsuspended,
result= suspended on %A_MM%/%A_DD%/12.
WinWaitActive, Claim Requests for Claim ID, 
ToolTip,
WinClose, Claim Requests for Claim ID, 
WinWaitActive, Patient Claim, 
Sleep, 300

end:
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel,
	sleep, 100
	send, %result%{down}
	sleep, 100

tooltip, 
TrayTip,
}
MsgBox, done
return
Pause::Pause
esc::
ExitApp
