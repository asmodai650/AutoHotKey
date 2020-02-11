InputBox, c,Unidentified script, Please open up the Safar window: Unidentified Provider Claim No for any claim.`nHighlight the first cell in the spreadsheet.`n`nHow many claims?
; get the uni #
FormatTime, btim, , h:mmtt 
Clipboard=
loop, %c%
{
TrayTip,, %a_index% of %c%.,

WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel -, , WinActivate, Microsoft Excel -, 
	WinWaitActive, Microsoft Excel -, 
	Sleep, 100
	Send, ^c
	ClipWait,


WinWait, Unidentified Provider Claim No., 
	IfWinNotActive, Unidentified Provider Claim No., , WinActivate, Unidentified Provider Claim No., 
	WinWaitActive, Unidentified Provider Claim No., 
	Sleep, 100
	Send, {CTRLDOWN}f{CTRLUP}


WinWaitActive, Search, 
	sleep, 100
	Send, {CTRLDOWN}v{CTRLUP}{ENTER}{TAB}{ENTER}

WinWaitActive, Unidentified Provider Claim No., 
	sleep, 500
	Send, {ALTDOWN}t{ALTUP}
	sleep, 300,
	Send, {DOWN  3}{ENTER}
	

WinWaitActive, Edit Amount
	sleep, 100
	send, {enter}
WinWaitActive, Search, 

WinWait, Microsoft Excel -, 
	IfWinNotActive, Microsoft Excel -, , WinActivate, Microsoft Excel -, 
	WinWaitActive, Microsoft Excel -, 
	Sleep, 100
	Send, {RIGHT  7}
	Sleep, 100
	Send, ^c
	sleep, 100
	Send, {right 2}


;search for claim to apply to
WinWait, Search, 
	IfWinNotActive, Search, , WinActivate, Search, 
	WinWaitActive, Search, 
	Send, ^v{ENTER}
	sleep, 300

IfWinExist, SafariCorp Search
	{
		WinWaitActive, SafariCorp Search, 
			MouseClick, left,  136,  8
			Sleep, 100
			Send, {ENTER}
		WinWaitActive, Search, 
			WinClose, Search
			Sleep, 100
		WinWaitActive, Reapply Provider Unidentified Claim,
			WinClose, Reapply Provider Unidentified Claim 
 		WinWaitActive, Unidentified Provider Claim No, 	
		note=Claim id not found
	goto, noclaimid
	}
	sleep, 300
	/*
WinGetActiveTitle, title
If title contains dd Claim
	{
		MsgBox, test 4
		WinWaitActive, Add Claim, 
			Sleep, 100
			Send, {tab}{Enter}
			
		WinWaitActive, Reapply Provider Unidentified Claim,
			WinClose, Reapply Provider Unidentified Claim 
 		WinWaitActive, Unidentified Provider Claim No, 	
		note=CB difference
	goto, noclaimid
	}
*/
		WinWaitActive, Search, 
	send,{TAB}{ENTER}


WinWaitActive, Reapply Provider Unidentified Claim #:, 
	sleep, 100
	mousemove, 139, 63
	Sleep, 100
	MouseClick, left,  139, 63
	Sleep, 100

WinWait, Reapply?, 
	IfWinNotActive, Reapply?, , WinActivate, Reapply?, 
	WinWaitActive, Reapply?, 
	Sleep, 100
;MsgBox, did it work 3	
	Send, {ENTER}

WinWaitActive, Unidentified Provider Claim No., 
;worked
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel -, , WinActivate, Microsoft Excel -, 
	WinWaitActive, Microsoft Excel -, 
	Sleep, 100
	Send, complete at %A_Hour%:%A_Min%:%A_Sec%
	sleep, 100,
	send, {down}{Left 9}
continue
noclaimid:
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel -, , WinActivate, Microsoft Excel -, 
	WinWaitActive, Microsoft Excel -, 
	Sleep, 100
	Send, %note%
	note=
	sleep, 100,
	send, {down}{Left 9}
TrayTip,
}
FormatTime, etim, , h:mmtt 
MsgBox, Done - %c%`nStart time %btim%`nEnd time %etim%

Pause::Pause
Esc::ExitApp
