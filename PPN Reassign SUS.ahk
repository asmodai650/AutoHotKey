;InputBox,a, Resolve SUS, What comments should be entered while removing the SUS?
;	if errorLevel =1
;		ExitApp
; macro removes SUS.
; Runs from an open SUS Request window, and a spreadsheet with one column of Safari Claim #s.
; update the "close" reason below.
; update the "comments".
;*****resolve sus
Loop,
{
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
Dept=%cell2%
ResearchReason=%cell3%
note=%cell4%

	claim := Clipboard
		if claim =
			break

IfWinNotExist, Suspense Request #
	MsgBox, Please OPEN a "Suspense Request #12345678 for Claim xxx01-0001-1" windwow
WinWait, Suspense Request, 
	IfWinNotActive, Suspense Request, , WinActivate, Suspense Request, 
	WinWaitActive, Suspense Request, 
		Clipboard =
	Sleep, 100
	Send, ^f
	sleep, 200

WinWait, Search, 
	IfWinNotActive, Search, , WinActivate, Search, 
	WinWaitActive, Search, 
	sleep, 100
	Send, %claim%{ENTER} ; safari claim id
	Sleep 300
	Send {TAB}{end}{ENTER}
	sleep, 2000
	
	
	SetTitleMatchMode, 1
WinWaitActive, Suspense Request, 
	SetTitleMatchMode, 2
WinWaitActive, %claim%, 
	SetTitleMatchMode, 1
	sleep, 500
	MouseMove, 555,172
	sleep,100
	MouseClick, right,  555,  175
	Sleep, 1000
	;MouseMove, 334, 236
	;Pause
	PixelGetColor, color, 334, 236 ; green in close approved
	if color = 0x008000
			{
	; can "close" sus options:
			;Send, {down} ; close-deny claim
			;Send, {down 2} ; close-Pass QA
			;Send, {down 3} ; close- Reactivate
			;Send, {down 4} ; close- Approved
			Send, {down 5} ; Reassign
			Sleep, 300
			send, {enter}
			;MouseClick, left,  369,  233 ; close approved
				Sleep, 100


			WinWaitActive, Suspense Request Tracking - Claim # , 
; "comments"	
				sleep, 300
				send, %dept%{tab} ; aim client
				sleep, 200
				send, demo{tab}
				sleep, 200
				send, %ResearchReason%{tab}
				sleep, 200
				send, %note%
				sleep, 200
				send,^s
				WinWaitNotActive, Suspense Request Tracking - Claim # 
			IfWinExist, save, 
				{
				sleep, 100
				Send, y
				}
				/*
			WinWaitActive, Close Request, 
				Sleep, 100
				Send, y
*/
			IfWinExist, Suspense Request Tracking
				sleep, 1000
			WinWaitActive, Suspense Request #, 
			sleep,100
			result = Reassigned
			}
		else
			result = not in SUS
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep 100
Send, %result%{Down}
}
MsgBox, Done.
return




Pause::Pause
esc::ExitApp