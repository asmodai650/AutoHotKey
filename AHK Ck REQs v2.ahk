#SingleInstance, force

WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	;Sleep, 100
	;send, ^{home}
	;Sleep, 100
	;Send, {down}{Right 4}


Loop 
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

unid=%cell1%
prov=%cell2%
comm=%cell3%
genNote=%cell4%	

if unid=0
	break

WinWait,Unidentified Provider Claim No.:, 
	IfWinNotActive,Unidentified Provider Claim No.:, , WinActivate,Unidentified Provider Claim No.:, 
	WinWaitActive,Unidentified Provider Claim No.:, 
	Sleep, 100
	Send, ^f

WinWaitActive, Search, 
		Sleep, 100
		Send, %unid%{enter}
		sleep, 100
		send, {TAB}{ENTER}
	

WinWaitActive,Unidentified Provider Claim No.:, 
	Send, {SPACE}

WinWaitActive, Search, 
		Sleep, 100
	Send,{SHIFTDOWN}{TAB  2}{SHIFTUP}{HOME}pp{TAB  2}
	Sleep, 100
	send, %prov%{enter}
	sleep, 100
	send, {TAB}{ENTER}


WinWaitActive,Unidentified Provider Claim No.:, 
	Sleep, 100
	send, ^s
	sleep, 1000
	Send, !t
	Sleep, 100
	send,{DOWN  2}{RIGHT  2}{ENTER}

WinWaitActive, New Provider Unidentified Check Request, 
	Sleep, 100
	Send, {TAB  2}
	Sleep, 100
	Send, 2{TAB}
	sleep, 100
	send, {tab}
	send, n{TAB  2}
	sleep, 100
	Send, o{TAB  2}
	Sleep, 100
clipboard = %comm%
;	sendinput, %comm%
send, ^V
	sleep, 100
	
send, ^s
	Sleep, 1000

/*
	send, {F2}
tooltip, opening notes window

WinWaitActive, Notes
				; Wait for New Notes button to become active
				Loop 
				{ 
					sleep, 300
				ImageSearch, FoundX, FoundY, 0,0, 700,700, \\Aim\accounts receivable$\AIM 2012 AR Reports\HCA\HCA Check Request Macro\Safari Notes-New.bmp
				if ErrorLevel = 0
				Break
				}
			Sleep, 100
			Send, ^n ; New button

	Sleep, 500
	loop, if %A_Cursor%=Wait
		sleep, 400
	
WinWaitActive, Notes
		loop
			{
			ControlGetFocus, c
			if c=Edit3
				break
			else
				{
				ToolTip, waiting for notes window.
				sleep, 300
				}
			}
ToolTip,
sleep, 4000	
WinWaitActive, Notes
Loop 
				{ 
				tooltip, searching %a_index%	
				sleep, 500
				ImageSearch, FoundX, FoundY, 0,0, 700,700, \\Aim\accounts receivable$\AIM 2012 AR Reports\HCA\HCA Check Request Macro\fullnotes.bmp
				if ErrorLevel = 0
					Break
				}
tooltip,
			sleep, 1000
			Send, {Tab 3}
			Sleep, 250
			Send, g{Tab 3}
			Sleep, 250
			Send, %genNote%
			Sleep, 1000
			Click, 77,42 ; Save button
			result= x
			Sleep, 1000
		;Pause
			WinClose, Notes
			WinWaitClose, Notes
	
*/	
result= x
WinWaitActive, Check Request #, 
WinClose, Check Request

WinWaitActive,Unidentified Provider Claim No.:, 
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 100
	send, x
	Sleep, 100
	Send, {down}
}
MsgBox, Request completed.
ESC::ExitApp
Pause::Pause

