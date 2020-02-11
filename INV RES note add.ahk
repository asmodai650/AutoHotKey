#SingleInstance, force

MsgBox, This macro Reassigns INV RES claims in Safari.`nSet up in Excel:`nClaimNo`nDept`nAssign To`nReason`nNote

;InputBox,a, Resolve SUS, What comments should be entered while removing the SUS?
;	if errorLevel =1
;		ExitApp
; macro removes SUS.
; Runs from an open SUS Request window, and a spreadsheet with one column of Safari Claim #s.
; update the "close" reason below.
; update the "comments".
;*****resolve sus
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 100
	send, {Home}{Right 5}


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

claim = %cell1%
Dept = %cell2%
AsgnTo = %cell3%
ResearchReason = %cell4%
note = %cell5%

	if claim =
		break

IfWinNotExist, Research Request
	MsgBox, Please OPEN a "Research Request #12345678 for Claim xxx01-0001-1" windwow
WinWait, Research Request, 
	IfWinNotActive, Research Request, , WinActivate, Research Request, 
	WinWaitActive, Research Request, 
		Clipboard =
	Sleep, 100
	Send, ^f
	sleep, 200

WinWait, Search, 
	IfWinNotActive, Search, , WinActivate, Search, 
	WinWaitActive, Search, 
		sleep, 100
	send, {ShiftDown}{tab 2}{ShiftUp} 
		sleep, 100
	send, i{tab 2} ; ins claim id
		sleep, 100
	Send, %claim%{ENTER} ; safari claim id
		Sleep 300
	Send {TAB}{end}{ENTER}
	sleep, 2000
	
	
	SetTitleMatchMode, 1
WinWaitActive, Research Request, 

sleep, 5000
	;SetTitleMatchMode, 2
;WinWaitActive, %claim%, 
	SetTitleMatchMode, 1
	sleep, 2000
		ImageSearch, oX, oY, 0,0, 1000, 1000,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Reassign Research\StatusCR.bmp
	if ErrorLevel = 2
    MsgBox Could not conduct the search.
else if ErrorLevel = 1
    result = not in research
else
{	
	send, {f2}
	ToolTip, opening notes
		sleep, 2000
WinWaitActive, Notes
		sleep, 2000
WinWaitActive, Notes
		mousemove, 0, 0
		send, !n
		sleep, 2000
	Loop
	{
	ToolTip, Opening New Notes window. %a_index% 
	WinWaitActive, Notes
			sleep, 1000
	ImageSearch, oX, oY, 0,0, 500, 500,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Reassign Research\New Note window.bmp
			if ErrorLevel = 2
				MsgBox Could not conduct the search.
	else if ErrorLevel = 1
		continue
	else
		break
	}
	
		WinWaitActive, Notes
		while A_Cursor = Wait
					sleep, 100
		loop, 10
			{
			ToolTip, Waiting for new notes.`r%A_Index% of 10
			sleep, 1000
			}
			ToolTip
			WinWaitActive, Notes
				send, {tab 3}
					sleep, 200
				send, r{tab 3}
					sleep, 200
				SendInput, %note%
					sleep, 200
				;send, {shiftdown}{tab}{shiftup}
					sleep, 200
			IfWinNotActive, Notes, , WinActivate, Notes, 
			WinWaitActive, Notes,
				 sleep, 1000
			; *******************************saving
				loop
				{
					mousemove, 0, 0
					PixelGetColor, black, 74, 47
					if black = 0x000000
						{
						MouseClick, left,  78, 45 ; save
							ToolTip, Saving
						Sleep, 2000
							while A_Cursor = Wait
								{
								tooltip, %a_index% Saving
								sleep, 1000
								}
						}
					else
						break
				}
						
				ToolTip, Save complete
				mousemove, 530, 10
					sleep, 200
				MouseClick, left,  539,  17 ; close
				mousemove, 10,10
			WinWaitActive, Research Request #, 
				sleep,100
				result = note added
			ToolTip
	}		
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