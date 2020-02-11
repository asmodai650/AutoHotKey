; this macro adds the SF based on a spreadsheet. Set up, safari claim # then UID
#SingleInstance, Force

MsgBox, 33, , This macro adds the UID. In Excel list the Safari claim ID followed by the UID.`nContinue?
	IfMsgBox cancel 
		ExitApp
	
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep 100
	send, {home}
	Sleep 100



Loop
{
Beginning:

WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep 100
	ClipBoard := 
	Sleep 100
	Send {SHIFTDOWN}{RIGHT}{SHIFTUP}
	Sleep 300
	send, ^c
		ClipWait
	Send {RIGHT 2}
	Sleep 100
	
StringSplit, cell, clipboard, %A_Tab%,
Claim = %cell1%		
UID = %cell2%

if claim = 
	break

;Pulling Claim in Safari
FindClaim:
sleep, 1000

IfWinExist, Patient Claim
	{
	WinActivate, Patient Claim
	WinWaitActive, Patient Claim,, 10
	}
IfWinExist, Claim ID 
	{
	WinActivate, Claim ID 
	WinWaitActive, Claim ID,, 10
	}
ToolTip,

sleep, 100
send, ^f
TrayTip, ,%A_Index%
Sleep 500


IfWinExist, Save Changes,
	{
	WinWait, Save Changes, 
		IfWinNotActive, Save Changes, , WinActivate, Save Changes, 
		WinWaitActive, Save Changes, 
		Sleep, 100
		Send, {ALTDOWN}n{ALTUP}
		Sleep 300
	}

;Identify Claims w/ Multiple Refunds
WinWait, Search
	IfWinNotActive, Search, , WinActivate, Search, 
	WinWaitActive, Search, 
	Sleep, 300
	Send, %Claim%
	Sleep, 300
	Send {ENTER}
	Sleep, 300
IfWinExist, SafariCorp Search
	{
	Sleep 300
	WinClose, SafariCorp Search
	WinWaitClose, SafariCorp Search
	Sleep 300
	send, !c
	WinClose, search
	WinWaitClose, search

WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep 300
	Send, No Items that match your selection criteria.
	Sleep 200
	Send {DOWN}{HOME}
	Sleep 300
	GoTo Beginning
	}

; selects claim to see if multiple exists
Send, {TAB}{DOWN}
Sleep 200

PixelSearch, Px, Py, 337, 203, 350, 203, 0xC56A31, 3, Fast
if ErrorLevel = 0
	{
	Sleep 300
	WinClose
	Sleep 300
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep 300
	Send Review- Multiple Refunds
	Sleep 200
	Send {DOWN}{HOME}
	Sleep 300
	GoTo Beginning
	}

; selects claim if only one
Sleep 500
Send {ENTER}
Sleep 500



; makes sure the Patient claim screen has loaded
WinWaitActive, Patient Claim
	sleep, 1000
	send, ^r
	WinWaitActive, Claim Requests for Claim ID
		WinClose, Claim Requests for Claim ID
WinWaitActive, Patient Claim

;Add UID
WinWaitActive, Patient Claim
	send !v
	sleep, 100
	send, u
	sleep, 200
	if A_CaretY = 201 ; UID line location
		send, %UID%{tab}
	else
		{
		loc= %A_CaretY%
		MsgBox, Cursor not in correct location.  %loc%
		Pause
		}
	
	result= added
	
	ControlGetText, groupno, ThunderRT6TextBox17, Patient 
		if groupno is not alnum
			{
				;MsgBox, %groupno% is an alnum.
				;~ ControlFocus, ThunderRT6TextBox17, Patient
				;~ Control, EditPaste, 0, ThunderRT6TextBox17, Patient 
				;~ ControlFocus, ThunderRT6TextBox17, Patient
				;~ send, {tab}
				ControlFocus, ThunderRT6TextBox17, Patient
				Control, EditPaste, 0, ThunderRT6TextBox17, Patient 
				ControlFocus, ThunderRT6TextBox17, Patient
				send, {tab}
			}
		else if groupno =
			{
				;MsgBox, %groupno% is an null.
				ControlFocus, ThunderRT6TextBox17, Patient
				Control, EditPaste, 0, ThunderRT6TextBox17, Patient 
				ControlFocus, ThunderRT6TextBox17, Patient
				send, {tab}
				;Pause
			}
			
	;~ ControlGetText,explanation, Edit1, Patient 
		;~ if explanation = Other
			;~ {			
			;~ ControlFocus, Edit1, Patient
			;~ send, Overpaid - See Comments{tab}
			;~ }

sleep, 100	
send, ^s
tooltip,Saving,500,500
sleep, 1000	

				MouseMove, 10,10,,R
				loop, if %A_Cursor%=Wait
					{
					sleep, 1000
					ToolTip, %a_index%
					}
					ToolTip
				WinWait, Claim ID , , 10
					{
					IfWinActive, Missing Fields
						{
						WinClose, Missing Fields
						WinWaitClose, Missing Fields
						sleep, 100
						}
					else
						sleep, 100
					
					IfWinActive, Failed Validation
						{
						Send {Enter}
						Result = Failed Validation
						}
					else
						{
						IfWinNotActive, Claim ID , , WinActivate, Claim ID , 
						WinWaitActive, Claim ID,,15 
						sleep, 200
					
						}
					}	



WinWaitActive, Claim ID,,10
	ToolTip,

WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 300
	Send %result%{DOWN}{HOME}
tooltip,Waiting for Safari,500,500

FoundX =
FoundY =
}
tooltip
MsgBox, done
Esc::ExitApp
Pause::Pause

]::
{
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 300
	Send {HOME}
}
Reload