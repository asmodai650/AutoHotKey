 #SingleInstance force

MsgBox, Open a patient claim window in Safari.

MsgBox, Select the first claim in excel.

InputBox, c, , How many claims?, ,200 ,125

loop, %c%
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 100
Clipboard =
Sleep, 100
Send, ^c
Sleep, 100
ClipWait
Sleep, 100
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 100
Send, ^f
Sleep, 100
WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 100
Send, ^v
Sleep, 100
Send, {ENTER}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {ENTER}
Sleep, 100
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 1000
ImageSearch, FoundX, FoundY, 1, 560, 225, 607, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\DREQ.PNG
	if ErrorLevel
		{
		goto, NotInDREQ
		}
Send, ^r
Sleep, 100
WinWait, Claim Requests for Claim ID
IfWinNotActive, Claim Requests for Claim ID
WinActivate, Claim Requests for Claim ID
WinWaitActive, Claim Requests for Claim ID
Sleep, 100
Send, {END}
Sleep, 100
Send, {ENTER}
Sleep, 100
WinWait, Claim Denial Request #
IfWinNotActive, Claim Denial Request #
WinActivate, Claim Denial Request #
WinWaitActive, Claim Denial Request #
Sleep, 100
MouseClick, Left, 206, 325, 2
Sleep, 100
WinWait, DM Account Manager Approval
IfWinNotActive, DM Account Manager Approval
WinActivate, DM Account Manager Approval
WinWaitActive, DM Account Manager Approval
Sleep, 100
Send, d
Sleep, 100
Send, {TAB}
Sleep, 100
SendInput, Removing denial request.
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {ENTER}
Sleep, 100
WinWait, Claim Denial Request #
IfWinNotActive, Claim Denial Request #
WinActivate, Claim Denial Request #
WinWaitActive, Claim Denial Request #
Sleep, 100
Send, ^s
Sleep, 1000
Send, !e
Sleep, 100
Send, {ENTER}
Sleep, 100
WinWait, Close Request
IfWinNotActive, Close Request
WinActivate, Close Request
WinWaitActive, Close Request
Sleep, 100
Send, {ENTER}
Sleep, 100
WinWait, Claim Requests for Claim ID
IfWinNotActive, Claim Requests for Claim ID
WinActivate, Claim Requests for Claim ID
WinWaitActive, Claim Requests for Claim ID
Sleep, 100
Send, ^{F4}
Sleep, 100
WinWaitClose, Claim Requests for Claim ID
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 100
Send, {F5}
Sleep, 100
Loop
{
ImageSearch, FoundX, FoundY, 1, 560, 225, 607, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\DREQ.PNG
	if ErrorLevel
		{
		Sleep, 1000
		}
	else
		{
		goto, DREQRemoved
		}
}

NotInDREQ:
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 100
Send, {Right}
Sleep, 100
SendInput, Claim is not in DREQ
Sleep, 100
Send, {DOWN}
Sleep, 100
Send, {LEFT 2}
Sleep, 100
updatetext =  %A_ComputerName%|%A_UserName%|%A_Now%|%A_ScriptName%
	loop 3
	{
	FileAppend, %updatetext%, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Public Macros\Macro Tracker\%A_ComputerName%.%A_UserName%.txt
					if errorlevel
					{
									sleep, 1000
									continue
					}
	else
					break
	}
goto, LoopEnd

DREQREMOVED:
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 100
Send, {Right}
Sleep, 100
SendInput, DREQ removed
Sleep, 100
Send, {DOWN}
Sleep, 100
Send, {LEFT 2}
Sleep, 100
updatetext =  %A_ComputerName%|%A_UserName%|%A_Now%|%A_ScriptName%
loop 3
{
FileAppend, %updatetext%, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Tracker\%A_ComputerName%.%A_UserName%.txt
                if errorlevel
                {
                                sleep, 1000
                                continue
                }
else
                break
}
LoopEnd:
}

MsgBox, Done!

ExitApp

Pause::Pause
Esc::ExitApp