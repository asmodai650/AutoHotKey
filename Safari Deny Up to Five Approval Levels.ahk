#SingleInstance Force

InputBox, Name, , Enter your first name exactly as it appears in Safari.

InputBox, Reason, , Enter 4 if you have a CB profile in Safari or 3 if you have a DM profile.

Loop
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep 100
Send, {Home}
Sleep, 100
Clipboard =
Sleep, 100
Send, {ShiftDown}{Space}{ShiftUp}
Sleep, 100
Send, ^c
ClipWait, 5
	if Clipboard =
		break
Sleep, 100
StringReplace, Clipboard, Clipboard, `r`n, , All
StringSplit, Cell, Clipboard, %A_Tab%
ClaimID = %Cell1%
Note = %Cell2%
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep 100
Send, ^f
Sleep, 100
WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep 100
Send, %ClaimID%
Sleep, 100
Send, {Enter}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 150, 300, 300, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Macros\Macro Images\Safari Status Check\DEN.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Already Denied
		WinClose, Search
		WinWaitClose, Search
		goto, Excel
		}
ImageSearch, FoundX, FoundY, 0, 150, 300, 300, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Macros\Macro Images\Safari Status Check\SUS.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Needs SUS Resolved
		WinClose, Search
		WinWaitClose, Search
		goto, Excel
		}
ImageSearch, FoundX, FoundY, 0, 150, 300, 300, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Macros\Macro Images\Safari Status Check\DREQ.PNG
	if ErrorLevel = 0
		{
		Send, {Tab}
		Sleep, 100
		Send, {Enter}
		Sleep, 100
		WinWait, Patient Claim
		IfWinNotActive, Patient Claim
		WinActivate, Patient Claim
		WinWaitActive, Patient Claim
		Sleep 100
		Send, ^r
		Sleep, 100
		WinWait, Claim Requests for Claim ID:
		IfWinNotActive, Claim Requests for Claim ID:
		WinActivate, Claim Requests for Claim ID:
		WinWaitActive, Claim Requests for Claim ID:
		Sleep 100
		;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
		;Send, {Down 5}
		Send, {End}
		Sleep, 500
		Send, {Enter}
		Sleep, 100
		WinWait, Claim Denial
		IfWinNotActive, Claim Denial
		WinActivate, Claim Denial
		WinWaitActive, Claim Denial
		Sleep 100
		goto, DREQ
		}
Send, {Tab}
Sleep, 100
Send, {Enter}
Sleep, 100
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep 100
Send, ^r
Sleep, 100
WinWait, Claim Requests for Claim ID:
IfWinNotActive, Claim Requests for Claim ID:
WinActivate, Claim Requests for Claim ID:
WinWaitActive, Claim Requests for Claim ID:
Sleep 100
Send, !e
Sleep, 500
Send, {Enter}
Sleep, 100
WinWait, Claim Denial
IfWinNotActive, Claim Denial
WinActivate, Claim Denial
WinWaitActive, Claim Denial
Sleep 100
Send, %Name%
Sleep, 500
Send, {Enter}
Sleep, 100
MouseClick, Left, 48, 166, 1
WinWait, Transaction Reasons
IfWinNotActive, Transaction Reasons
WinActivate, Transaction Reasons
WinWaitActive, Transaction Reasons
Sleep 100
Send, {Tab 2}
Sleep, 100
Send, {Down %Reason%}
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {Down}
Sleep, 100
Send, {Enter}
Sleep, 100
WinWaitClose, Transaction Reasons
WinActivate, Claim Denial
WinWaitActive, Claim Denial
Send, !f
Sleep, 500
Send, {Down}
Sleep, 100
Send, {Enter}
Sleep, 100
WinWaitClose, Claim Denial
WinActivate, Claim Requests for Claim ID:
WinWaitActive, Claim Requests for Claim ID:
Sleep, 1000
Send, {F5}
Sleep, 100
ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
DREQ:
WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
MouseClick, Left, 107, 326, 1
Sleep, 100
Send, {Enter}
Sleep, 100
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 100
Send, {Tab}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 100
	SendInput, a
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	Send, !f
	Sleep, 500
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinWaitClose, Claim Denial Request
	WinActivate, Claim Requests for Claim ID:
	WinWaitActive, Claim Requests for Claim ID:
	Sleep, 1500
	Send, {F5}
	Sleep, 1500
	ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
MouseClick, Left, 107, 326, 1
Sleep, 100
Send, {Down}
Sleep, 100
Send, {Enter}
Sleep, 100
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 100
Send, {Tab}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 100
	SendInput, a
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	Send, !f
	Sleep, 500
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinWaitClose, Claim Denial Request
	WinActivate, Claim Requests for Claim ID:
	WinWaitActive, Claim Requests for Claim ID:
	Sleep, 1500
	Send, {F5}
	Sleep, 1500
	ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
MouseClick, Left, 107, 326, 1
Sleep, 100
Send, {Down 2}
Sleep, 100
Send, {Enter}
Sleep, 100
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 100
Send, {Tab}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 100
	SendInput, a
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	Send, !f
	Sleep, 500
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinWaitClose, Claim Denial Request
	WinActivate, Claim Requests for Claim ID:
	WinWaitActive, Claim Requests for Claim ID:
	Sleep, 1500
	Send, {F5}
	Sleep, 1500
	ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
MouseClick, Left, 107, 326, 1
Sleep, 100
Send, {Down 3}
Sleep, 100
Send, {Enter}
Sleep, 100
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 100
Send, {Tab}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 100
	SendInput, a
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	Send, !f
	Sleep, 500
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinWaitClose, Claim Denial Request
	WinActivate, Claim Requests for Claim ID:
	WinWaitActive, Claim Requests for Claim ID:
	Sleep, 1500
	Send, {F5}
	Sleep, 1500
	ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
MouseClick, Left, 107, 326, 1
Sleep, 100
Send, {Down 4}
Sleep, 100
Send, {Enter}
Sleep, 100
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 100
Send, {Tab}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 100
	SendInput, a
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	Send, !f
	Sleep, 500
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinWaitClose, Claim Denial Request
	WinActivate, Claim Requests for Claim ID:
	WinWaitActive, Claim Requests for Claim ID:
	Sleep, 1500
	Send, {F5}
	Sleep, 1500
	ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
MouseClick, Left, 107, 326, 1
Sleep, 100
Send, {End}
Sleep, 100
Send, {Enter}
Sleep, 100
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 100
Send, {Tab}
Sleep, 100
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	;WinClose, Level
	;WinWaitClose, Level
	ExcelNote = Done
	WinClose, Claim Denial Request
	WinClose, Claim Requests for Claim ID:
	goto, Excel
	}
	else
	{
	Send, {Tab 3}
	Sleep, 100
	SendInput, a
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	Send, {Tab}
	Sleep, 100
	Send, {Enter}
	Sleep, 100
	Send, !f
	Sleep, 500
	Send, {Down}
	Sleep, 100
	Send, {Enter}
	Sleep, 500
	WinWaitClose, Claim Denial Request
	WinActivate, Claim Requests for Claim ID:
	WinWaitActive, Claim Requests for Claim ID:
	Sleep, 1500
	Send, {F5}
	Sleep, 1500
	ImageSearch, FoundX, FoundY, 379, 147, 487, 160, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
	;CHANGED LINE BELOW TO ALLOW MACRO TO GO BOTTOM OF THE LIST 20170216
	Send, {End}
	Sleep, 100
	Send, {Enter}
	;Send, {Down 5}{Enter}
	Sleep, 100
	}

Excel:
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep 100
Send, {Right 2}
Sleep, 100
Send, %ExcelNote%
Sleep, 100
Send, {Down}
Sleep, 100

updatetext =  %A_ComputerName%|%A_UserName%|%A_Now%|%A_ScriptName%

loop 3
{
FileAppend, %updatetext%, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Tracker\%A_ComputerName%.%A_UserName%.txt
                if errorlevel
                {
				Sleep, 1000
				continue
                }
else
                break
}

}

MsgBox, Done!

Pause::Pause
Esc::ExitApp