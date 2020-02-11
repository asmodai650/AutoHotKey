#SingleInstance Force

InputBox, Name, , Enter your first name exactly as it appears in Safari.

InputBox, Reason, , Enter 4 if you have a CB profile in Safari or 3 if you have a DM profile.

Loop
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 1000
Send, {Home}
Sleep, 1000
Clipboard =
Sleep, 1000
Send, {ShiftDown}{Space}{ShiftUp}
Sleep, 1000
Send, ^c
ClipWait, 5
	if Clipboard =
		break
Sleep, 1000
StringReplace, Clipboard, Clipboard, `r`n, , All
StringSplit, Cell, Clipboard, %A_Tab%
ClaimID = %Cell1%
Note = %Cell2%
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 1000
Send, ^f
Sleep, 1000
WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 1000
Send, %ClaimID%
Sleep, 1000
Send, {Enter}
Sleep, 1000
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
		Sleep, 1000
		Send, {Enter}
		Sleep, 1000
		WinWait, Patient Claim
		IfWinNotActive, Patient Claim
		WinActivate, Patient Claim
		WinWaitActive, Patient Claim
		Sleep, 1000
		Send, ^r
		Sleep, 1000
		WinWait, Claim Requests for Claim ID:
		IfWinNotActive, Claim Requests for Claim ID:
		WinActivate, Claim Requests for Claim ID:
		WinWaitActive, Claim Requests for Claim ID:
		Sleep, 1000
		Send, {Down 5}
		Sleep, 1500
		Send, {Enter}
		Sleep, 1000
		WinWait, Claim Denial
		IfWinNotActive, Claim Denial
		WinActivate, Claim Denial
		WinWaitActive, Claim Denial
		Sleep, 1000
		goto, DREQ
		}
Send, {Tab}
Sleep, 1000
Send, {Enter}
Sleep, 1000
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 1000
Send, ^r
Sleep, 1000
WinWait, Claim Requests for Claim ID:
IfWinNotActive, Claim Requests for Claim ID:
WinActivate, Claim Requests for Claim ID:
WinWaitActive, Claim Requests for Claim ID:
Sleep, 1000
Send, !e
Sleep, 1500
Send, {Enter}
Sleep, 1000
WinWait, Claim Denial
IfWinNotActive, Claim Denial
WinActivate, Claim Denial
WinWaitActive, Claim Denial
Sleep, 1000
Send, %Name%
Sleep, 1500
Send, {Enter}
Sleep, 1000
MouseClick, Left, 48, 166, 1
WinWait, Transaction Reasons
IfWinNotActive, Transaction Reasons
WinActivate, Transaction Reasons
WinWaitActive, Transaction Reasons
Sleep, 1000
Send, {Tab 2}
Sleep, 1000
Send, {Down %Reason%}
Sleep, 1000
Send, {Enter}
Sleep, 1000
Send, {Down}
Sleep, 1000
Send, {Enter}
Sleep, 1000
WinWaitClose, Transaction Reasons
WinActivate, Claim Denial
WinWaitActive, Claim Denial
Send, !f
Sleep, 1500
Send, {Down}
Sleep, 1000
Send, {Enter}
Sleep, 1000
WinWaitClose, Claim Denial
WinActivate, Claim Requests for Claim ID:
WinWaitActive, Claim Requests for Claim ID:
Sleep, 10000
Send, {F5}
Sleep, 1000
ImageSearch, FoundX, FoundY, 379, 147, 487, 460, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Claim Denied Blue.PNG
	if ErrorLevel = 0
		{
		ExcelNote = Done
		WinClose, Claim Requests for Claim ID:
		goto, Excel
		}
Send, {Down 5}
Sleep, 1000
Send, {Enter}
Sleep, 1000
DREQ:
WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 1500
MouseClick, Left, 107, 326, 1
Sleep, 1000
Send, {Enter}
Sleep, 1000
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 1000
Send, {Tab}
Sleep, 1000
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 1000
	SendInput, a
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	SendInput, %Note%
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	Send, !f
	Sleep, 1500
	Send, {Down}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1500
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
	Send, {Down 5}{Enter}
	Sleep, 1000
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 1500
MouseClick, Left, 107, 326, 1
Sleep, 1000
Send, {Down}
Sleep, 1000
Send, {Enter}
Sleep, 1000
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 1000
Send, {Tab}
Sleep, 1000
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 1000
	SendInput, a
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	SendInput, %Note%
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	Send, !f
	Sleep, 1500
	Send, {Down}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1500
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
	Send, {Down 5}{Enter}
	Sleep, 1000
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 1500
MouseClick, Left, 107, 326, 1
Sleep, 1000
Send, {Down 2}
Sleep, 1000
Send, {Enter}
Sleep, 1000
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 1000
Send, {Tab}
Sleep, 1000
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 1000
	SendInput, a
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	SendInput, %Note%
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	Send, !f
	Sleep, 1500
	Send, {Down}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1500
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
	Send, {Down 5}{Enter}
	Sleep, 1000
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 1500
MouseClick, Left, 107, 326, 1
Sleep, 1000
Send, {Down 3}
Sleep, 1000
Send, {Enter}
Sleep, 1000
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 1000
Send, {Tab}
Sleep, 1000
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 1000
	SendInput, a
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	SendInput, %Note%
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	Send, !f
	Sleep, 1500
	Send, {Down}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1500
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
	Send, {Down 5}{Enter}
	Sleep, 1000
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 1500
MouseClick, Left, 107, 326, 1
Sleep, 1000
Send, {Down 4}
Sleep, 1000
Send, {Enter}
Sleep, 1000
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 1000
Send, {Tab}
Sleep, 1000
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	;WinClose, Level
	;WinWaitClose, Level
	}
	else
	{
	Send, {Tab 3}
	Sleep, 1000
	SendInput, a
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	SendInput, %Note%
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	Send, !f
	Sleep, 1500
	Send, {Down}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1500
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
	Send, {Down 5}{Enter}
	Sleep, 1000
	}

WinWait, Claim Denial Request
IfWinNotActive, Claim Denial Request
WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 1500
MouseClick, Left, 107, 326, 1
Sleep, 1000
Send, {Down 5}
Sleep, 1000
Send, {Enter}
Sleep, 1000
SetTitleMatchMode, 2
WinWait, Level
IfWinNotActive, Level
WinActivate, Level
WinWaitActive, Level
Sleep, 1000
Send, {Tab}
Sleep, 1000
ImageSearch, FoundX, FoundY, 0, 30, 150, 75, \\aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Images\Safari - Approved.PNG
	if ErrorLevel = 0
	{
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
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
	Sleep, 1000
	SendInput, a
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	SendInput, %Note%
	Sleep, 1000
	Send, {Tab}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1000
	Send, !f
	Sleep, 1500
	Send, {Down}
	Sleep, 1000
	Send, {Enter}
	Sleep, 1500
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
	Send, {Down 5}{Enter}
	Sleep, 1000
	}

Excel:
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 1000
Send, {Right 2}
Sleep, 1000
Send, %ExcelNote%
Sleep, 1000
Send, {Down}
Sleep, 1000

updatetext =  %A_ComputerName%|%A_UserName%|%A_Now%|%A_ScriptName%

loop 3
{
FileAppend, %updatetext%, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Tracker\%A_ComputerName%.%A_UserName%.txt
                if errorlevel
                {
				Sleep, 10000
				continue
                }
else
                break
}

}

MsgBox, Done!

Pause::Pause
Esc::ExitApp