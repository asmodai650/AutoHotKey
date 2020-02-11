#SingleInstance, force

ToolTip, Starting Macro
run, \\Aim\client_services$\MACROS- Chris Taylor\Matt\Brigett Recodes\Safari Error.ahk
;Loop
ToolTip,
MsgBox, 33, UNH PPN Recode Macro, This macro recodes a claim to the UnithedHealtchare (PPN) contract. `n`nExcel setup:`n	Safari Claim No`; `n	SF1 - FACs Directory`; `n	SF2 - FACs Debtor#`;`n	SF3 - Platform`n`nWould you like to ontinue?`n, 30

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep 100
send, {home}
Sleep 100

loop
{
Beginning:

;Save all data from spreadsheet for claim
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep 100
ClipBoard := 
Sleep 100
ToolTip, Grabing info,0,0

Send {Home}{SHIFTDOWN}{RIGHT 3}{SHIFTUP}
Sleep 300
send, ^c
ClipWait
Send {RIGHT 4}
Sleep 100
StringSplit, cell, clipboard, %A_Tab%,

ClaimID = %cell1%		
sf1 = %cell2%	
sf2 = %cell3%		
sf3 = %cell4%		

StringReplace, sf3, sf3, `r`n, , All



if ClaimID =
	{
		MsgBox Done
		break
	}


;Pulling Claim in Safari
searchpart:
ToolTip
IfWinExist, Claim ID 
	{
	WinActivate, Claim ID 
	}
	else
		{
		WinActivate, Patient Claim, 
		WinWaitActive, Patient Claim
		}
sleep, 100


send, ^f

IfWinExist, Save Changes,
{
WinWait, Save Changes, 
IfWinNotActive, Save Changes, , WinActivate, Save Changes, 
WinWaitActive, Save Changes, 
MouseClick, left,  238,  14

sleep, 100
Send, {ALTDOWN}n{ALTUP}
Sleep 300
}


;Search for claim

;Identify Claims w/ Multiple Refunds
WinWait, Search,, 15
	if errorlevel = 1
		goto, searchpart
	IfWinNotActive, Search, , WinActivate, Search, 
	WinWaitActive, Search, 
	Sleep, 300
	Send, %ClaimID%
	Sleep, 300
	Send {ENTER}
	Sleep, 300
	
	IfWinExist, SafariCorp Search
		{
		Sleep 300
		WinClose, SafariCorp Search
		Sleep 300
		send, !c
		WinClose, search
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep 300
		Send, No Items that match your selection criteria.
		Sleep 200
		Send, {DOWN}{HOME}
		Sleep 300
		GoTo Beginning
		}

	

;Checks for Suspense Status
	ImageSearch, FoundX, FoundY, 6,  165,  159,  240,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Safari Status Check\Images\SUS.bmp
	if errorlevel = 2
		{
		MsgBox Unable to find needed images. Please copy it and restart the macro.
		ExitApp
		}
	if ErrorLevel = 0
		{
		Sleep 300
		WinClose
		Sleep 300
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep 300
		Send Claim is in Suspense
		Sleep 200
		Send {DOWN}{HOME}
		Sleep 300
		GoTo Beginning
		}
;Checks for Suspense Status
	ImageSearch, FoundX, FoundY, 6,  165,  159,  240,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Safari Status Check\Images\DREQ.bmp
	if errorlevel = 2
		{
		MsgBox Unable to find needed images. Please copy it and restart the macro.
		ExitApp
		}
	if ErrorLevel = 0
		{
		Sleep 300
		WinClose
		Sleep 300
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep 300
		Send Claim is in DREQ
		Sleep 200
		Send {DOWN}{HOME}
		Sleep 300
		GoTo Beginning
		}

;Checks for Invoiced Status
	ImageSearch, FoundX, FoundY, 6,  165,  159,  240,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Safari Status Check\Images\Inv.bmp
	if ErrorLevel = 0
		{
		Sleep 300
		WinClose
		Sleep 300
		WinWait, Microsoft Excel, 
		IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		WinWaitActive, Microsoft Excel, 
		Sleep 300
		Send Claim has invoiced
		Sleep 200
		Send {DOWN}{HOME}
		Sleep 300
		GoTo Beginning
		}

Send, {TAB}{DOWN}
Sleep 200

; checks for multiple refunds
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
Sleep 500
Send {ENTER}
Sleep 500

WinWaitActive, Patient Claim
ToolTip, Loading Patient Claim, 0,0
MouseClick, left,  281,  465
Sleep, 2000




IfWinExist, Recode Policy
	{
	Sleep 300
	WinWait, Recode Policy, 
	IfWinNotActive, Recode Policy, , WinActivate, Recode Policy, 
	WinWaitActive, Recode Policy, 
	Sleep 300
	Send {ENTER}
		;~ Sleep 100
		;~ WinWait, Microsoft Excel, 
		;~ IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
		;~ WinWaitActive, Microsoft Excel, 
		;~ Sleep, 300
		;~ Send Review
		;~ Sleep 300
		;~ Send {DOWN}{HOME}
		;~ Sleep 300
		;~ GoTo Beginning
	}

Send ^e
Sleep 500
loop,
{
	IfWinExist, Recode Policy
		break
	else
	{
		IfWinExist, Lookup Payor Contract
			break	
		else
			sleep, 500
	}
}
ToolTip
IfWinExist, Recode Policy
	{
	Sleep 300
	WinWait, Recode Policy, 
	IfWinNotActive, Recode Policy, , WinActivate, Recode Policy, 
	WinWaitActive, Recode Policy, 
	Send {ENTER}
	}
	
WinWait, Lookup Payor Contract, 
	IfWinNotActive, Lookup Payor Contract, , WinActivate, Lookup Payor Contract, 
	WinWaitActive, Lookup Payor Contract, 
	Sleep, 100
	Send, !f
	Sleep 100
	Send n
	Sleep 100

WinWait, Lookup Payor Contract, 
	IfWinNotActive, Lookup Payor Contract, , WinActivate, Lookup Payor Contract, 
	WinWaitActive, Lookup Payor Contract, 
	Sleep, 100
	Send {TAB 2}
	ControlFocus, ThunderRT6TextBox2, Lookup Payor
	send {home}400000208{ENTER}
	Sleep 200	
	Send {ENTER}
	Sleep 300
;~ WinWait, Report Code, 
;~ IfWinNotActive, Report Code, , WinActivate, Report Code, 
;~ WinWaitActive, Report Code, 
;~ Sleep 100
;~ Send %ReportCode%
;~ Sleep 100
;~ MouseClick, left,  272,  98
;~ Sleep, 100
;~ Send, {TAB}
;~ sleep, 100
;~ send, {ENTER}
;~ Sleep 1000
;pause

;~ loop,
;~ {
	;~ IfWinActive, Recode Policy
		;~ break
	;~ else
	;~ {
		;~ IfWinActive, Patient Claim
			;~ break
		;~ else
			;~ sleep, 200
	;~ }
;~ }


;~ IfWinExist, Recode Policy
	{
	WinWait, Recode Policy, 
	IfWinNotActive, Recode Policy, , WinActivate, Recode Policy, 
	WinWaitActive, Recode Policy, 
	Send {ENTER}
	}

WinWaitActive, Patient Claim, 
Sleep 500

;Special Fields Entry
SpecialFieldOrder:
WinWait, Patient Claim, 
	IfWinNotActive, Patient Claim, , WinActivate, Patient Claim, 
	WinWaitActive, Patient Claim, 
sleep, 100
;Check SF Order
ImageSearch, FoundX, FoundY, 545,  224, 714,  450, \\Aim\client_services$\MACROS- Chris Taylor\Matt\UHC\400000208sfOrder.bmp
	if ErrorLevel = 0
		GoTo Next
	else
		ImageSearch, FoundX, FoundY, 545,  224, 714,  450, \\Aim\client_services$\MACROS- Chris Taylor\Matt\UHC\400000208sfOrder2.bmp
			if ErrorLevel = 0
				GoTo Next
	
			;MsgBox %errorlevel%		
		
Sleep 100
MouseClick, left,  566,  251
send, {home}
Sleep, 1000
MouseClick, left,  552, 177
Sleep, 300

MouseClick, left,  675,  267
sleep, 100
ImageSearch, FoundX, FoundY, 545,  224, 714,  450, \\Aim\client_services$\MACROS- Chris Taylor\Matt\UHC\400000208sfOrder2.bmp
			if ErrorLevel = 0
				GoTo Next
MouseClick, left,  552, 177
GoTo SpecialFieldOrder

; adding SF
Next:
Sleep 300
MouseClick, left,  949,  352
MouseClick, left,  949,  352
;Send, .{ENTER}.{ENTER}{ENTER 2}
Sleep, 100
Send %SF1%{ENTER}
Sleep, 100
Send %SF2%{ENTER}
Sleep, 100
Send %SF3%{ENTER}





;Check for Blank SS#
ImageSearch, FoundX, FoundY, 194,311,319,355,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Brigett Recodes\SSN Blank.bmp
	if ErrorLevel = 0
	{
	Sleep, 100
	MouseClick, left,  273,  331
	Sleep, 100
	Send 111111111{TAB}
	Sleep, 100
	}
	if ErrorLevel = 1
	{
	Sleep, 500
	}


;Check for Blank Group#
ImageSearch, FoundX, FoundY, 9,  375,  326,  446,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Brigett Recodes\GroupNumberBlank.bmp
	if ErrorLevel = 0
	{
	Sleep, 100
	MouseClick, left,  193,  406
	Sleep, 100
	Send .
	Sleep, 100
	}
	if ErrorLevel = 1
	{
	Sleep, 500
	}



Send ^s
ToolTip, saving
Sleep 2000
;Sleep 500

/* 01.14.13
Loop
{
ImageSearch, FoundX, FoundY, 5,237,113,280,\\Aim\client_services$\MACROS- Chris Taylor\Matt\Safari Status Check\Images\Recoded.bmp
	if ErrorLevel = 0
	Break
	if ErrorLevel = 1
	{
	Sleep 500
	}
}
*/


; Claim Detail fix 7.31.12
				
	;Send, ^s

sleep, 200
IfWinExist, Save Changes
	send, y
else 

IfWinExist Missing Fields
	{
	WinClose
		Sleep 100
	IfWinExist, Failed V
		Send {enter}
	
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 300
	Send missing field{DOWN}{HOME}{LEFT}
	continue
	}
	
sleep 1000
	WinWait, Claim ID,,10
/*
; confirms In House and RET are set

				WinWaitActive, Patient Claim Detail
				ToolTip
				Sleep, 1000
				Send, {Tab}
				Sleep, 100
				Send, I ; IH/In House 
				Sleep, 100
				Send, {Tab 2} 
				Sleep, 500
				Send, R ; RET/Retraction
				Sleep, 100
				Send, {Tab} 
				Sleep, 500
				Send, !fv  ; Save and Close
				WinWaitClose, Patient Claim Detail
*/



	Sleep 100
	WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 300
	Send x{DOWN}{HOME}{LEFT}

}
ToolTip
MsgBox Done
Esc::
ExitApp
Pause::Pause

]::
{
Send {HOME}
}
Reload