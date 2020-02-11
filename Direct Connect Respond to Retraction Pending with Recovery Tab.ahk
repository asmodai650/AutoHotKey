#SingleInstance, Force
SetTitleMatchMode, 2

MsgBox, In your Excel workbook, select the cell in Column D next to the UID you want to start with. Column A should contain UID, column B contains the Request Subject, and column C contains the Macro Comment. Now log into Direct Connect under UnitedHealthcare.

InputBox, Count, How many UID's?, How many UID's do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
		;MsgBox, Ensure Excel is open.
		GoTo, CHECK	
	}


FormatTime, TimeBegin,, Time

Loop, %Count%
{
TrayTip,,%a_index% of %Count%,30	
Clipboard =

UID =
Comment =

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
	Sleep, 500
	Send, +{Home}
	Sleep, 500
	Send, ^c
	ClipWait, 1
StringSplit, MyArray, clipboard, %A_Tab%
	UID = %MyArray1% 
	Subject = %MyArray2%	
	Comment = %MyArray3%

If UID =
{
Msgbox The macro is complete.
ExitApp
}

	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 61000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 46000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 36000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 16000
			SplashTextOff
		}
		else
		{
			break
		}
	}
	
	WinActivate, Inventory
	IfWinNotExist, Inventory
	{
		;MsgBox, Confirm that the correct Direct Connect Window is open.
		WinActivate Error
			IfWinNotExist, Error
			{
				;MsgBox, Confirm that the correct Direct Connect Window is open.
				WinActivate Partner Account View
				Sleep, 1000
				MouseMove, 916, 358
				Sleep, 200
				Click
				sleep, 2000
				send, {Home 2}
				sleep, 200
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 200
				MouseMove, 20,125 ; Close Find Window
				Sleep, 500
				Click, 20,125 ; Close Find Window
				Sleep, 500
				MouseMove, 860,200 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
			else
			{
				sleep, 2000
				send, {Home 2}
				sleep, 200
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 200
				MouseMove, 20,125 ; Close Find Window
				Sleep, 500
				Click, 20,125 ; Close Find Window
				Sleep, 500
				MouseMove, 860,200 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
	}
	WinActivate, Inventory
	MouseMove, 860, 200 ; Inventory button
	Sleep, 500
	Click,  890, 250 ; Search button
	sleep, 500

sleep, 1000

Loop,
	{
	mouseclickdrag, left, 37, 278, 215, 278
	Send, ^c
	clipwait, 1
		If clipboard contains Search Inventory
			{
			;MsgBox, Search Inventory page loaded.
			Break
			}
		Else
		;If clipboard contains An error has occurred
			{
			;MsgBox, Search Inventory page loaded.
			sleep, 2000
			send, {Home}
			sleep, 200
			Send, {Home}
			MouseMove, 860,200 ; Inventory button
			Sleep, 500
			Click, 890, 250 ; Search button
			Sleep, 500
			}
	}
			 
Sleep, 500
Send, {Tab} ;Activates the claim type drop-down field
Sleep, 500
Send, {u} ;Go up to Unique ID
Sleep, 500
Send, {Tab} ;Tab over to search field
Sleep, 500
Send, %UID%
Sleep, 500
Send, {Enter}
Sleep, 2000

; Wait for page to load

;Determine RESID Status. If Resolved, re-open and goto REQUEST. If anything else, go back to excel.
	Loop, 
	{
		Clipboard = 
		sleep, 200
		mouseclickdrag, left, 848, 436, 918, 436
		send, ^c
		clipwait, 1
		
		If clipboard contains In Process
		{
			;this clicks hyper link for acct
			Sleep, 200
			MouseMove, 867, 436
			Sleep, 200
			Click 2
			Sleep, 2000
			Clipboard =
			break
		}
		If clipboard contains Resolved
		{
			;this clicks hyper link for acct
			Sleep, 200
			MouseMove, 867, 436
			Sleep, 200
			Click 2
			Sleep, 2000
			Clipboard =
			
			;Re-Open account
			mouseclickdrag, left, 45, 278, 189, 278
			Sleep, 500
			Send, ^c
			clipwait, 1
			
			If clipboard contains Account View
			;Goes to new request button
				Sleep, 1000
				MouseMove, 784, 253
				Sleep, 200
				Click 1
				Sleep, 2000
			break
		}
		If Clipboard = 
		{
			;MsgBox, RESID Not Found
			MouseMove, 860, 200 ; Inventory button
			Sleep, 500
			Click,  890, 250 ; Search button
			sleep, 500
			comment = RESID Not Found
			goto, EXCEL
		}
		Else
		{
			;MsgBox, RESID Not Resolved
			MouseMove, 860, 200 ; Inventory button
			Sleep, 500
			Click,  890, 250 ; Search button
			sleep, 500
			comment = %clipboard% ; copies whatever text was selected and will paste it in Excel
			goto, EXCEL
		}
	}


;Check page to see if Open Requests: is an option. If found, continue. If NOT found, return to Excel.
	WinWaitActive, Partner Account View
	IfWinNotExist, Partner Account View
	{
		MsgBox, Confirm that the correct Direct Connect Window is open.
	}

mouseclickdrag, left, 45, 278, 189, 278
Sleep, 500
Send, ^c
clipwait, 1

If clipboard contains Account View
	;Goes to new request button
		Sleep, 2000
		Send, ^f
		Sleep, 500
		Comment = New Request
		Send, %Comment%
		sleep, 500
		Send, {Tab}
		sleep, 250
		Send, {Enter}
		sleep, 250
		Send, {Tab}
		sleep, 250
		Send, +{Tab}
		sleep, 250
		Send, {Enter}
		sleep, 1000
		goto, REQUEST

REQUEST:
	
	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 61000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 46000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 36000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 16000
			SplashTextOff
		}
		else
		{
			break
		}
	}	
	
	;Opens new request window
	Sleep, 1000
	Send, {Tab}
	Sleep, 200
	Send, {f}
	Sleep, 200
	Send, {tab}
	Sleep, 200
	Send, %MyArray2%
	Sleep, 200
	Send, {tab}
	Sleep, 200
	Send, {tab}
	Sleep, 200
	Send, {tab}
	Sleep, 200
	Send, %MyArray3%
	Sleep, 200
	Send, {tab}
	Sleep, 200
	Send, {tab}
	Sleep, 200
	
;pauses script for Direct Connect refresh every 15 minutes
	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 61000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 46000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 36000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 16000
			SplashTextOff
		}
		else
		{
			break
		}
	}

	; closes new request window
	Send, {Enter}
	Sleep, 2000

	Clipboard = 
	
;Determine if Direct Connect loaded correctly
	Sleep, 500
	Send, {home}
	Sleep, 500
	MouseMove, 72, 306
	Sleep, 500
	click 2
	Sleep, 250
	Send, ^c
	clipwait, 1

	If clipboard contains Search
	{
		MouseMove, 20,125 ; Close Find Window
		Sleep, 500
		Click, 20,125 ; Close Find Window
		Sleep, 500
		Clipboard = 
		Comment = Request Entered
		Goto, EXCEL
	}
	Else
	{
		sleep, 2000
		send, {Home 2}
		sleep, 1000
		Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
		Sleep, 200
		MouseMove, 20,125 ; Close Find Window
		Sleep, 500
		Click, 20,125 ; Close Find Window
		Sleep, 500
		MouseMove, 860,200 ; Inventory button
		Sleep, 1000
		Clipboard = 
		Comment = Double Check Claim
		GoTo, EXCEL
	}

;Update Excel with comment based on actions above
EXCEL:
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	clipboard = 
	Sleep, 500
	Send, %Comment%
	Sleep, 500
	Send, {Down}
	Sleep, 500
	clipboard = 
}


FormatTime, TimeEnd,, Time
MsgBox, %Count% UID's resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
pause::Pause
ESC::Pause