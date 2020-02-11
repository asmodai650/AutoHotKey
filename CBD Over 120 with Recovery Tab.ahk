#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect.  Column A should contain the Account Resolution ID, and column B New Request note of "Claim will be closed via MACRO.".
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
	MsgBox, Ensure Excel is open.
	GoTo, CHECK	
	}

FormatTime, TimeBegin,, Time

Loop, %Count%
{
TrayTip,,%a_index% of %Count%,30	

Account =
;Note should state "Claim will be closed via MACRO."
Note = 

Clipboard = 

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 750
Send, +{Home}
Sleep, 750
Send, ^c
ClipWait, 2
StringSplit, MyArray, clipboard, %A_Tab%
Account = %MyArray1% 
Note = %MyArray2%

If Account =
{
Msgbox The macro is complete.
ExitApp
}

	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		;else
		;{
			;break
		;}
	}

	WinActivate, Inventory
	IfWinNotExist, Inventory
	{
		;MsgBox, Confirm that the correct Direct Connect Window is open.
		WinActivate Error
			IfWinNotExist, Error
			{
				;MsgBox, Confirm that the correct Direct Connect Window is open.
				WinActivate Partner
					IfWinNotExist, Partner Account View
						;MsgBox, Confirm that the correct Direct Connect Window is open.
						WinActivate Account View
						{
							WinActivate Internet Explorer
							clipboard = 
							sleep, 2000
							send, {Home 2}
							Sleep, 7500
							mousemove, 185, 46
							Sleep, 750
							Click, 185, 46
							Sleep, 750
							comment = https://directconnect.optum.com/Inventory/MyInventory
							sleep, 200
							send, %comment%
							sleep, 200
							send, {enter}
							sleep, 200
							WinWaitActive Internet Explorer
							sleep, 2000
							Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
							Sleep, 750
							MouseMove, 20,125 ; Close Find Window
							Sleep, 750
							Click, 20,125 ; Close Find Window
							Sleep, 750
							MouseMove, 870,200 ; Inventory button
							Sleep, 7500
							Click
							Sleep, 1000
							Clipboard = 
							GoTo, EXCEL2 ; restarts macro with same claim
							
							;Sleep, 1000
							;send, {Home}
							;sleep, 1000
							;send, {Home}
							;sleep, 1000
							;Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
							;Sleep, 200
							;MouseMove, 20,125 ; Close Find Window
							;Sleep, 750
							;Click, 20,125 ; Close Find Window
							;Sleep, 750
							;send, {Home}
							;sleep, 1000
							;send, {Home}
							;sleep, 1000
							;MouseMove, 870,200 ; Inventory button
							;Sleep, 7500
							;Click
							;Sleep, 1000
						}
			}
			else
			{
				WinActivate Internet Explorer
				clipboard = 
				sleep, 2000
				send, {Home 2}
				Sleep, 7500
				mousemove, 185, 46
				Sleep, 750
				Click, 185, 46
				Sleep, 750
				comment = https://directconnect.optum.com/Inventory/MyInventory
				sleep, 200
				send, %comment%
				sleep, 200
				send, {enter}
				sleep, 200
				WinWaitActive Internet Explorer
				sleep, 2000
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 750
				MouseMove, 20,125 ; Close Find Window
				Sleep, 750
				Click, 20,125 ; Close Find Window
				Sleep, 750
				MouseMove, 870,200 ; Inventory button
				Sleep, 7500
				Click,
				Sleep, 1000
				Clipboard = 
				GoTo, EXCEL2 ; restarts macro with same claim
			}
	}
	WinActivate, Inventory
	MouseMove, 870, 200 ; Inventory button
	Sleep, 750
	Click,  870, 240 ; Search button
	sleep, 1000

; Wait for page to load then go to left side
;Click to select type of claim number
mousemove, 118, 315 
Sleep, 750
Click 2
Sleep, 750
;Go up to Resolution ID and Tab over to search field
Send, {r} 
Sleep, 750
Send, {Tab}
Sleep, 750
Send, %Account%
Sleep, 750
Send, {Enter}
Sleep, 1500

;Check to see what claim status is
Sleep, 1500
Loop,
	{
	;mouseclickdrag, left, 450, 427, 508, 427
	mouseclickdrag, left, 850, 440, 900, 440
	Send, ^c
	ClipWait, 2
		
		If clipboard contains In Process
			{
				;msgbox, Claim In Process
				Sleep, 200
				MouseMove, 867, 436
				Sleep, 200
				Click 2
				Sleep, 2000
				Clipboard =
				GoTo, REQUEST
			}
		; If claim is not In Process, look to see if it is resolved
		else
			If clipboard contains Resolved
			Loop,
				{
					Sleep, 200
					MouseMove, 867, 436
					Sleep, 200
					Click 2
					Sleep, 2000
					Clipboard =
					Sleep, 1000
					MouseMove, 100, 275
					Sleep, 750
					Click 2
					Send, ^c
					ClipWait, 2
					{
					; Reopen account
					If clipboard contains Account
						Sleep, 2000
						Send, {TAB}
						Sleep, 750
						Send, {Enter}
						Sleep, 7500
						GoTo, REQUEST
						Break
					}
				}
			Else
				{
				If Clipboard = 
					{
					;MsgBox, UID Not Found
					MouseMove, 870, 200 ; Inventory button
					Sleep, 750
					Click,  870, 240 ; Search button
					Sleep, 750
					comment = UID Not Found
					goto, EXCEL
					Sleep, 1000
					Break
					}
				else
					{
					;MsgBox, UID Not Resolved
					MouseMove, 870, 200 ; Inventory button
					Sleep, 750
					Click,  870, 240 ; Search button
					Sleep, 750
					comment = %clipboard%
					goto, EXCEL
					Sleep, 1000
					Break
					}
				}
	}

; Go to New Request and send comment to Provider stating claim is being closed through a macro
;Open New Request, Category = Information

REQUEST:

Sleep, 1000
Send, ^a
Sleep, 250
Send, ^c
ClipWait, 2
Loop,
	{
		If clipboard contains INACTIVE
		{
			Sleep, 750
			goto, RESOLVE
			Sleep, 750
			break
		}
		Else
			Clipboard =
			Sleep, 1000
			break
	}

Sleep, 250
MouseMove, 50, 240
Sleep, 750
click
Sleep, 250

Loop,
	{
		mousemove, 100, 275
		Sleep, 250
		click 2
		Send, ^c
		ClipWait, 2
		If clipboard contains Account
		{
			Sleep, 2000
			Send, ^f
			Sleep, 750
			Comment = New Request
			Send, %Comment%
			Sleep, 750
			Send, {Tab}
			sleep, 250
			;Send, {Enter}
			;sleep, 250
			Send, {Tab}
			sleep, 250
			Send, +{Tab}
			sleep, 250
			Send, {Enter}
			Sleep, 750
			Break
		}
		Else
			If clipboard contains = An error
			sleep = 1000
				{
					;MsgBox, Search Inventory page loaded.
					clipboard = 
					sleep, 1000
					MouseMove, 950,205 ; Inventory button
					Sleep, 750
					Click, 988, 247 ; Search button
					Sleep, 750
					Comment = An Error Has Occurred
					GoTo, EXCEL
					Break
				}
	
	}
	

;New Request Subject = Please Review
Sleep, 750
Send, {TAB}
Sleep, 250
Send, {i}
Sleep, 250
Send, {TAB}
sleep, 250
send, {down 3}
Sleep, 250
Send, {TAB 3}
Sleep, 250
;New Request Comment = Claim will be closed via MACRO.
Send, %MyArray2%
Sleep, 750
;Save New Request, NOT save and close
Send, {TAB}

;Wait for Direct Connect refresh every 15 minutes
	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		;else
		;{
			;break
		;}
	}

Sleep, 750
Send, {Enter}
Sleep, 2000

;Open Resolve Window
Sleep, 250
; NEW COMMAND BELOW
MouseMove, 20,125 ; Close Find Window
Sleep, 750
Click ; Close Find Window
Sleep, 750

Loop,
	{
	;sleep, 1000
	mousemove, 100, 275
	Sleep, 250
	click 2
	Sleep, 250
	Send, ^c
	ClipWait, 2
	
		If clipboard contains Account
			{
			Sleep, 750
			goto, RESOLVE
			Break
			}
		Else
			clipboard = An error has occurred
			sleep = 1000
				{
				;MsgBox, Search Inventory page loaded.
				clipboard = 
				sleep, 1000
				MouseMove, 950,205 ; Inventory button
				Sleep, 750
				Click ; Search button
				Sleep, 750
				Comment = An Error Has Occurred
				GoTo, EXCEL
				}
	}

;Wait for Direct Connect refresh every 15 minutes
	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		;else
		;{
			;break
		;}
	}

RESOLVE:
clipboard = 
Sleep, 750
;mousemove, 100, 275
;Sleep, 250
;Click 2
;Sleep, 250
;Send, ^c
;ClipWait, 2
;Loop,
	;{
		;If clipboard contains Account 
		;{
			;Sleep, 2000
			;Send, {TAB}
			;Sleep, 750
			;Send, {Enter}
			;Sleep, 7500
			;Break
		;}
		;else
		;{
			;MsgBox, Search Inventory page loaded.
			;clipboard = 
			;sleep, 1000
			;Send, {home 2}
			;Sleep, 750
			;MouseMove, 950,205 ; Inventory button
			;Sleep, 750
			;Click ; Search button
			;Sleep, 750
			;Comment = An Error Has Occurred
			;GoTo, EXCEL
		;}
	;}
;Wait for Direct Connect refresh every 15 minutes
	Loop, 5
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000
			SplashTextOff
		}
		else
		{
			break
		}
	}

;Resolve Reason = Created in Error
mousemove, 100, 275
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {enter}
Sleep, 500
Send, {TAB}
Sleep, 250
Send, c ;CHANGE WHERE NEEDED! C = CREATED IN ERROR / W = WRITE OFF / D = DISAPPROVED / A = APPROVED / P = PREVIOUSLY PROCESSED
Sleep, 250
Send, {TAB}
Sleep, 750
Send, {Enter}
sleep, 1000
Clipboard = 
	
;Determine if Direct Connect loaded correctly
	Sleep, 750
	Send, {home}
	Sleep, 1000
	mousemove, 100, 275
	Sleep, 250
	Click 2
	Sleep, 250
	Send, ^c
	ClipWait, 2 

	If clipboard contains Account
	{
		;Close claim and return to main inventory screen
			;MsgBox, Search Inventory page loaded.
			clipboard = 
			sleep, 1000
			Send, {pgup 2}
			Sleep, 750
			MouseMove, 870,200 ; Inventory button
			Sleep, 750
			Click ; Inventory button
			Sleep, 750
			Clipboard= 
			Comment = Account Resolved
			GoTo, EXCEL
		;Close claim and return to main inventory screen
			;Sleep, 750
			;Send, {TAB}
			;Sleep, 750
			;Send, {Enter}
			;sleep, 1000
			;Send, {Home}
			;WinWaitActive Inventory
			;Sleep, 1000
			;clipboard =
			;Comment = Account Resolved
			;Sleep, 250
			;Goto, EXCEL
	}
	Else
	{
		sleep, 2000
		send, {Home 2}
		sleep, 1000
		Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
		Sleep, 200
		MouseMove, 20,125 ; Close Find Window
		Sleep, 750
		Click, 20,125 ; Close Find Window
		Sleep, 750
		MouseMove, 870,200 ; Inventory button
		Sleep, 1000
		Clipboard = 
		Comment = Double Check Claim
		GoTo, EXCEL
	}

EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 750
Send, %Comment%
Sleep, 750
Send, {Down}
Sleep, 750

EXCEL2:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 1000

}

FormatTime, TimeEnd,, Time
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause