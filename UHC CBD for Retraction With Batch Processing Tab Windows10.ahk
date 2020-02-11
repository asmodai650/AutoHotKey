#SingleInstance, Force
SetTitleMatchMode, 2


MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect and select UnitedHealthcare (Payer) from the provider list. The Excel Columns should be as follows: Column A - Resolution ID and Column B - Comment for Open Request 1.

InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
{
	MsgBox, Ensure Excel is open.
	GoTo, CHECK	
}

;IfWinNotExist,  Inventory
;{
;	MsgBox, Confirm that the correct Direct Connect Window is open.
;	WinActivate Internet Explorer
;	WinWaitActive Inventory
;	GoTo, CHECK
;}

FormatTime, TimeBegin,, Time

Loop, %Count% ;This is the main loop that should go just above Format Time
{
	TrayTip,,%a_index% of %Count%,30

	Clipboard =

	RESID =
	RequestResponseA =

	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel

	Sleep, 500
	Send, {Home}{Right 2}
	Send, +{Space}
	Sleep, 500
	Send, ^c
	ClipWait, 1
	StringSplit, MyArray, clipboard, %A_Tab%
	RESID = %MyArray1% 
	RequestResponseA = %MyArray2%

	If RESID =
	{
		Msgbox The macro is complete.
		ExitApp
	}

	;Pause script if minutes are 00, 15, 30, 45 to compensate for direct connect time-out
	; Message box will close automatically after 2 seconds

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
				MouseMove, 20, 125 ; Close Find Window
				Sleep, 500
				Click, 20, 125 ; Close Find Window
				Sleep, 500
				MouseMove, 872, 207 ; Inventory button
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
				MouseMove, 20, 125 ; Close Find Window
				Sleep, 500
				Click, 20, 125 ; Close Find Window
				Sleep, 500
				MouseMove, 872, 207 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
	}
	WinActivate, Inventory


	MouseMove, 872, 207 ; Inventory button
	Sleep, 500
	Click,  880, 246 ; Search button
	sleep, 2500

	Loop,
	{
		MouseMove, 66, 278
		Sleep, 1000
		click 2
		Sleep, 1000
		Send, ^c
		clipwait, 1
		
		If clipboard contains Search
		{
			Break
		}
		Else
		{
			Sleep, 200
			Send, ^f
			Sleep, 200
			MouseMove, 20, 125
			Sleep, 200
			Click, 20, 125 ; Close Find Window
			sleep, 200
			send, {Home 2} ; Send page back to top
			sleep, 500
			MouseMove, 872, 207 ; Inventory button
			Sleep, 500
			Click, 880, 246 ; Search button
			Sleep, 500
		}
	}

	Sleep, 500
	Send, {Tab} ;Activates the claim type drop-down field
	Sleep, 500
	Send, {r} ;Goes up to Resolution ID
	;Send, {u} ;Goes up to Unique ID	
	Sleep, 500
	Send, {Tab} ;Tab over to search field
	Sleep, 500
	Send, %RESID% ;Enters RES ID from Excel
	Sleep, 500
;pause
	Send, {Enter}
	Sleep, 1000

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
			break
		}
		If clipboard contains Resolved
		{
			break
		}
		If Clipboard = 
		{
			MsgBox, RESID Not Found
			MouseMove, 872, 207 ; Inventory button
			Sleep, 500
			Click,  880, 246 ; Search button
			sleep, 500
			comment = RESID Not Found
			goto, EXCEL
		}
		Else
		{
			MsgBox, RESID Not Resolved
			MouseMove, 872, 207 ; Inventory button
			Sleep, 500
			Click,  880, 246 ; Search button
			sleep, 500
			comment = %clipboard% ; copies whatever text was selected and will paste it in Excel
			goto, EXCEL
		}
	}

;this clicks hyper link for acct
	Sleep, 200
	MouseMove, 867, 436
	Sleep, 200
	Click 2
	Sleep, 2000

	Clipboard =

;Check page to see if Open Requests: is an option. If found, continue. If NOT found, return to Excel.
	WinWaitActive, Partner Account View
	IfWinNotExist, Partner
	{
		MsgBox, Confirm that the correct Direct Connect Window is open.
	}
	Sleep, 1000
	Send, {end 2}
	Sleep, 1000
	Send, ^a
	Sleep, 1000
	Send, ^c
	ClipWait, 10, 1

	Loop,
	{
		If clipboard contains Open Requests:
		{
			Sleep, 500
			Send, {home 2}
			MouseMove, 14, 278
			Sleep, 500
			Click
			Sleep, 1000
			break
		}
		else
		{
			Sleep, 500
			Send, {home 2}
			Sleep, 1000
			MouseMove, 14, 278
			Sleep, 500
			Click
			MouseMove, 872, 207 ; Inventory button
			Sleep, 500
			Click,  880, 246 ; Search button
			sleep, 500
			Clipboard = 
			Comment = Open Request Not Found
			Goto, EXCEL
		}
	}
	Loop,
	{
		mouseMove, 83, 278
		Sleep, 500
		Click 2
		Sleep, 200
		Send, ^c
		clipwait, 1
		
		If clipboard contains Account
		{
			;Goes to new request button
			Sleep, 2000
			Send, ^f
			Sleep, 500
			Comment = Open Requests:
			Send, %Comment%
			sleep, 500
			Send, {Tab 2}
			sleep, 250
			Send, {Enter}
			sleep, 250
			;pauses script for Direct Connect refresh every 15 minutes
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
			Clipboard = 
			goto, POPUPWINDOWCHECK
		}
	}

POPUPWINDOWCHECK:

;Opens new request window
	Sleep, 5000
	;MouseMove, 582, 385
	MouseMove, 590,380	
	Sleep, 500
	Click 2
	Sleep, 200
	Send, ^c
	clipwait, 2
	
	If clipboard = Retraction
	{
		Sleep, 1000
		Goto RETRACTION
	}
	else
	{
		Sleep, 500
		MouseMove, 916, 339
		;MouseMove, 590,380
		Sleep, 200
		Click
		Sleep, 1000
		Goto POPUPWINDOWCHECK2
	}

RETRACTION:

;If Retraction was found in the pop-up window the script continues here
	Sleep, 250
	Send, {tab 2}
	Sleep, 250
	Send, {r}
	Sleep, 250
	Send, {tab}
	Sleep, 250
	Send, %RequestResponseA%
	Sleep, 250
	Send, {Tab 2}
	sleep, 250
	;pauses script for Direct Connect refresh every 15 minutes
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
			break ;If not paused for one of the minutes above, continue
		}
	}

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
		MouseMove, 20, 125 ; Close Find Window
		Sleep, 500
		Click, 20, 125 ; Close Find Window
		Sleep, 500
		Clipboard = 
		Comment = Request Entered
		Goto, EXCEL
	}
	Else
	{
		sleep, 2000
		send, {Home 2}
		sleep, 200
		Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
		Sleep, 200
		MouseMove, 20, 125 ; Close Find Window
		Sleep, 500
		Click, 20, 125 ; Close Find Window
		Sleep, 500
		MouseMove, 872, 207 ; Inventory button
		Sleep, 1000
		Clipboard = 
		Comment = Double Check Claim
		GoTo, EXCEL
	}

POPUPWINDOWCHECK2:

;If Retraction was NOT found in the pop-up window the script continues here 
	Clipboard = 
	Sleep, 1000
	Send, {home 2}
	Sleep, 500
	Send, ^f
	Sleep, 500
	Comment = Open Requests:
	Send, %Comment%
	sleep, 500
	Send, {Tab 3}
	sleep, 250
	Send, {Enter}
	sleep, 1000
	;MouseMove 582 , 399
MouseMove, 590,380
	Sleep, 500
	Click 2
	Sleep, 250
	Send, ^c
	clipwait, 1
	Loop, ;determine if second popup window can be used
	{
		If clipboard NOT contains Retraction
		{
			Sleep, 500
			Goto, RETRACTIONNOTFOUND
		}
		If clipboard = Retraction
		{
			Sleep, 200
			Clipboard = 
			Send, ^a
			Sleep, 200
			Send, ^c
			ClipWait, 2
			If Clipboard Contains Responded by
			{
				Sleep, 500
				Goto, RETRACTIONNOTFOUND
			}
			else
			{
				Sleep, 500
				MouseMove, 635,366
				Sleep, 200
				Click
				Sleep, 200
				MouseMove, 582, 399
				Sleep, 500
				Click 2
				Sleep, 200
				Send, ^c
				clipwait, 2
				If clipboard = Retraction
					{
						Sleep, 500
						Goto RETRACTION
					}
				else
					{
						Sleep, 500
						Goto, RETRACTIONNOTFOUND
					}
			}
		}
	}

RETRACTIONNOTFOUND:

;Unable to find Retraction Request for Res ID
	Sleep, 1000
	MouseMove, 916, 342
	Sleep, 200
	Click
	Sleep, 1000
	send, {Home 2}
	sleep, 500
	Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
	Sleep, 200
	MouseMove, 20, 125 ; Close Find Window
	Sleep, 500
	Click, 20, 125 ; Close Find Window
	Sleep, 500
	MouseMove, 872, 207 ; Inventory button
	Sleep, 1000
	Clipboard = 
	Comment = Retraction Request Not Found
	GoTo, EXCEL

EXCEL:

;Update Excel with comment based on actions above
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	Sleep, 1000
	Send, %Comment%
	Sleep, 500
	Send, {Down}
	Sleep, 500
}

FormatTime, TimeEnd,, Time
MsgBox, %Count% RESID's resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
Esc::Pause