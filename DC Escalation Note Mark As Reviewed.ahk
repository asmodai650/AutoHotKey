#SingleInstance, Force
SetTitleMatchMode, 2


MsgBox, In your Excel workbook, select the cell in column A on to the account you want to start with. Now log into Direct Connect and select UnitedHealthcare (Payer) from the provider list. The Excel Columns should be as follows: Column A - Unique ID and Column B - Worklist Note.

InputBox, Count, How many accounts?, How many accounts do you want to add the note to?

CHECK:
IfWinNotExist, Microsoft Excel
{
	MsgBox, Ensure Excel is open.
	GoTo, CHECK	
}

FormatTime, TimeBegin,, Time

Loop, %Count% ;This is the main loop that should go just below Format Time
{
	TrayTip,,%a_index% of %Count%,30 ;Clear clipboard, copy data from excel, and go to Direct Connect

	Clipboard =

	UNIQUEID =
	NOTE =

	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel

	Sleep, 500
	Send, +{Space}
	Sleep, 500
	Send, ^c
	ClipWait, 1
	StringSplit, MyArray, clipboard, %A_Tab%
	UNIQUEID = %MyArray1% 
	NOTE = %MyArray2%

	If UNIQUEID = ;If no UID in Excel Cell, end Macro
	{
		Msgbox The macro is complete.
		ExitApp
	}

	Loop, 5 ;Pause script if minutes are 00, 15, 30, 45 to compensate for direct connect time-out
	{
		if A_Min = 00
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000 	;Message box will close automatically after 2 seconds
			SplashTextOff
		}
		if A_Min = 15
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000 	;Message box will close automatically after 2 seconds
			SplashTextOff
		}
		if A_Min = 30
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000 	;Message box will close automatically after 2 seconds
			SplashTextOff
		}
		if A_Min = 45
		{
			SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
			sleep, 15000 	;Message box will close automatically after 2 seconds
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
				MouseMove, 978,200 ; Inventory button
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
				MouseMove, 978,200 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
	}
	WinActivate, Inventory


	MouseMove, 978, 200 ; Inventory button
	Sleep, 500
	Click,  978, 240 ; Search button
	sleep, 2500

	Loop,
	{
		MouseMove, 66, 278
		Sleep, 1000
		click 2
		Sleep, 1000
		Send, ^c
		clipwait, 1
		
		If clipboard contains Search ;Look for Search, if not found, reload Direct Connect Page
		{
			Break
		}
		Else
		{
			Sleep, 2000
			Send, {Home 2}
			Sleep, 200
			Send, ^f
			Sleep, 200
			MouseMove, 20,125
			Sleep, 200
			Click, 20,125 ; Close Find Window
			sleep, 200
			send, {Home 2} ; Send page back to top
			;NEW TEST
			Sleep, 500
			Send, ^f
			Sleep, 500
			Comment = Request / Response
			Send, %Comment%
			sleep, 500
			Send, {Tab 2}
			Sleep, 500
			Send, {Enter}
			Sleep, 200
			MouseMove, 20,125
			Sleep, 200
			Click, 20,125 ; Close Find Window
			sleep, 200
			send, {Home 2} ; Send page back to top
			;END NEW TEST
			sleep, 1000
			MouseMove, 978,200 ; Inventory button
			Sleep, 500
			Click, 988, 247 ; Search button
			Sleep, 500
		}
	}

	Sleep, 500
	Send, {Tab} ;Activates the claim type drop-down field
	Sleep, 500
	Send, {u} ;Goes up to Unique ID
	Sleep, 500
	Send, {Tab} ;Tab over to search field
	Sleep, 500
	Send, %UNIQUEID% ;Enters Unique ID from Excel
	Sleep, 500
	Send, {Enter}
	Sleep, 1000

; Wait for page to load
;Determine UNIQUEID Status. If In Process goto REQUEST. If anything else, go back to excel.
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
		If Clipboard = 
		{
			;MsgBox, UNIQUEID Not Found
			MouseMove, 978, 200 ; Inventory button
			Sleep, 500
			Click,  988, 247 ; Search button
			sleep, 500
			comment = UNIQUEID Not Found
			goto, EXCEL
		}
		Else
		{
			;MsgBox, UNIQUEID Not In Process
			MouseMove, 978, 200 ; Inventory button
			Sleep, 500
			Click,  988, 247 ; Search button
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

;Look for Open Requests:
	If clipboard contains Open Requests:
	{
		;Goes to Please Review request
		Sleep, 1000
		Send, {home}
		Sleep, 500
		Send, ^f
		Sleep, 500
		Comment = Open Requests:
		Send, %Comment%
		sleep, 500
		Send, {Tab 2}
		;sleep, 250
		;Send, +{Tab}
		Sleep, 250
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
		
		;Opens Please Review window
		Send, {Enter}
		sleep, 250
		Clipboard = 
		GoTo, POPUPWINDOWCHECK2
	}
	Else
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
		MouseMove, 978,200 ; Inventory button
		Sleep, 1000
		Clipboard = 
		Comment = Double Check Claim
		GoTo, EXCEL
	}

POPUPWINDOWCHECK2:

;Checks to make sure New Request window is open
	Sleep, 5000
	MouseMove, 540, 380
	Sleep, 500
	Click 2
	Sleep, 200
	Send, ^c
	clipwait, 2
	
	If clipboard = Review
	{
		Sleep, 1000
		Goto REVIEWED
	}
	else
	{
		Sleep, 500
		;MouseMove, 916, 328
		MouseMove, 915, 340
		Sleep, 200
		Click
		Sleep, 500
		Send, {Home 2}
		Sleep, 500
		Clipboard = 
		Comment = Please Review Not Available
		GoTo, EXCEL
	}

REVIEWED:

	Sleep, 500
	Send, {Tab 2}
	Sleep, 500
	Send, {r}
	Sleep, 500
	Send, {Tab 3}
	Sleep, 500
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
		
	;Opens Please Review window
	Send, {Enter}
	Sleep, 2000

;go to the top of the browser window, close the find window, and determine if Direct Connect loaded correctly
	Sleep, 1000
	Send, {home}
	Sleep, 250
	Send, {home}
	Sleep, 250
	MouseMove, 20,125 
	Sleep, 250
	Click ; Close Find Window

	Sleep, 500	
	mousemove, 80, 275
	Sleep, 250
	Click 2
	Sleep, 250
	Send, ^c
	ClipWait, 2 

	If clipboard contains Search
		{
			;Close claim and return to main inventory screen
			Sleep, 500
			Send, {Home}
			WinWaitActive Inventory
			Sleep, 500
			clipboard =
			Comment = Claim Reviewed
			Sleep, 500
			Goto, EXCEL
		}
	Else
		{
			sleep, 500
			send, {Home 2}
			sleep, 1000
			Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
			Sleep, 500
			MouseMove, 20,125 ; Close Find Window
			Sleep, 500
			Click, 20,125 ; Close Find Window
			Sleep, 500
			MouseMove, 978,200 ; Inventory button
			Sleep, 500
			Clipboard = 
			Comment = Double Check Claim
			GoTo, EXCEL
		}


EXCEL:

;Update Excel with comment based on actions above
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	Sleep, 500
	Send, {Tab 2}
	Sleep, 500
	Send, %Comment%
	Sleep, 500
	Send, {Down}
	Sleep, 500
	Send, {home}
	Sleep, 500
}

FormatTime, TimeEnd,, Time
MsgBox, %Count% UNIQUEID's resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
Esc::Pause