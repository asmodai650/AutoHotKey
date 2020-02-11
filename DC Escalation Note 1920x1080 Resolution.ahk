#SingleInstance, Force
SetTitleMatchMode, 2


MsgBox, In your Excel workbook, select the cell in column A that has the account you want to start with. Now log into Direct Connect and select UnitedHealthcare (Payer) from the provider list. The Excel Columns should be as follows: Column A - Unique ID and Column B - Worklist Note.

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

	Sleep, 250
	Send, +{Space}
	Sleep, 250
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
				sleep, 1000
				send, {Home 2}
				sleep, 200
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 200
				MouseMove, 20,1 25 ; Close Find Window
				Sleep, 500
				Click, 20, 125 ; Close Find Window
				Sleep, 500
				MouseMove, 1515, 200 ; Inventory button
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
				;MouseMove, 978, 200 ; Inventory button
				MouseMove, 1515, 200 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
	}
	WinActivate, Inventory


	;MouseMove, 978, 200 ; Inventory button	
	MouseMove, 1515, 200 ; Inventory button
	Sleep, 500
	Click,  1535, 250 ; Search button
	sleep, 1500

	Loop,
	{
		MouseMove, 66, 278
		Sleep, 500
		click 2
		Sleep, 500
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
			MouseMove, 1515, 200 ; Inventory button
			Sleep, 500
			Click, 1535, 250 ; Search button
			Sleep, 500
		}
	}

	Sleep, 250
	Send, {Tab} ;Activates the claim type drop-down field
	Sleep, 250
	Send, {u} ;Goes up to Unique ID
	Sleep, 250
	Send, {Tab} ;Tab over to search field
	Sleep, 250
	Send, %UNIQUEID% ;Enters Unique ID from Excel
	Sleep, 250
	Send, {Enter}
	Sleep, 1000

; Wait for page to load
;Determine UNIQUEID Status. If In Process goto REQUEST. If anything else, go back to excel.
	Loop, 
	{
		Clipboard = 
		sleep, 200
		mousemove, 848, 430
		MouseClickDrag, Left, 848, 430, 915, 430
		;send, ^a
		sleep, 500
		send, ^c
		clipwait, 5,1
		
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
			;MsgBox, UNIQUEID Not Found
			MouseMove, 1515, 200 ; Inventory button
			Sleep, 500
			Click,  1535, 250 ; Search button
			sleep, 500
			comment = UNIQUEID Not Found
			goto, EXCEL
		}
		Else
		{
			;MsgBox, UNIQUEID Not In Process
			MouseMove, 1515, 200 ; Inventory button
			Sleep, 500
			Click,  1535, 250 ; Search button
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
	Sleep, 500
	Send, {end 2}
	Sleep, 500
	Send, ^a
	Sleep, 500
	Send, ^c
	ClipWait, 10, 1

	Loop,
	{
		If clipboard contains %NOTE%
		{
			Sleep, 500
			Send, {home 2}
			Sleep, 500
			MouseMove, 14, 278
			Sleep, 500
			Click
			MouseMove, 1515, 200 ; Inventory button
			Sleep, 500
			Click,  1535, 250 ; Search button
			sleep, 500
			Clipboard = 
			Comment = Request previously entered
			Goto, EXCEL
		}
		else
		{
			;Sleep, 500
			break
		}
	}
	Loop,
	{
		If clipboard contains New Request
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
			MouseMove, 1515, 200 ; Inventory button
			Sleep, 500
			Click,  1535, 250 ; Search button
			sleep, 500
			Clipboard = 
			Comment = New Request Not Found
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
			Sleep, 1000
			Send, ^f
			Sleep, 500
			Comment = New Request
			Send, %Comment%
			sleep, 500
			Send, {Tab 2}
			sleep, 250
			Send, +{Tab}
			Sleep, 250
			;Opens new request window
			Send, {Enter}
			sleep, 250
			;pauses script for Direct Connect refresh every 15 minutes
			;Loop, 5
			;{
				;if A_Min = 00
				;{
					;SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
					;sleep, 15000
					;SplashTextOff
				;}
				;if A_Min = 15
				;{
					;SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
					;sleep, 15000
					;SplashTextOff
				;}
				;if A_Min = 30
				;{
					;SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
					;sleep, 15000
					;SplashTextOff
				;}
				;if A_Min = 45
				;{
					;SplashTextOn, 500, 100, Direct Connect Inactive, Waiting on Direct Connect Refresh!
					;sleep, 15000
					;SplashTextOff
				;}
				;else
				;{
					;break
				;}
			;}
			Clipboard = 
			goto, POPUPWINDOWCHECK
		}
	}

POPUPWINDOWCHECK:

;Checks to make sure New Request window is open
	Sleep, 2500
	MouseMove, 815, 470
	Sleep, 500
	Click 2
	Sleep, 200
	Send, ^c
	clipwait, 2
	
	If clipboard = Category
	{
		Sleep, 500
		Goto NEWREQUEST
	}
	else
	{
		Sleep, 500
		MouseMove, 1155, 420 ;close window
		Sleep, 200
		Click
		Sleep, 1000
		Goto INCORRECTWINDOW
	}

NEWREQUEST:

;If Category was found in the pop-up window the script continues here
	Sleep, 250
	Send, {tab}
	Sleep, 250
	Send, {i}
	Sleep, 250,
	Send, {Tab}
	Sleep, 250
	Send, {down 3}
	Sleep, 250
	Send, {tab 3}
	Sleep, 250
	;Enters the Worklist Note
	Send, %NOTE%
	Sleep, 250
	;Tabs down to the Save button, checks to see if the Direct Connect refresh is going on, then clicks save
	Send, {Tab}
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
	;clicks Save button
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

;Look for Please Review Request
	If clipboard contains Account
	{
		;Goes to Please Review request
		Sleep, 1000
		Send, ^f
		Sleep, 500
		Comment = Please Review
		Send, %Comment%
		sleep, 500
		Send, {Tab 2}
		sleep, 250
		Send, +{Tab}
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
		goto, POPUPWINDOWCHECK2
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
		MouseMove, 1515,200 ; Inventory button
		Sleep, 1000
		Clipboard = 
		Comment = Double Check Claim
		GoTo, EXCEL
	}

POPUPWINDOWCHECK2:

;Checks to make sure New Request window is open
	Sleep, 5000
	MouseMove, 860, 408
	Sleep, 500
	Click 2
	Sleep, 200
	Send, ^c
	clipwait, 2
	
	If clipboard = Review
	{
		Sleep, 500
		Goto REVIEWED
	}
	else
	{
		Sleep, 500
		MouseMove, 1154,421
		Sleep, 200
		Click
		Sleep, 1000
		Goto INCORRECTWINDOW
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
			MouseMove, 1515,200 ; Inventory button
			Sleep, 500
			Clipboard = 
			Comment = Double Check Claim
			GoTo, EXCEL
		}


RETRACTIONNOTFOUND:

;Unable to find Retraction Request for Unique ID
	Sleep, 1000
	;MouseMove, 915, 358
	MouseMove, 915, 340
	Sleep, 200
	Click
	Sleep, 500
	send, {Home 2}
	sleep, 500
	Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
	Sleep, 200
	MouseMove, 20,125 ; Close Find Window
	Sleep, 500
	Click, 20,125 ; Close Find Window
	Sleep, 500
	MouseMove, 1515,200 ; Inventory button
	Sleep, 1000
	Clipboard = 
	Comment = Retraction Request Not Found
	GoTo, EXCEL

INCORRECTWINDOW:

;If Category was NOT found in the pop-up window the script continues here 
	Clipboard = 
	Sleep, 1000
	Send, {home 2}
	Sleep, 500
	Send, ^f
	Sleep, 500
	Comment = New Request
	Send, %Comment%
	sleep, 500
	Send, {Tab 2}
	sleep, 250
	Send, +{Tab}
	Sleep, 250
	;Opens new request window
	Send, {Enter}
	sleep, 250
	Send, {Enter}
	sleep, 1000
	MouseMove 815, 470
	Sleep, 500
	Click 2
	Sleep, 250
	Send, ^c
	clipwait, 1
	Loop, ;determine if second popup window can be used
	{
		If clipboard contains Category
		{
			Sleep, 500
			Goto, NEWREQUEST
		}
		else
		{
			Sleep, 500
			MouseMove, 1154, 421			
			;MouseMove, 833, 389
			Sleep, 200
			Click
			Sleep, 1000
			Clipboard = 
			Comment = New Request Not Found
			Goto Excel
		}
	}

EXCEL:

;Update Excel with comment based on actions above
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	Sleep, 250
	Send, {Tab 2}
	Sleep, 250
	Send, %Comment%
	Sleep, 250
	Send, {Down}
	Sleep, 250
	Send, {home}
	Sleep, 250
}

FormatTime, TimeEnd,, Time
MsgBox, %Count% UNIQUEID's resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
Esc::Pause