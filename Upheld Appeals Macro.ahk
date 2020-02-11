#SingleInstance, Force
SetTitleMatchMode, 2


MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect and select UnitedHealthcare (Payer) from the provider list. The Excel Columns should be as follows: Column A - Unique ID and Column B - Appeal Response.

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

	UniqueID =
	AppealResponseA =

	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel

	Sleep, 500
	Send, +{Home}
	Sleep, 500
	Send, ^c
	ClipWait, 1
	StringSplit, MyArray, clipboard, %A_Tab%
	UniqueID = %MyArray1% 
	AppealResponseA = %MyArray2%

	If UniqueID =
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
				send, {PgUp}
				sleep, 200
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 200
				MouseMove, 22,126 ; Close Find Window
				Sleep, 500
				Click, 20,125 ; Close Find Window
				Sleep, 500
				MouseMove, 965,203 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
			else
			{
				sleep, 2000
				send, {PgUp}
				sleep, 200
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 200
				MouseMove, 22,126 ; Close Find Window
				Sleep, 500
				Click, 20,125 ; Close Find Window
				Sleep, 500
				MouseMove, 965,203 ; Inventory button
				Sleep, 5000
				Click
				Sleep, 1000
			}
	}
	WinActivate, Inventory
	Sleep, 200
	Send, {PgUp}

	MouseMove, 965, 203 ; Inventory button
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
		
		If clipboard contains Search
		{
			Break
		}
		Else
		{
			Sleep, 200
			Send, ^f
			Sleep, 200
			MouseMove, 22,126
			Sleep, 200
			Click, 20,125 ; Close Find Window
			sleep, 200
			send, {PgUp} ; Send page back to top
			sleep, 500
			MouseMove, 965,203 ; Inventory button
			Sleep, 500
			Click, 978, 240 ; Search button
			Sleep, 500
		}
	}

	Sleep, 500
	Send, {Tab} ;Activates the claim type drop-down field
	Sleep, 500
	Send, {u} ;Goes up to Resolution ID
	Sleep, 500
	Send, {Tab} ;Tab over to search field
	Sleep, 500
	Send, %UniqueID% ;Enters Unique ID from Excel
	Sleep, 500
	Send, {Enter}
	Sleep, 1000

; Wait for page to load
;Determine UniqueID Status. If anything else, go back to excel.
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
			;MsgBox, UniqueID Not Found
			MouseMove, 965, 203 ; Inventory button
			Sleep, 500
			Click,  978, 240 ; Search button
			sleep, 500
			comment = UniqueID Not Found
			goto, EXCEL
		}
		
		Else
		{
			;MsgBox, UniqueID Not Resolved
			MouseMove, 965, 203 ; Inventory button
			Sleep, 500
			Click,  978, 240 ; Search button
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
			Send, {PgUp}
			MouseMove, 14, 278
			Sleep, 500
			Click
			Sleep, 1000
			break
		}
		else
		;Look for "Additional Information Supplied" and if found "goto NEWREQUEST:" to enter new request of Financial
		{
			Sleep, 500
			Send, {PgUp}
			Sleep, 1000
			Send, ^a
			Sleep, 500
			Send, ^c
			ClipWait, 10, 1
			Sleep, 1000
				
				If clipboard contains Additional Information Supplied
				{
					;msgbox, TEXT FOUND!
					Sleep, 500
					clipboard = 
					Goto, NEWREQUEST
				}
				else
				{	
					;MsgBox, Open Request Not Found					
					Sleep, 500
					Send, {PgUp}
					Sleep, 1000
					MouseMove, 14, 278
					Sleep, 500
					Click
					MouseMove, 965, 203 ; Inventory button
					Sleep, 500
					Click,  978, 240 ; Search button
					sleep, 500
					Clipboard = 
					Comment = Open Request Not Found
					Goto, EXCEL
				}
		}
	}
	Loop,
	{
		send, {PgUp}
		Sleep, 500
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
			send, {PgUp}
			Sleep, 500
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
			
			Clipboard = 
			goto, POPUPWINDOWCHECK
		}
	}

POPUPWINDOWCHECK:

;Opens new request window
	Sleep, 5000
	MouseMove, 535, 380
	Sleep, 500
	Click 2
	Sleep, 200
	Send, ^c
	clipwait, 2
	
	If clipboard Contains Supply
	{
		Sleep, 1000
		Goto SUPPLY
	}

;**********ADD REVIEW SEARCH CODE HERE**********
	If clipboard Contains Review
	{
		Sleep, 1000
		Goto REVIEW
	}
;**********END REVIEW SEARCH CODE HERE**********

	else
	{
		Sleep, 500
		MouseMove, 914, 337
		Sleep, 200
		Click
		Sleep, 1000
		Goto POPUPWINDOWCHECK2
	}

;**********ADD REVIEW CODE HERE**********
REVIEW:

;If REVIEW was found in the pop-up window the script continues here
	Sleep, 250
	Send, {tab 2}
	Sleep, 250
	Send, {r}
	Sleep, 250
	Send, {Tab 2}
	sleep, 250
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
			break ;If not paused for one of the minutes above, continue
		}
	}

	Send, {Enter}
	Sleep, 2000
	Clipboard = 
	GoTo, NEWREQUEST
;**********END REVIEW CODE HERE**********	

SUPPLY:

;If SUPPLY was found in the pop-up window the script continues here
	Sleep, 250
	Send, {tab 2}
	Sleep, 250
	Send, {a}
	Sleep, 250
	Send, {Tab 2}
	sleep, 250
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
			break ;If not paused for one of the minutes above, continue
		}
	}

	Send, {Enter}
	Sleep, 2000
	Clipboard = 
	GoTo, NEWREQUEST
	
NEWREQUEST:
; ENTER NEW REQUEST CODE HERE!!!!!
	sleep, 2000
	send, {PgUp}
	sleep, 1000
	Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page

;Goes to new request button
	Sleep, 500
	Comment = New Request
	Send, %Comment%
	sleep, 500
	Send, {Tab 2}
	sleep, 250
	Send, +{Tab}
	sleep, 250
	Send, {Enter}
	sleep, 250
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
	
	Clipboard = 
	sleep, 2000
	MouseMove 491, 435
	Sleep, 500
	Click 2
	Sleep, 250
	Send, ^c
	clipwait, 2
	
	Loop, ;determine if second popup window can be used
	{
		If clipboard contains Category
			{
				;MsgBox, Category Found
				Sleep, 200
				Clipboard = 
				Sleep, 500
				Goto, NEWREQUEST2
			}
			else
			{
				Sleep, 500
				MouseMove, 836,387
				Click 1
				Sleep, 3000
				send, {PgUp}
				sleep, 1000
				Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
				Sleep, 200
				MouseMove, 22,126 ; Close Find Window
				Sleep, 500
				Click, 20,125 ; Close Find Window
				Sleep, 500
				MouseMove, 965,203 ; Inventory button
				Sleep, 1000
				Clipboard = 
				Comment = New Request Not Found
				GoTo, EXCEL
			}
	}

NEWREQUEST2:
;If Category was found in the pop-up window the script continues here
	Sleep, 250
	Send, {tab}
	Sleep, 250
	Send, {a}
	Sleep, 250
	
;check for approval or financial
	MouseMove, 491, 439
	Sleep, 200
	Send, ^a
	Sleep, 200
	Send, ^c
	sleep, 250
	ClipWait, 2, 1
	sleep, 500

	Loop, ;determine if second popup window can be used. If Select a Category is found, then Approval was not an option and Financial is the correct choice
	{
		If clipboard contains Category
			{
				;Msgbox, SELECT FOUND
				Sleep, 500
				clipboard = 
				Sleep, 200
				click, 491, 439
				sleep, 250
				Send, {tab}
				Sleep, 200
				Send, {f}
				;msgbox, financial entered
				Sleep, 750
				Send, {tab}
				Sleep, 750
				Send, {p}
				Sleep, 200
				Send, {tab 3}
				Sleep, 250
				SendRaw, %AppealResponseA%
				Sleep, 250
				Send, {Tab 2}
				sleep, 250
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
							break ;If not paused for one of the minutes above, continue
						}
					}
				Send, {Enter}
				Sleep, 2000
				Clipboard = 
				GoTo, CLOSECLAIM
			}
			else ;Financial not found, so Approval is the correct option
			{
				;MsgBox, Approval Found
				Sleep, 200
				click, 491, 439
				Sleep, 200
				break
			}
	}
	
	Send, {tab 5}
	Sleep, 250
	SendRaw, %AppealResponseA%
	Sleep, 250
	Send, {Tab 2}
	sleep, 250
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
			break ;If not paused for one of the minutes above, continue
		}
	}

	Send, {Enter}
	Sleep, 2000

	Clipboard = 
	GoTo, CLOSECLAIM

POPUPWINDOWCHECK2:

;If SUPPLY was NOT found in the pop-up window the script continues here 
	Clipboard = 
	Sleep, 1000
	Send, {PgUp}
	Sleep, 500
	Send, {PgUp}
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
	MouseMove 536 , 378
	Sleep, 500
	Click 2
	Sleep, 250
	Send, ^c
	clipwait, 2
	Loop, ;determine if second popup window can be used
	{
		;If clipboard NOT contains Supply
		;{
			;Sleep, 500
			;Goto, SUPPLYNOTFOUND
		;}
		If clipboard contains Supply
		;{
			;MsgBox, Supply Found
			Sleep, 200
			Clipboard = 
			Send, ^a
			Sleep, 200
			Send, ^c
			ClipWait, 2
			If Clipboard Contains Responded by
			{
				Sleep, 500
				Goto, SUPPLYNOTFOUND
			}
			else
			{
				Sleep, 500
;;;; MAY NEED TO CHANGE CO-ORDINATES
				MouseMove, 635,366
				Sleep, 200
				Click
				Sleep, 200
				MouseMove, 536, 378
				Sleep, 500
				Click 2
				Sleep, 200
				Send, ^c
				clipwait, 2
				If clipboard = Supply
					{
						Sleep, 500
						Goto SUPPLY
					}
				else
					{
						Sleep, 500
						Goto, SUPPLYNOTFOUND
					}
			}
		;}
	}

SUPPLYNOTFOUND:

;Unable to find SUPPLY Request for Res ID
	Sleep, 1000
	MouseMove, 913, 338
	Sleep, 200
	Click
	Sleep, 1000
	send, {PgUp}
	sleep, 500
	Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
	Sleep, 200
	MouseMove, 22,126 ; Close Find Window
	Sleep, 500
	Click, 20,125 ; Close Find Window
	Sleep, 500
	MouseMove, 965,203 ; Inventory button
	Sleep, 1000
	Clipboard = 
	Comment = Please Supply Additional Information Not Found
	GoTo, EXCEL

CLOSECLAIM:

;Determine if Direct Connect loaded correctly
	Sleep, 500
	Send, {PgUp}
	Sleep, 500
	MouseMove, 72, 306
	Sleep, 500
	click 2
	Sleep, 250
	Send, ^c
	clipwait, 1

	If clipboard contains Search
	{
		MouseMove, 22,126 ; Close Find Window
		Sleep, 500
		Click, 20,125 ; Close Find Window
		Sleep, 1000
		Clipboard = 
		Comment = Please Approve Overpayment Request Entered
		Goto, EXCEL
	}
	Else
	{
		sleep, 2000
		send, {PgUp}
		sleep, 200
		Send, ^f ;opens find window just in case it isnt open to keep from clicking off this page
		Sleep, 200
		MouseMove, 22,126 ; Close Find Window
		Sleep, 500
		Click, 20,125 ; Close Find Window
		Sleep, 500
		MouseMove, 965,203 ; Inventory button
		Sleep, 1000
		Clipboard = 
		Comment = Double Check Claim
		GoTo, EXCEL
	}


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
MsgBox, %Count% UniqueID's resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
Esc::Pause