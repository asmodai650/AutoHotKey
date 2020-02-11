;Optum Direct Connect
;Select United Healthcare (Payer)>Inventory>Search



#SingleInstance Force
;WinWait Search Inventory

IfWinNotExist Microsoft Excel
  {
   Msgbox The spreadsheet must be named "OPD Appeal and Inquiry Resolution Report" in order for the macro to work. Thanks.
   ExitApp
  }

Beginning:
Loop
{
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

WinActivate Microsoft Excel
WinWaitActive Microsoft Excel
Sleep, 750
Send {HOME}{SHIFTDOWN}{RIGHT}{SHIFTUP}
ClipBoard =
ResolutionDescription =
Send ^c
ClipWait

StringSplit, cell, clipboard, %A_Tab%,

ClaimID = %cell1%
ResolutionDescription = %cell2%
StringReplace, ResolutionDescription, ResolutionDescription, `r`n, , All
RegExReplace(ResolutionDescription, "[`r`n`t]+$")
ClipBoard =

If ResolutionDescription =
{
Msgbox The macro is complete.
ExitApp
}

SearchInventory:
IfWinNotExist Search Inventory
  {

IfWinExist My Inventory
{
WinActivate My Inventory
WinWaitActive My Inventory
}

IfWinExist Partner Account View
{
WinActivate Partner Account View
WinWaitActive Partner Account View
}


IfWinExist Frontier
{
WinActivate Frontier
WinWaitActive Frontier
Sleep, 750
Send {BACKSPACE}
Sleep, 7500
}

IfWinExist Error
{
WinActivate Error
WinWaitActive Error
}

Sleep, 750
If A_UserName contains abaugh1,cmoor61
MouseMove 1234, 200
Else
MouseMove 980, 200
Sleep, 750
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
Click 985, 250
Sleep 2000
  }

WinWait, Search Inventory, ,3
IfWinNotExist Search Inventory
  GoTo SearchInventory

SearchInventory2:
WinWait, Search Inventory, ,3

IfWinNotExist Search Inventory
{
Send {BACKSPACE}
GoTo SearchInventory
}

WinActivate Search Inventory
WinWaitActive Search Inventory
Sleep, 750

;Send ^f
;Sleep, 750
;Send claim number{ESC}
;Sleep, 750

click 132, 315
Send u
sleep, 200
Send, {enter}
Sleep, 200
	Send, {Tab} ;Tab over to search field
	Sleep, 500
	Send, %ClaimID% ;Enters RES ID from Excel
	Sleep, 500
	Send, {Enter}
	Sleep, 1000



;send, {Tab 2} %ClaimID% {ENTER}
;Sleep 2000

Click 645, 175
Sleep 1000

Send ^a
Sleep, 750

Send ^c
ClipWait, 2

If ClipBoard contains No items to display
{
Sleep, 750
If A_UserName contains abaugh1,cmoor61
MouseMove 1234, 200
Else
MouseMove 980, 200
Sleep, 750
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
Click 985, 250
Sleep 2000
Sleep 4000
  GoTo SearchInventory2
}

If ClipBoard not contains In Process
  If ClipBoard contains %ClaimID%
   {
   Send {HOME}
   Sleep, 750
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   MouseMove 980, 200
   Sleep, 750
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   Click 985, 250
   Sleep 1000
   WinActivate Microsoft Excel
   WinWaitActive Microsoft Excel
   Sleep, 750
   Send {RIGHT 4}Not {"}In Process{"}{DOWN}
   ClipBoard =
   GoTo Beginning
   }
  Else
   GoTo SearchInventory

ClipBoard =
click 65, 438, 2
Sleep, 750

;Send ^f
;Sleep, 750
;Send in process{ESC}
;Sleep 1000
;msgbox %A_CaretX%, %A_CaretY%
;ControlGetPos , X, Y, Width, Height, ListView20WndClass1, Claim ID

;Send +{TAB}

;Sleep, 750
;Send {ENTER}

WinWaitClose Search Inventory
WinWait, Partner Account View, ,3

IfWinNotExist Partner Account View
{
Send {BACKSPACE}
Sleep 2000
GoTo SearchInventory2
}


;Checking for Window
OpenRequests:
{
; added line below on 20170504
Sleep, 2000
ClipBoard =

Send ^a
Sleep, 750

Send ^c
ClipWait, 2
}

;ADDED 204-227 TO KEEP MACRO FROM STALLING IF THERE IS NO OPEN REQUEST FOR CLAIM - DCURTIS 20160617
;If ClipBoard not contains Open Requests
  ;GoTo OpenRequests
If ClipBoard not contains Open Requests
  If ClipBoard contains %ClaimID%
   {
   Send {HOME}
   Sleep, 750
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   MouseMove 980, 200
   Sleep, 750
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   Click 985, 250
   Sleep 1000
   WinActivate Microsoft Excel
   WinWaitActive Microsoft Excel
   Sleep, 750
   Send {RIGHT 4}No {"}Open Requests{"}{DOWN}
   ClipBoard =
   GoTo Beginning
   }
  Else
   GoTo SearchInventory

ClipBoard =
click 65, 438, 2
Sleep, 750

Send ^f
Sleep, 750
Send open requests{ESC}
Sleep, 750

Send {TAB}{ENTER}
Sleep, 750

;Checking for Window
RequestResponse:
ClipBoard =

Send ^a
Sleep, 750

Send ^c
ClipWait
RequestResponseData :=ClipBoard
ClipBoard =

If RequestResponseData not contains Response Comments
  GoTo RequestResponse

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

ClipBoard =
ClipBoard := ResolutionDescription
Send ^f
Sleep, 750
Send request comments{ESC}
Sleep, 750
Send {TAB 2}n{tab}^v{TAB}{ENTER}
;new section
;Send {TAB 2}n{tab}
;Sleep, 750
;sendraw %clipboard%
;Sleep, 750
;Send {TAB}{ENTER}
Sleep, 750
;end new section
ClipBoard =

If RequestResponseData contains Response Date
{
Sleep, 7500
Send ^f
Sleep, 750
Send dialog{TAB}{ENTER}{ESC}
Sleep, 750
Send {TAB}{ENTER}
Sleep 1000
Send ^f
Sleep, 750
Send category{ESC}
Sleep, 750
Send {TAB 2}{DOWN 6}{TAB 3}
Sleep, 750

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

ClipBoard =
ClipBoard := ResolutionDescription

Send ^v{TAB}{ENTER}
Sleep, 750
ClipBoard =
}
Sleep 3000

;Checking for Window
AccountView:
ClipBoard =

Send ^a
Sleep, 750

Send ^c
ClipWait

If ClipBoard not contains Account View
  GoTo AccountView

ResolveAccount:
Send ^f
Sleep, 750
Send edit
Sleep, 750
Send {ESC}
Sleep, 750
Send +{TAB}
Sleep, 750

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

Send {ENTER}
Sleep, 750

;Checking for Window

ClipBoard =

Send ^a
Sleep, 750

Send ^c
ClipWait, 2

If ClipBoard not contains Resolve Reason
  GoTo ResolveAccount

Sleep, 750

Send ^f
Sleep, 750
Send resolve reason{ESC}
Sleep, 750
Send {TAB}c
Sleep, 750

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

Send {TAB}{ENTER}
Sleep 7000

WinActivate Microsoft Excel
WinWaitActive Microsoft Excel
Sleep, 750
Send {RIGHT 4}x{DOWN}

}

pause::pause