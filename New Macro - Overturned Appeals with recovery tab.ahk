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
If A_UserName contains abaugh1,cmoor61,derrick,dreece150
MouseMove 1234, 200
Else
MouseMove 860, 200
Sleep, 750
If A_UserName contains abaugh1,cmoor61,derrick,dreece150
Click 1250, 250
Else
Click 890, 250
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
Send, {Enter}
Sleep, 1000
	Send, {Tab} ;Tab over to search field
	Sleep, 500
	Send, %ClaimID% ;Enters UID from Excel
	Sleep, 500
	Send, {Enter}
	Sleep, 1000


MouseMove, 62, 432 ;this clicks hyper link for acct
Sleep, 3000
Click, 62, 432
Click, 62, 432


Sleep, 5000

MouseMove, 434, 427
Click, 474, 429
Click, 474, 429
Click, 474, 429
Sleep, 800
Send, ^c
Sleep, 3000

If clipboard contains In Process
{	
Clipboard = 
Sleep, 2000
GoTo, NEWREQUEST
}	

If clipboard contains Resolved
{
Clipboard = 
Sleep, 2000
GoTo, REOPEN
}

If clipboard contains Complete
{
Clipboard = 
Sleep, 2000
Comment = Account Complete
GoTo, Excel
}

REOPEN:
Sleep, 2000
Send, ^f
Sleep, 750
Send, edit
Sleep, 750
Send, {ESC}
Sleep, 750
Send, +{TAB}
Sleep, 750
Send, {Enter}
Sleep, 3000

GoTo NEWREQUEST



NEWREQUEST:

Sleep, 2000
Send, ^f
Sleep, 750
Send, New Request
Sleep, 750
Send, {ESC}
Sleep, 750
Send, {Enter}
Sleep, 3000

Send, {TAB}
Sleep, 750
Send, Information
Sleep, 750
Send, {TAB}
Sleep, 750
Send, Please Review
Sleep, 750
Send,{TAB}
Sleep, 750
Send, {TAB}
Sleep, 750
Send, {TAB}
Sleep, 750
Send, %cell2%
Sleep, 750
Send, {TAB}
Sleep, 750
Send, {ENTER}
Sleep, 3000
GoTo, ResolveAccount


ResolveAccount:
Sleep, 3000
Send, ^f
Sleep, 750
Send, edit
Sleep, 750
Send, {ESC}
Sleep, 750
Send, +{TAB}
Sleep, 750
Send, {ENTER}
Sleep, 750
Send, {TAB}
Sleep, 750
Send, Created in Error
Sleep, 750
Send, {TAB}
Sleep, 750	
Send, {ENTER}
Sleep, 5000

Comment = Macro Resolved


Excel:
WinActivate Microsoft Excel
WinWaitActive Microsoft Excel
Sleep, 750
Send, {Right}
Send, {Right}
Send, {Right}
Send, {Right}
Send, %comment%
Send, {DOWN}
Sleep, 3000

}

pause::pause