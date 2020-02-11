#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\mreece2\Desktop\Macro Images

MsgBox, In your Excel workbook, select the cell in column F next to the account you want to start with. Now log into Direct Connect.  Column A should contain the Account Resolution ID, column B the "Resolved Reason", column c CBO Organization, column D Payor Organization, column E should be the account status, .
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
	MsgBox, Ensure Excel is open.
	GoTo, CHECK	
	}

IfWinNotExist,  Internet Explorer
	{
	MsgBox, Ensure the home Direct Connect page is open.
	GoTo, CHECK	
	}

FormatTime, TimeBegin,, Time

	

Loop, %Count%

{
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

TrayTip,,%a_index% of %Count%,30	

Account =
Reason = 
Category =
Subject =

Clipboard = 


WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
Account = %MyArray1% 
Category = %MyArray2%


WinActivate, Internet Explorer
WinWaitActive, Internet Explorer
;MouseMove, 865, 207 ; Inventory button
MouseMove, 1515, 200 ; Inventory button
Sleep, 1000
;mousemove, 868,251
mousemove, 1535, 250
sleep, 500
;Click, 868,251 ; Search button
Click, 1535, 250 ; Search button
Sleep, 500

; Wait for page to load

mousemove, 70, 240
click, 70, 240
sleep, 1000

Click, 118, 315 ; Click to select type of claim number
Sleep, 500
Click, 118, 315
Sleep, 500
Send, {up} ; Go up to Resolution ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 500
Send, %Account%
Sleep, 500
Send, {Enter}
Sleep, 1500
; Wait for page to load

MouseMove, 62, 432 ;this clicks hyper link for acct
Sleep, 5000
Click, 62, 432
Click, 62, 432

; Wait for page to load
Sleep, 5000

Loop,
	{
		mouseclickdrag, left, 453, 425, 508, 425
		Send, ^c
		clipwait, 
		If clipboard contains Complete
			{
				comment = Account Complete
				goto, EXCEL
			}
		else
			break
	}

Loop,
	{
		sleep, 5000
		MouseMove, 50, 240
		Sleep, 100
		mouseclickdrag, left, 45, 278, 190, 278
		Send, ^c
		clipwait, 
		If clipboard contains An error has occurred
			{
				SoundPlay *48
				MsgBox, Error has Occurred
				ExitApp
			}
		If clipboard contains Account View
			{
				Break
			}
	}

Sleep, 5000
Send, {TAB}
Sleep, 5000
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 500

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

;Wait for New Request Window screen to pop up
;MouseMove, 628, 516
MouseMove, 760, 419
Sleep, 5000
;Click, 628, 516 ; click to select category drop down menu
Send, {Tab}
Sleep, 500
Send, %MyArray2%
Sleep, 500
;Send, {Enter}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}
Sleep, 2000


; Wait for New Request Window screen to disappear
Sleep, 5000
Loop,
	{
		sleep, 500
		MouseMove, 50, 240
		Sleep, 100
		mouseclickdrag, left, 45, 278, 190, 278
		Send, ^c
		clipwait, 
		If clipboard contains An error has occurred
			{
				SoundPlay *48
				MsgBox, Error has Occurred
				ExitApp
			}
		If clipboard contains Account View
			{
				Break
			}
	}

Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 5000

EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, Account Resolved
Sleep, 500
Send, {Down}
Sleep, 500
}

FormatTime, TimeEnd,, Time
SoundPlay *48
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause