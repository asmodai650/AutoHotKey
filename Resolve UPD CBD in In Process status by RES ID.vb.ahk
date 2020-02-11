#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column F next to the account you want to start with. Now log into Direct Connect.  Column A should contain the Account Resolution ID, column B the "Resolved Reason", column c CBO Organization, column D Payor Organization, column E should be blank, .
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
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

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
MouseMove, 865, 207 ; Inventory button
Sleep, 1000
mousemove, 868,251
sleep, 500
Click, 868,251 ; Search button
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

Sleep, 3000

Loop, {
MouseMove, 50, 240
Sleep, 100
mouseclickdrag, left, 45, 278, 190, 278
Send, ^c
clipwait, 
If clipboard contains Account View

{
Break
}
}


Sleep, 3000
Clipboard =
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

sleep, 5000
Loop, {
MouseMove, 50, 240
Sleep, 100
mouseclickdrag, left, 45, 278, 190, 278
Send, ^c
clipwait, 
If clipboard contains Account View

{
Break
}
}

Sleep, 5000
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 500

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

;Wait for New Request Window screen to pop up
MouseMove, 628, 516
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
Sleep, 1500


; Wait for New Request Window screen to disappear
Sleep, 2000
Loop, {
sleep, 500
MouseMove, 50, 240
Sleep, 100
mouseclickdrag, left, 45, 278, 190, 278
Send, ^c
clipwait, 
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
sleep, 500

Send, {Home}
Sleep, 500
Send, {Home}
Sleep, 500

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
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause