#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column F next to the account you want to start with. Now log into Direct Connect.  Column A should contain the UID #, column B Resolved Reason, column c CBO Organization, column D Payor Organization, column E should be blank, .
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

WinActivate, Dire
WinWaitActive, Dire
;WinActivate, Inventory
;WinWaitActive, Inventory
MouseMove, 865, 207 ; Inventory button
Sleep, 500
Click, 868,251 ; Search button
Sleep, 500


sleep, 1000
Loop, {
mouseclickdrag, left, 35, 277, 212, 277
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}
					
; Wait for page to load

Sleep, 1500
Send, {Tab}

;Sleep,500
Send, {Down} ; Go down to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 500
Send, %Account%
Sleep, 500
Send, {Enter}

; Wait for page to load
Loop, 
	{
	PixelSearch, FX, YX, 33,425, 140,450, 0xDDCCFF, 5, Fast
	If ErrorLevel = 1
		{
		;MsgBox, Account # results loaded.
		Sleep, 1000
		Break
		}
	}
;this clicks hyper link for acct
MouseMove, 104,426
Sleep, 100
Click, 104,426
Click, 104,426
Click, 104,426

; Wait for page to load
Clipboard =  ; clears the clipboard
Sleep, 2000


Loop, {
sleep, 2000
mouseclickdrag, left, 45, 278, 189, 278
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Image found.
Break
}
}

sleep, 3000

Sleep, 2000
Clipboard = 

Loop, {
mouseclickdrag, left, 458, 427, 506, 427
Send, ^c
clipwait, 
If clipboard contains Complete
	
{
;MsgBox, Image found.
comment = Account Complete
goto, EXCEL
}
else
	break

}


sleep, 1000
Loop, {
mouseclickdrag, left, 45, 278, 189, 278
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
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

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000




; Wait for New Request Window screen to disappear
Sleep, 3000
Loop, {
sleep, 500
mouseclickdrag, left, 45, 278, 189, 278
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
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 1500



EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, Account Reopened
Sleep, 500
Send, {Down}
Sleep, 500

}

FormatTime, TimeEnd,, Time
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause