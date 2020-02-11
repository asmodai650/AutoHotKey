#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column D next to the account you want to start with. Column A should contain UID, column B should contain new worklist and column C should contacin Discovery Status . Now log into Direct Connect under UHC org.   This uses 11 tabs after line 156 to reassign to new list.
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
	MsgBox, Ensure Excel is open.
	GoTo, CHECK	
	}

IfWinNotExist, Internet Explorer
	{
	MsgBox, Ensure the home Direct Connect page is open.
	GoTo, CHECK	
	}

FormatTime, TimeBegin,, Time

Loop, %Count%
{	
TrayTip,,%a_index% of %Count%,30	

Account =
Reason =

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
New_Worklist = %MyArray2% 

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

WinActivate, Internet Explorer
WinWaitActive, Internet Explorer
MouseMove, 970, 205 ; Inventory button
Sleep, 1000
Click, 985, 245 ; Search button
Sleep, 1000

; Wait for page to load

sleep, 1000

Click, 118, 315 ; Click to select claim number drop down
Sleep, 1000
Click, 118, 315
Sleep, 1000
Send, {Up} ; Go up to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 1000
Send, %Account%
Sleep, 1000
Send, {Enter}
 

MouseMove, 64, 436 ;Click hyper link for acct
Sleep, 750
Click, 64, 436
Click, 64, 436
; Wait for page to load
Sleep, 2000

Loop, {
mouseclickdrag, left, 453, 425, 508, 425
Send, ^c
clipwait, 
sleep, 750
If clipboard contains Complete
	
{
;MsgBox, Image found.
comment = Account Complete
goto, EXCEL
}
else
	break

}

Loop, {
sleep, 1000
mouseclickdrag, left, 45, 278, 190, 278
Send, ^c
clipwait, 
sleep, 750
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Send, {Tab} ; tabs down to worklist box
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}
Sleep, 750
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

Send, %MyArray2% ;enters new worklist
Sleep, 1000
Send, {Tab}
Sleep, 1000

Send, {Home} ;sends web browser to top
Sleep, 1000
Send, {Home} ;send web browser to top again to make sure it didn't stop part way
Sleep, 1000

Sleep, 500
MouseMove, 35, 225
Click, 35, 225
Loop, {
mouseclickdrag, left, 45, 278, 190, 278
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Sleep, 500 ;goes to save box
Send, {Tab}
Sleep, 500
Send, {Enter}

sleep, 1000
MouseMove, 35, 225
Click, 35, 225
Loop, {
mouseclickdrag, left, 45, 278, 190, 278
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Sleep, 500 ;goes to cancel box
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}

sleep, 1000
MouseMove, 35, 225
Click, 35, 225
Loop, {
mouseclickdrag, left, 34, 278, 213, 278
Send, ^c
clipwait, 
Sleep, 500
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}


Loop,
	{
	ImageSearch, FX, YX, 398,370, 553,450, %Imagelocation%BI New Request Window.bmp
	If ErrorLevel = 1
		{
		;MsgBox, New Request Window closed.
		Comment = Reassigned
		GoTo, EXCEL
		}
	}



EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, %Comment%
Sleep, 500
Send, {Down}
Sleep, 500

}

FormatTime, TimeEnd,, Time
SoundPlay *48
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause