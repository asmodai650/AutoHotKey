#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column G "Status", across from the account you want to start with. Column A should contain ODAR UID, column B MACRO worklist, column C Worklist, column D Dialog Age, column E Discovery Status, column F blank,. Now log into Direct Connect under UnitedHealthcare (Payer). This uses 11 tabs after line 156 to reassign to the new Inquiries worklist.
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
MouseMove, 970, 204 ; Inventory button
Sleep, 1000
Click, 985, 245 ; Search button
Sleep, 1000
;Mousemove, 10, 10

; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 28,256, 156, 256, %Imagelocation%IE11 DC Search Inventory Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Search Inventory page loaded.
		;Break
		;}
;if Errorlevel = 1
;Msgbox, Image not found
	;}

sleep, 1000
Loop, {
;mouseclickdrag, left, 28, 256, 156, 256
;Send, ^c
;clipwait, 
;If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}

Click, 118, 315 ; Click to select 
Sleep, 1000
Click, 118, 315
Sleep, 1000
Send, {Up} ; Go down to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 1000
Send, %Account%
Sleep, 1000
Send, {Enter}

;this clicks hyper link for acct
MouseMove, 62, 427
Sleep, 750
Click, 62, 427
Click, 62, 427


; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 50, 254, 153, 254, %Imagelocation%IE11 DC Account View Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Account information page loaded.
		;Sleep, 500
		;Break
		;}
	;}
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

Send, {Tab}
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
Send, {Tab}
Sleep, 1000

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

Send, %MyArray2%
Sleep, 1000
Send, {Tab}
Sleep, 1000

Send, {Home}
Sleep, 500
Send, {Home}
Sleep, 500

sleep, 500
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

Sleep, 500
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

Sleep, 500
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

;send, {Alt}
;Sleep, 500
;send, 1
;Sleep, 10000
}

FormatTime, TimeEnd,, Time
SoundPlay *48
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause