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

IfWinNotExist,  Windows Internet Explorer
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


;MsgBox, %MyArray2%
;MsgBox, %MyArray3%

WinActivate, Windows Internet Explorer
WinWaitActive, Windows Internet Explorer
MouseMove, 969, 216 ; Inventory button
Sleep, 1000
Click, 974, 262 ; Search button
Sleep, 1000
;Mousemove, 10, 10

; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 30,269, 157, 269, %Imagelocation%DC Search Inventory Image.bmp
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
;mouseclickdrag, left, 30, 269, 157, 269
;Send, ^c
;lipwait, 
;If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}

Click, 103, 331 ; Click to select 
Sleep, 1000
Click, 103, 331
Sleep, 1000
Send, {Up} ; Go down to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 1000
Send, %Account%
Sleep, 1000
Send, {Enter}

;this clicks hyper link for acct
MouseMove, 68, 443
Sleep, 750
Click, 68, 443
Click, 68, 443


; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 41,270, 145, 270, %Imagelocation%DC Account View Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Account information page loaded.
		;Sleep, 1000
		;Break
		;}
	;}
Sleep, 2000

Loop, {
mouseclickdrag, left, 437, 442, 505, 442
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
mouseclickdrag, left, 47,292, 189, 292
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
Sleep, 1000
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


Send, %MyArray2%
Sleep, 1000
Send, {Tab}
Sleep, 1000

Send, {Home}
Sleep, 1000

sleep, 1000
MouseMove, 50, 240
Sleep, 100
Click, 50, 240
Loop, {
mouseclickdrag, left, 47,292, 189, 292
Send, ^c
clipwait, 
sleep, 750
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
MouseMove, 50, 240
Sleep, 100
Click, 50, 240
Loop, {
mouseclickdrag, left, 47,292, 189, 292
Send, ^c
clipwait, 
sleep, 750
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
MouseMove, 50, 240
Sleep, 100
Click, 50, 240
Loop, {
mouseclickdrag, left, 36, 291, 216, 291
Send, ^c
clipwait, 
sleep, 750
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