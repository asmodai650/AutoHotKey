#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO

MsgBox, In your Excel workbook, select the first empty cell in first empty column next to the account you want to start with. Now log into Direct Connect.  Column A should contain the RES ID#, column B anything, and column C should contain new request comments.
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
	;MsgBox, Ensure Excel is open.
	GoTo, CHECK	
	}

IfWinNotExist,  Inventory
	{
	;MsgBox, Ensure the home Direct Connect page is open.
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
Subject =
Note =
Comment =

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
Account = %MyArray1% 
Subject = %MyArray2% 
Note = %MyArray3%
Comment  = %MyArray4%


;MsgBox, %Account% - %Reason%

WinActivate, Inventory
WinWaitActive, Inventory
MouseMove, 865, 207 ; Inventory button
Sleep, 500
Click, 868,251 ; Search button
Sleep, 500


sleep, 3000
Loop, {
mouseclickdrag, left, 36, 278, 213, 278
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}

Sleep, 1000
Send, {Tab}

Sleep,500
Send, {Up} ; Go down to RES ID
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
MouseMove, 96,424
Sleep, 500
Click, 96,424
Click, 96,424
Click, 96,424



Sleep, 4000


Loop, {
sleep, 1000
mouseclickdrag, left, 45, 278, 190, 278
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
mouseclickdrag, left, 453, 425, 508, 425
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


sleep, 3000
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

Send, {TAB}
Sleep, 100

Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {Enter}
sleep, 5000




; Wait for New Request Window screen to pop up
;Loop, 
	;{
	;ImageSearch, FX, YX, 318,277, 541,358, %Imagelocation%Request Response Image.bmp
	;If ErrorLevel = 0
	;{
		;MsgBox, New Request Window loaded.
		;Sleep, 1000
		;Break
		;}
	;}
;NEW REQUEST APPROVAL

Send, {TAB}
Sleep, 500
Send, {DOWN}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500

Send, %MyArray3%
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
Sleep, 5000


sleep, 1000
Loop, {
mouseclickdrag, left, 36, 278, 213, 278
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Comment = Request Approved
Goto, Excel
}
}


; Wait for New Request Window screen to disappear
;Loop,
	;{
	;ImageSearch, FX, YX, 318,277, 541,358, %Imagelocation%Request Response Image.bmp
	;If ErrorLevel = 1
		;{
		;MsgBox, New Request Window closed.
		;Comment = Request Approved
		;GoTo, EXCEL
		;}
	;}


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
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause