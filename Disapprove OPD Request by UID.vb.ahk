#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO

MsgBox, In your Excel workbook, select the cell in column F next to the account you want to start with. Now log into Direct Connect.  Column A should contain the UID# and column B should contain a Response Comment. 
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
Note = %MyArray2% 



;MsgBox, %Account% - %Reason%

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
;REQUEST APPROVAL

Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Down}
Sleep, 100
Send, {Down}
Sleep, 100
Send, {Down}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, %MyArray2%
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
Sleep, 2000


sleep, 2000
Loop, {
mouseclickdrag, left, 35, 277, 212, 277
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Comment = Request Disapproved
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