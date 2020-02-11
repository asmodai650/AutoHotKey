#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO

MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect.  Column A should contain the UID#, column B Request Note. Log into Direct Connect, and go to the Provider side.
InputBox, Count, How many accounts?, Please enter the amount of claims that you want to resolve.

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
TrayTip,,%a_index% of %Count%,30	
clipboard = 

UID =
Note =

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
UID = %MyArray1% 
Note = %MyArray2% 


WinActivate, Inventory
WinWaitActive, Inventory
MouseMove, 979, 202 ; Inventory button
Sleep, 500
Click, 984, 245 ; Search button
Sleep, 500


sleep, 1000
Loop, {
	mouseclickdrag, left, 35, 280, 211, 280
	Send, ^c
	clipwait, 
If clipboard contains Search Inventory

{
	Break
}
}

Sleep, 500
Send, {Tab}

Sleep,500
Send, {down} ; Go down to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 500
Send, %UID%
Sleep, 1000
Send, {Enter}

; Wait for page to load
;Loop, 
	;{
	;PixelSearch, FX, YX, 33,425, 140,450, 0xDDCCFF, 5, Fast
	;If ErrorLevel = 1
		;{
		;MsgBox, Account # results loaded.
		;Sleep, 1000
		;Break
		;}
	;}
;this clicks hyper link for acct
MouseMove, 92, 431
Sleep, 100
Click, 92, 431
Click, 92, 431
Click, 92, 431

sleep, 1000

Sleep, 1000
Clipboard = 

Loop, {
mouseclickdrag, left, 458, 428, 506, 428
Send, ^c
clipwait, 
If clipboard contains Complete
	
{
;MsgBox, Image found.
comment = Account Complete

Sleep, 2000
			Clipboard =
			
							mouseclickdrag, left, 46, 279, 185, 279
							Send, ^c
							clipwait, 
							If clipboard contains Account View
								
		
		Sleep, 500
		Send, {Tab}
		Sleep, 200
		Send, {Enter}
		sleep, 1500
goto, EXCEL
}

Else if clipboard contains Resolved
;If Else clipboard contains Resolved
{
comment = Account Resolved

Sleep, 1000
			Clipboard =
			
							mouseclickdrag, left, 46, 279, 185, 279
							Send, ^c
							clipwait, 
							If clipboard contains Account View
								
		
		Sleep, 500
		Send, {Tab}
		Sleep, 200
		Send, {Tab}
		Sleep, 200
		Send, {Tab}
		Sleep, 200
		Send, {Enter}
		sleep, 1500
goto, EXCEL
}

else
	break

}


sleep, 2000
Loop, {
mouseclickdrag, left, 56, 279, 198, 279
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

;Approve 1st Request
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
sleep, 500


if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

Sleep, 2000
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Down}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
Sleep, 2000

sleep, 500
Loop, {
mouseclickdrag, left, 56, 279, 198, 279
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

;Approve 2nd Request
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
sleep, 500

if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

Sleep, 2000
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {R}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, %Note%
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
Sleep, 2000

sleep, 500
Loop, {
mouseclickdrag, left, 35, 280, 211, 280
Send, ^c
clipwait, 1
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Comment = Check info entered
Goto, Excel
}

Else if clipboard contains has occurred.
{
;MsgBox, Search Inventory page loaded.
Comment = Request Incomplete
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