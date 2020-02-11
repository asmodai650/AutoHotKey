#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO

MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect and go to the Provider whose claims you are working.  Column A should contain the UID# and column B New request Subject of "Please Approve Overpayment".
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
	;MsgBox, Ensure Excel is open.
	GoTo, CHECK	
	}

IfWinNotExist,  Internet Explorer
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
FinancialRequest =

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
UID = %MyArray1% 
FinancialRequest = %MyArray2% 


;Wait for Direct Connect refresh
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
MouseMove, 978,200 ; Inventory button
Sleep, 500
Click, 978, 240 ; Search button
Sleep, 500


sleep, 1000

Loop,
	{
	mouseclickdrag, left, 37, 278, 215, 278
	Send, ^c
	clipwait, 
	If clipboard contains Search Inventory
		{
		;MsgBox, Search Inventory page loaded.
		Break
		}
	}

Sleep, 500
Send, {Tab}
Sleep,500
Send, {Down} ; Go down to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 500
Send, %UID%
Sleep, 500
Send, {Enter}

; Wait for page to load, then confirm correct screen
Sleep, 500

Loop, 
	{
	PixelSearch, FX, YX, 33,425, 140,450, 0xDDCCFF, 5, Fast
	If ErrorLevel = 1
		{
		;MsgBox, UID Found.
		Sleep, 1000
		Break
		}
	}

;this clicks hyper link for acct
MouseMove, 63, 430
Sleep, 100
Click, 63, 430
Click, 63, 430
Click, 63, 430

Sleep, 2000

Loop,
	{
	sleep, 1000
	mouseclickdrag, left, 47, 284, 187, 284
	Send, ^c
	clipwait, 
	Sleep, 200
	If clipboard contains Account View
		{
		;MsgBox, Image found.
		Break
		}
	else
		{
		sleep, 500
		;msgBox, UID NOT Found
		Comment = UID Not Found
		goto, EXCEL
		}
	break
	}
	

Sleep, 1000
Clipboard = 

Loop,
	{
	mouseclickdrag, left, 455, 427, 490, 427
	Send, ^c
	clipwait, 
		sleep, 200
		If clipboard contains In Process
			{
			;MsgBox, Image found.
			break
			}
		;If Account Complete, Resolved, or In Process is not found, return to Excel with Direct Connect Account Status
		else
		comment = %clipboard%
		Sleep, 2000
		Clipboard =
			
		mouseclickdrag, left, 47, 284, 187, 284
		Send, ^c
		clipwait,
		If clipboard contains Account View		
			Sleep, 500
			Send, {Tab}
			Sleep, 200
			Send, {Enter}
			Sleep, 2000		
			goto, EXCEL
		break
	}

;Look to see if page is loaded
sleep, 1000
Loop,
	{
	mouseclickdrag, left, 47, 284, 187, 284
	Send, ^c
	clipwait, 
	If clipboard contains Account View
		{
		;MsgBox, Account View page loaded.
		Break
		}
	}

;Tab Down page to Open Request
Send, {TAB 7}
Sleep, 200
;Send, {TAB}
;Sleep, 100
;Send, {TAB}
;Sleep, 100
;Send, {TAB}
;Sleep, 100
;Send, {TAB}
;Sleep, 100
;Send, {TAB}
;Sleep, 100
;Send, {TAB}
;Sleep, 100
Send, {Enter}
sleep, 2000

; Go to Response Reason in Request / Response Window, select Approved then the Financial Request from Array2
Send, {TAB 2}
Sleep, 500
;Send, {TAB}
;Sleep, 500
Send, {down}
Sleep, 500
Send, {TAB}
Sleep, 500
;If Financial Request is anything other than Please Set Up Retraction, change next two lines
;Send, {Down}
;Sleep, 500
Send, {TAB 2}
Sleep, 500
;Send, {TAB}
;Sleep, 500

;Wait for Direct Connect refresh
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

;Goes to Save and close
Send, {TAB}
Sleep, 500
Send, {Enter}
Sleep, 2000

;Returns to Excel
sleep, 2000
Loop,
	{
	mouseclickdrag, left, 37, 278, 215, 278
	Send, ^c
	clipwait,
	If clipboard contains Search Inventory
		{
		;MsgBox, Search Inventory page loaded.
		Comment = Done
		Goto, Excel
		}
	Else
	
	clipboard = An error has occurred
	sleep = 500
		{
		;MsgBox, Search Inventory page loaded.
		clipboard = 
		sleep, 1000
		MouseMove, 978,200 ; Inventory button
		Sleep, 500
		Click, 978, 200 ; Search button
		Sleep, 500
		Comment = Request Incomplete
		Goto, Excel
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
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause