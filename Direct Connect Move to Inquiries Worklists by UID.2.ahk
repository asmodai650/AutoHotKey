#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column D next to the account you want to start with. Column A should contain UID and column B should contain new worklist. Now log into Direct Connect under UHC org.
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
MouseMove, 70,245 ; Inventory button
Sleep, 500
Click, 70, 273 ; Search button

; Wait for page to load
Loop, 
	{
	ImageSearch, FX, YX, 15,265, 175, 320, %Imagelocation%DC Search Inventory Image.bmp
	If ErrorLevel = 0
		{
		;MsgBox, Search Inventory page loaded.
		Break
		}
	}
Click, 150,325 ; Click to select 
Sleep, 500
Click, 150,325
Sleep, 500
Send, {Up} ; Go down to Unique ID
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
MouseMove, 59,444
Sleep, 100
Click, 59,444
Click, 59,444
Click, 59,444

; Wait for page to load
Loop, 
	{
	ImageSearch, FX, YX, 15,265, 180, 315, %Imagelocation%DC Account View Image.bmp
	If ErrorLevel = 0
		{
		;MsgBox, Account information page loaded.
		Sleep, 1000
		Break
		}
	}

Sleep, 2000
Clipboard = 

Loop, {
mouseclickdrag, left, 440, 436, 483, 436
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
mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 
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
Sleep, 500
Send, {Tab}
Sleep, 500

Send, {Home}
Sleep, 500

sleep, 1000
Loop, {
mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Sleep, 100
Send, {Tab}
Sleep, 200
Send, {Enter}

sleep, 1000
Loop, {
mouseclickdrag, left, 41, 291, 145, 291
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
Loop, {
mouseclickdrag, left, 30, 290, 159, 290
Send, ^c
clipwait, 
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