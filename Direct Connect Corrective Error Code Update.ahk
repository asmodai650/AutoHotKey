#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect. 
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
Reason = %MyArray2% 
;MsgBox, %Account% - %Reason%

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
Send, {Down} ; Go down to Patien Account Number
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
MouseMove, 60,435
Sleep, 100
Click, 60,435

; Wait for page to load
Loop, 
	{
	ImageSearch, FX, YX, 15,265, 180, 315, %Imagelocation%DC BI Account View Image.bmp
	If ErrorLevel = 0
		{
		;MsgBox, Account information page loaded.
		Sleep, 1000
		Break
		}
	}
Sleep, 500
ImageSearch, FX, YX, 790,265, 1080, 350, %Imagelocation%DC Edit Button Image.bmp
If ErrorLevel = 0
	{
	;MsgBox, Edit button available.
	MouseMove, 1019,294
	Sleep, 100
	Click, 1019,294 ; Edit button
	}
If ErrorLevel = 1
	{
	;MsgBox, Edit button NOT available.
	Comment = Edit button unavailable.
	GoTo, EXCEL
	}
Sleep, 500

Send, {END}
MouseMove, 473,497
	Sleep, 100
	Click, 473,497

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel

Send, {LEFT}
Send, {CTRLDOWN}c{CTRLUP}
Send, {RIGHT}

WinActivate, Windows Internet Explorer
WinWaitActive, Windows Internet Explorer

Send, {CTRLDOWN}v{CTRLUP}

MouseMove, 492,495
	Sleep, 100
	Click, 492,495

Send, {HOME}
	
{
	ImageSearch, FX, YX, 1013,245, 1208, 356, %Imagelocation%DC Save Button Image.bmp
	If ErrorLevel = 0
{
	;MsgBox, Save button available.
	MouseMove, 1104,295
	Sleep, 100
	Click, 1104,295 ; Save button
	}
		{
		;MsgBox, Save Edit screen closed.
		Comment = Account Edited
		GoTo, EXCEL
		}
	}

MouseMove, 1104,295
	Sleep, 100
	Click, 1104,295



	{
	ImageSearch, FX, YX, 15,265, 175, 320, %Imagelocation%DC Search Inventory Image.bmp
	If ErrorLevel = 0

		{
		;MsgBox, Resolve account screen closed.
		Comment = Account resolved
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
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause