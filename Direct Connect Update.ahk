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
Click, 60,435
;Click, 60,445


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
ImageSearch, FX, YX, 790,265, 1000, 400, %Imagelocation%DC Resolve Button Image.bmp
If ErrorLevel = 0
	{
	;MsgBox, Resolve button available.
	MouseMove, 928,295
	Sleep, 100
	Click, 928,295 ; Resolve button
	}
If ErrorLevel = 1
	{
	;MsgBox, Resolve button NOT available.
	Comment = Resolve button unavailable.
	GoTo, EXCEL
	}

; Wait for Resovle Account screen to pop up
Loop, 
	{
	ImageSearch, FX, YX, 390,460, 820,570, %Imagelocation%DC Resolve Account Image.bmp
	If ErrorLevel = 0
	{
		;MsgBox, Resolve account screen loaded.
		Sleep, 1000
		Break
		}
	}
Click, 648,538 ; click to select reason drop down menu
Sleep, 500
Send, %Reason%
Sleep, 500
Send, {Enter} ; Select reason code
Sleep, 500
;Send, {Enter} ; Save selection and close pop up window
;Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}


; Wait for Resovle Account screen to disappear
Loop,
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