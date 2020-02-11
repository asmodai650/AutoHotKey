#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI Macro

MsgBox, In your Excel workbook, select the cell in column D next to the account you want to start with. Column A should contain UID, column B should contain Response Subject and column C should contain Response Comments. Now log into Direct Connect under UHC org.
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
Comment = %MyArray3%

;MsgBox, %MyArray2%
;MsgBox, %MyArray3%

WinActivate, Windows Internet Explorer
WinWaitActive, Windows Internet Explorer
MouseMove, 65,218 ; Inventory button
Sleep, 500
Click, 63, 249 ; Search button

; Wait for page to load
Loop, 
	{
	ImageSearch, FX, YX, 20,250, 182, 284, %Imagelocation%DC Search Inventory Image.bmp
	If ErrorLevel = 0
		{
		;MsgBox, Search Inventory page loaded.
		Break
		}
	}
Click, 115,305 ; Click to select 
Sleep, 500
Click, 115,305
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



Send, {End}
Sleep, 100
ImageSearch, FX, YX, 22,174, 310, 882, %Imagelocation%RET Subject Link.bmp
If ErrorLevel = 0

	{
	;MsgBox, Subject Link available.
	
	Sleep, 100
	MouseMove, %FX%,%YX% ; Open Request
	Click, 10, 10, relative
	}
If ErrorLevel = 1
{
	{
	send, {pgup}
	sleep, 500
	ImageSearch, FX, YX, 22,174, 310, 882, %Imagelocation%RET Subject Link.bmp

		If ErrorLevel = 0

		{
		;MsgBox, Subject Link available.
		
		Sleep, 100
		MouseMove, %FX%,%YX% ; Open Request
		Click, 10, 10, relative
		}
	
	
		If ErrorLevel = 1
	
		{
		
		
		send, {Home}
		Comment = Request Link Not Available
			GoTo, EXCEL
		}
	}	
  }

; Wait for Request/ Response Account screen to pop up
Loop, 
	{
	ImageSearch, FX, YX, 335,310, 495,370, %Imagelocation%Request Response Image.bmp
	If ErrorLevel = 0
	{
		;MsgBox, Request Response screen loaded.
		Sleep, 1000
		Break
		}
	}

sleep, 1000
Loop, {
mouseclickdrag, left, 371, 576, 469, 576
Send, ^c
clipwait, 
If clipboard contains Response Reason
	
{
;MsgBox, Image found.
Break
}
}



;MouseMove, 528,581
;Sleep, 100
;Click, 528,581 ; click to select reason drop down menu
Send, {TAB}
Sleep, 500
Send, %MyArray2%
Sleep, 500
Send, {Enter} ; Select Response reason
Sleep, 500
;Send, {Enter} ; Save selection and close pop up window
;Sleep, 500
Send, {Tab}
Sleep, 500
Send, %MyArray3%
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}

; Wait for Request/ Response Account screen to disappear
Loop,
	{
	ImageSearch, FX, YX, 335,310, 495,370, %Imagelocation%Request Response Image.bmp
	If ErrorLevel = 0
		{
		;MsgBox, Request Response screen closed.
		Comment = Request Responded
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