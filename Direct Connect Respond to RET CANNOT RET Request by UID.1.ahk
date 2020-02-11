#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column F next to the account you want to start with. Now log into Direct Connect.  Column A should contain the UID#, column B Response Subject of "Unable to Set Up Retraction", column c should contain New request category of "Financial", column D should contain new request subject of "Please Process Refund Check" and column E should contain "Comments".
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
Category =
Subject =

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
Category = %MyArray3%
Subject  = %MyArray4%
Comment = %MyArray5%

;MsgBox, %Account% - %Reason%

WinActivate, Windows Internet Explorer
WinWaitActive, Windows Internet Explorer
MouseMove, 66,244 ; Inventory button
Sleep, 500
Click, 65, 273 ; Search button
Sleep, 500
;Mousemove, 10, 10

; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 15,265, 175, 340, %Imagelocation%DC Search Inventory Image.bmp
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
mouseclickdrag, left, 280, 290, 16, 290
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}

Click, 105,325 ; Click to select 
Sleep, 500
Click, 105,325
Sleep, 500
Send, {up} ; Go down to Unique ID
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
MouseMove, 88,437
Sleep, 100
Click, 88,437
Click, 88,437


; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 15,265, 180, 315, %Imagelocation%DC Account View Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Account information page loaded.
		;Sleep, 1000
		;Break
		;}
	;}
Sleep, 2000


Loop, {
sleep, 1000
mouseclickdrag, left, 41, 290, 145, 290
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Image found.
Break
}
}

sleep, 3000

Send, {End}

Sleep, 100
ImageSearch, FX, YX, 51,148, 358, 768, %Imagelocation%RET Subject Link.bmp
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

; Wait for Request/ Response Account screen to pop up
Loop, 
	{
	ImageSearch, FX, YX, 335,310, 495,393, %Imagelocation%Request Response Image.bmp
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

;MouseMove, 537,587
;Sleep, 100
;Click, 537,587 ; click to select reason drop down menu
Send, {TAB}
Sleep, 500
Send, %Reason%
Sleep, 500
Send, {Enter} ; Select Response reason
Sleep, 500

Send, {Tab}
Sleep, 500
Send, %MyArray5%
Sleep, 500
Send, {Tab}
Sleep, 500

Send, {Enter}
Sleep, 500


; Wait for Request/ Response Account screen to disappear
;Loop,
	;{
	;ImageSearch, FX, YX, 335,310, 495,370, %Imagelocation%Request Response Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Request Response screen closed.
		;Sleep, 1000
		;Break
		;}
	;}


; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 15,265, 180, 315, %Imagelocation%DC Account View Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Account information page loaded.
		;Sleep, 1000
		;Break
		;}
	;}

sleep, 1000
Loop, {
mouseclickdrag, left, 41, 290, 145, 290
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Image found.
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
Send, {Enter}
sleep, 1000


;Send, {End}
Sleep, 500


;NEW REQUEST
;ImageSearch, FX, YX, 51,152, 281, 764, %Imagelocation%BI New Request Button.bmp
;If ErrorLevel = 0
	{
	;MsgBox, BI New Request Button available.
	
	;Sleep, 100
	;MouseMove, %FX%,%YX% ; Open Request
	;Click, 10, 10, relative
	}
;If ErrorLevel = 1
;MsgBox, BI New Request Button not available

; Wait for New Request Window screen to pop up
Loop, 
	{
	ImageSearch, FX, YX, 411,357, 555, 423, %Imagelocation%BI New Request Window.bmp
	If ErrorLevel = 0
	{
		;MsgBox, New Request Window loaded.
		Sleep, 1000
		Break
		}
	}
;CATEGORY
MouseMove, 587,425
Sleep, 100
Click, 587,425 ; click to select category drop down menu
Sleep, 500
Send, %Category%
Sleep, 500
Send, {Enter}
Sleep, 2000

;SUBJECT
MouseMove, 595,458
Click, 595,458 ; Select Response reason
;Send, {TAB} ; Select Response reason
Sleep, 500
send, %Subject%
sleep, 1000
Send, {Enter}
sleep, 1000

Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, %MyArray5%

Sleep, 500
Send, {Tab}
Send, {Enter}

Sleep, 1500
; Wait for New Request Window screen to disappear
Loop,
	{
	ImageSearch, FX, YX, 398,370, 553,450, %Imagelocation%BI New Request Window.bmp
	If ErrorLevel = 1
		{
		;MsgBox, New Request Window closed.
		Comment = New Request Entered
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