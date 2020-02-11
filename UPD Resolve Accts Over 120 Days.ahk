#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column F next to the account you want to start with. Now log into Direct Connect.  Column A should contain the Account Resolution ID, column B Resolved Reason of "Created in Error", column c CBO Organization, column D Payor Organization, column E should be blank, .
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECK:
IfWinNotExist, Microsoft Excel
	{
	MsgBox, Ensure Excel is open.
	GoTo, CHECK	
	}

IfWinNotExist,  Internet Explorer
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

Clipboard = 

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
Account = %MyArray1% 
Category = %MyArray2%


;MsgBox, %Account% - %Reason%

WinActivate, Internet Explorer
WinWaitActive, Internet Explorer
MouseMove, 963, 203 ; Inventory button
Sleep, 500
mousemove, 986,248
sleep, 500
Click, 986, 248 ; Search button
Sleep, 500
;Mousemove, 10, 10

; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 28,256, 156, 256, %Imagelocation%IE11 DC Search Inventory Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Search Inventory page loaded.
		;Break
		;}
;if Errorlevel = 1
;Msgbox, Image not found
	;}

mousemove, 70, 240
click, 70, 240
sleep, 500
Loop, {
mouseclickdrag, left, 34, 275, 212, 275
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}

Click, 96, 318 ; Click to select 
Sleep, 500
Click, 96, 318
Sleep, 500
Send, {up} ; Go up to Resolution ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 500
Send, %Account%
Sleep, 500
Send, {Enter}

; Wait for page to load
Loop, 
	{
	;PixelSearch, FX, YX, 33,425, 140,450, 0xDDCCFF, 5, Fast
	;If ErrorLevel = 1
		{
		;MsgBox, Account # results loaded.
		;Sleep, 500
		Break
		}
	}
;this clicks hyper link for acct
MouseMove, 65, 430
Sleep, 500
Click, 65, 430
Click, 65, 430


; Wait for page to load
;Loop, 
	;{
	;ImageSearch, FX, YX, 50, 254, 153, 254, %Imagelocation%IE11 DC Account View Image.bmp
	;If ErrorLevel = 0
		;{
		;MsgBox, Account information page loaded.
		;Sleep, 500
		;Break
		;}
	;}
Sleep, 1000

Loop, {
mouseclickdrag, left, 453, 427, 506, 427
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

Loop, {
sleep, 500
MouseMove, 50, 240
Sleep, 100
mouseclickdrag, left, 45, 280, 190, 280
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Image found.
Break
}
}

Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 500


;}






;Wait for New Request Window screen to pop up
;Loop, 
	;{
;ImageSearch, FX, YX, 460,478, 576, 486, %Imagelocation%IE11 DC Resolved Button.bmp
	;If ErrorLevel = 0
	    ;{
		;MsgBox, New Request Window loaded.
		;Sleep, 500
		;Break
		;}
	;}
;CATEGORY
MouseMove, 615, 505
Sleep, 500
Click, 615, 505 ; click to select category drop down menu
Sleep, 500
Send, %MyArray2%
Sleep, 500
Send, {Enter}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}
Sleep, 1500


; Wait for New Request Window screen to disappear
;Loop,
	;{
;	ImageSearch, FX, YX, 398,370, 553,450, %Imagelocation%BI New Request Window.bmp
	;If ErrorLevel = 1
		;{
		;MsgBox, New Request Window closed.
		;Comment = Account Resolved
		;GoTo, EXCEL
		;}
	;}

Loop, {
sleep, 500
MouseMove, 50, 240
Sleep, 100
mouseclickdrag, left, 45, 280, 190, 280
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Image found.
Break
}
}

Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 500

EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, Account Resolved
Sleep, 500
Send, {Down}
Sleep, 500

}

FormatTime, TimeEnd,, Time
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause