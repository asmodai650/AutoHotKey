#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, Optum provisioning must start on the Users tab. In your Excel workbook, select the cell in first empty column next to the account you want to start with. Now log into Optum Provisioning and select the Users tab.  Column A should contain the User name to be removed from the Direct Connect group in provisioning.

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
;if A_Min between 00 and 01
;	sleep, 60000
;if A_Min between 15 and 16
;	sleep, 60000
;if A_Min between 30 and 31
;	sleep, 60000
;if A_Min between 45 and 46
;	sleep, 60000

TrayTip,,%a_index% of %Count%,30	

User =
 

Clipboard = 

	
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
User = %MyArray1% 



WinActivate, Provisioning
WinWaitActive, Provisioning

;Copies search above search section
Loop, {
MouseMove, 50, 400
Sleep, 100
mouseclickdrag, left, 50, 400, 101, 400
Send, ^c
clipwait, 
If clipboard contains Search

{
Break
}
}

;tabs to user name box
Send, {Tab}
Sleep, 500
;enters user name in field
Send, %MyArray1%
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Tab}
Sleep, 500
;hits enter on search button
Send, {Enter}
Sleep, 3000

MouseMove, 189,714

Loop, 
	{
	PixelSearch, FX, YX, 189,714, 189,714, 0xCCCCCC, 5, Fast
	;MsgBox, %ErrorLevel%
	;MsgBox, 
	
	If ErrorLevel = 0
		{
		;MsgBox, gray found
				Goto, EXCEL
		
		;MsgBox, gray not found
		Sleep, 1000
	    }	
	else
		
		Break
		
		;
	}



MouseMove, 54, 779 ; Moves to user name
Sleep, 1000
sleep, 500
Click, 54, 779 ; Clicks User name
Sleep, 3000

MouseMove, 363, 618 ; Moves to Remove Button
Sleep, 1000
sleep, 500
Click, 363, 618 ; Clicks Remove Button
Sleep, 3000

Send, {Enter} ; Enter to select Yes on "are you sure" window








Sleep, 3000

;if A_Min between 00 and 01
;	sleep, 60000
;if A_Min between 15 and 16
;	sleep, 60000
;if A_Min between 30 and 31
;	sleep, 60000
;if A_Min between 45 and 46
;	sleep, 60000




EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, User Removed from Direct Connect group
Sleep, 500
Send, {Down}
Sleep, 500

}

FormatTime, TimeEnd,, Time
SoundPlay *48
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause