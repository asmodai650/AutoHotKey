#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\dcurtis1\Desktop\BI MACRO\

MsgBox, Direct Connect must start in the Manage users section under the Administration tab. In your Excel workbook, select the cell in column C next to the account you want to start with. Now log into Direct Connect.  Column A should contain the Email address to be inactivated and column B the email address with the word INACTIVE inserted.

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
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

TrayTip,,%a_index% of %Count%,30	

Email =
NewEmail = 

Clipboard = 

	
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
Email = %MyArray1% 
NewEmail = %MyArray2%


WinActivate, Direct Connect
WinWaitActive, Direct Connect
MouseMove, 529, 398 ; Filter icon for email
Sleep, 1000
;mousemove, 1193, 382
sleep, 500
Click, 529, 398 ; Manage users
Sleep, 3000

Send, {Tab}
Sleep, 500
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
Send, {Enter}

					


MouseMove, 489, 429 ;this clicks line to take to edit screen
Sleep, 5000
Click, 489, 429
Click, 489, 429

; Wait for page to load
Sleep, 5000


Loop, {
MouseMove, 47, 279
Sleep, 100
mouseclickdrag, left, 47, 279, 141, 279
Send, ^c
clipwait, 
If Clipboard contains Users
goto, EXCEL

sleep, 1000

;If clipboard contains Edit User




{
	
Break
}
}

If clipboard contains Edit User
	

Sleep, 3000
Clipboard =


Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, %MyArray2%
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 5000

Loop, {
MouseMove, 47, 279
Sleep, 100
mouseclickdrag, left, 47, 279, 141, 279
Send, ^c
clipwait, 
If clipboard contains Edit User

{
Break
}
}

Sleep, 5000


if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000


;checking Inactive box
			MouseMove, 189, 490
			Click, 189, 490



Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, {Enter}
sleep, 5000



EXCEL:
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, Account Inactivated
Sleep, 500
Send, {Down}
Sleep, 500

}

FormatTime, TimeEnd,, Time
SoundPlay *48
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause