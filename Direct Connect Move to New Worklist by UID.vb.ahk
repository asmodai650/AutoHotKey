#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column C next to the account you want to start with. Column A should contain UID and column B should contain new worklist. Now log into Direct Connect under UHC org.
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

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
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

	
TrayTip,,%a_index% of %Count%,30	

Account =
Reason =

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 1500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
Account = %MyArray1% 
New_Worklist = %MyArray2% 


;MsgBox, %MyArray2%
;MsgBox, %MyArray3%

WinActivate, Inventory
WinWaitActive, Inventory
MouseMove, 865, 207 ; Inventory button
Sleep, 1500
Click, 868,251 ; Search button

Sleep, 5000
; Wait for page to load
				Loop, 
					{
					  mouseclickdrag, left, 36, 278, 213, 278
					 Send, ^c
					 clipwait, 
					 
					 If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
					 If clipboard contains Search Inventory

						{
						  ;MsgBox, Search Inventory page loaded.
						  Break
						}
					}
					


;Click, 150,325 ; Click to select 
Sleep, 2000
;Click, 150,325
Send, {Tab}
Sleep, 500
Send, {Up} ; Go up to Unique ID
Sleep, 500
;send, U
;click, 129, 316
Sleep, 500
;Send, {Tab}
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
		Sleep, 3000
		Break
		}
	}
;this clicks hyper link for acct
MouseMove, 96,424
Sleep, 100
Click, 96,424
Click, 96,424
Click, 96,424
Sleep, 5000

; Wait for page to load
Loop, {
			mouseclickdrag, left, 45, 278, 190, 278
			Send, ^c
			clipwait, 
			If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
			If clipboard contains Account View

			{
			;MsgBox, Account View page loaded.
			Break
			}
			}

			Sleep, 5000
			Clipboard =

						 

		;Loop, {
					mouseclickdrag, left, 453, 425, 508, 425
					Send, ^c
					clipwait, 
					If clipboard contains Complete
	
					{
						;MsgBox, Image found.
						comment = Account Complete
						
						mouseclickdrag, left, 45, 278, 190, 278
						Send, ^c
						clipwait, 
						
						If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
						If clipboard contains Account View
						
						Sleep, 1000
						Send, {Tab}
						Sleep, 100
						Send, {Tab}
						Sleep, 100
						Send, {Enter}
						
						goto, EXCEL
					}
						sleep, 5000

			Clipboard = 

			
					mouseclickdrag, left, 45, 278, 190, 278
					Send, ^c
					clipwait, 
					
					If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
					If clipboard contains Account View


;clicks edit button
Sleep, 1000
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}
;Sleep, 500
;Send, {End}
Sleep, 3000

clipboard = 

			
					mouseclickdrag, left, 45, 278, 190, 278
					Send, ^c
					clipwait, 
					
					If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
					If clipboard contains Account View

Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 1000
Send, {Tab}
Sleep, 1000
Send, G
;Send, {Down}
;Sleep, 50
;Send, {Down}
Sleep, 200
Send, {Tab}
Sleep, 500
Send, {Home}
Sleep, 500

sleep, 3000
Loop, {
mouseclickdrag, left, 45, 278, 190, 278
;mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 

If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

;clicks save button
Sleep, 100
Send, {Tab}
Sleep, 200
Send, {Enter}

sleep, 3000
Loop, {
mouseclickdrag, left, 45, 278, 190, 278
;mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 

If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

;clicks cancel button
Sleep, 3000
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}

sleep, 3000
Loop, {
mouseclickdrag, left, 36, 278, 213, 278
Send, ^c
clipwait, 

If clipboard contains An error has occurred
			{
				MsgBox, Error has Occurred
				ExitApp
			}
			
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
Sleep, 1500
Send, {Down}
Sleep, 1500

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