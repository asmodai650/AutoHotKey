#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column D next to the account you want to start with. Column A should contain UID, column B should contain New Request Category of "Financial", column C should contain New Request Subject of "Please Process Refund Check" and column D should contain New Request Comments. Now log into Direct Connect under UHC org.
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
		 Category =
		 Subject  =
		 Comment =

		 WinActivate, Microsoft Excel
		 WinWaitActive, Microsoft Excel
		 Sleep, 2000
		 Send, +{Home}
		 Sleep, 500
		 Send, ^c
		 ClipWait
		 StringSplit, MyArray, clipboard, %A_Tab%
		 Account = %MyArray1% 
		 Category = %MyArray2%
		 Subject  = %MyArray3%
		 Comment = %MyArray4%
		 

		 ;MsgBox, %MyArray2%
		 ;MsgBox, %MyArray3%

		 WinActivate, Internet Explorer
		 WinWaitActive, Internet Explorer
		 MouseMove, 969, 216 ; Inventory button
		 Sleep, 500
		 Click, 974, 262 ; Search button
		 Sleep, 500

			  sleep, 1000
			  Loop, {
					  mouseclickdrag, left, 36, 291, 216, 291
					 Send, ^c
					 clipwait, 
					 If clipboard contains Search Inventory

						{
						  ;MsgBox, Search Inventory page loaded.
						  Break
						}
					}
			Click, 103, 331 ; Click to select 
			Sleep, 500
			Click, 103, 331
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
				;PixelSearch, FX, YX, 33,425, 140,450, 0xDDCCFF, 5, Fast
				;If ErrorLevel = 1
			{
					;MsgBox, Account # results loaded.
					;Sleep, 1000
					Break
			}
	    }
					;this clicks hyper link for acct
					MouseMove, 68, 443
					Sleep, 100
					Click, 68, 450
					Click, 68, 450
					Click, 68, 450 ; REMOVE?

					; Wait for page to load
					;Loop, 
	        {
				;ImageSearch, FX, YX, 15,265, 180, 350, %Imagelocation%DC Account View Image.bmp
				;If ErrorLevel = 0
				{
					;MsgBox, Account information page loaded.
					;Sleep, 1000
					;Break
				}
			}
			Sleep, 2000
			Clipboard = 

		Loop, {
							mouseclickdrag, left, 437, 442, 505, 442
							Send, ^c
							clipwait, 
							If clipboard contains Complete
					
						    {
								;MsgBox, Image found.
								comment = Complete
								goto, EXCEL
							}
								else
									break

			}
		
		
		Sleep, 2000
			Clipboard =
			
			mouseclickdrag, left, 47, 292, 189, 292
							Send, ^c
							clipwait, 
							If clipboard contains Account View
								
		
		Sleep, 500
		Send, {Tab}
		Sleep, 200
		Send, {Enter}
		
		Sleep, 500
		
		mouseclickdrag, left, 47, 292, 189, 292
							Send, ^c
							clipwait, 
							If clipboard contains Account View
		
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
		Send, {TAB}
		Sleep, 100
		Send, {Enter}

		;NEW REQUEST
		; Wait for New Request Window screen to pop up
		;Loop, 
			;{
					 ;ImageSearch, FX, YX, 411,357, 555, 423, %Imagelocation%BI New Request Window.bmp
					 ;If ErrorLevel = 0
				;{
						 ;MsgBox, New Request Window loaded.
						 ;Sleep, 1000
						 ;Break
				;}
			;}
			;CATEGORY
		MouseMove, 620,520
		Sleep, 100
		Click, 620,520 ; click to select category drop down menu
		Sleep, 500
		Send, %MyArray2%
		Sleep, 500
		Send, {Enter}
		Sleep, 2000

		;SUBJECT
		;MouseMove, 595,458
		;Click, 595,458 ; Select Response reason
		Send, {TAB} ; Select Response reason
		Sleep, 500
		send, %MyArray3%
		sleep, 1000
		Send, {Enter}
		sleep, 1000

		Send, {Tab}
		Sleep, 500
		Send, {Tab}
		Sleep, 500
		Send, {Tab}
		Sleep, 500
		Send, %MyArray4%

		;Sleep, 500
		;Send, {Tab}
		Sleep, 500
		Send, {Tab}
		Sleep, 500
		Send, {Enter}

		sleep, 2000
				
		; Wait for New Request Window screen to disappear
		;Loop,
	;{
		;ImageSearch, FX, YX, 411,357, 555, 423, %Imagelocation%BI New Request Window.bmp
		;If ErrorLevel = 1
		;{
			;MsgBox, Request Response screen closed.
			;Comment = Request Entered
			;GoTo, EXCEL
		;}
	;}
		
			
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