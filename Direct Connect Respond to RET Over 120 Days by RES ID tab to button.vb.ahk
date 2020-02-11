#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column D next to the account you want to start with. Column A should contain UID, column B should contain New Request Category of "Financial", column C should contain New Request Subject of "Please Process Refund Check" and column D should contain New Request Comments. Now log into Direct Connect under UHC org.
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

		 WinActivate, Inventory
		 WinWaitActive, Inventory
		 MouseMove, 865, 207 ; Inventory button
		 Sleep, 500
		 Click,  868,251 ; Search button

			  sleep, 5000
			  Loop, {
					  mouseclickdrag, left, 35, 277, 212, 277
					 Send, ^c
					 clipwait, 
					 If clipboard contains Search Inventory

						{
						  ;MsgBox, Search Inventory page loaded.
						  Break
						}
					}
			 
			Sleep, 500
			Send, {Tab}
			Sleep, 500
			Send, {Home} ; Go up to Unique ID
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
					MouseMove, 104,426
					Sleep, 500
					Click, 104,426
					Click, 104,426
					Click, 104,426

					
			
			;Sleep, 2000
			;Clipboard =
			
			;Loop, {
					;mouseclickdrag, left, 45, 278, 189, 278
					;Send, ^c
					;clipwait, 
					;If clipboard contains Account View

						;{
						  ;MsgBox, Search Inventory page loaded.
						  ;Break
						;}
					;}
					
					
			Sleep, 5000
			Clipboard = 

		Loop, {
							mouseclickdrag, left, 458, 427, 506, 427
							Send, ^c
							clipwait, 
							If clipboard contains Complete
					
						    {
								;MsgBox, Image found.
								comment = Complete
								
								Sleep, 5000
			Clipboard =
			
							mouseclickdrag, left, 45, 278, 189, 278
							Send, ^c
							clipwait, 
							If clipboard contains Account View
								
		
		Sleep, 500
		Send, {Tab}
		Sleep, 200
		Send, {Tab}
		Sleep, 200
		Send, {Enter}
		
		Sleep, 5000
								goto, EXCEL
							}
								else
									break

			}
		
		Sleep, 5000
			Clipboard =
			
							mouseclickdrag, left, 45, 278, 189, 278
							Send, ^c
							clipwait, 
							If clipboard contains Account View
								
		
		Sleep, 500
		Send, {Tab}
		Sleep, 200
		Send, {Tab}
		Sleep, 200
		Send, {Enter}
		
		Sleep, 5000
		
					
							mouseclickdrag, left, 45, 278, 189, 278
							Sleep, 500
							Send, ^c
							clipwait, 
							If clipboard contains Account View
		
							
		Sleep, 1000
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

		
			;CATEGORY
		
		Sleep, 1000
		
		Send, {Tab}
		Sleep, 1000
		Send, %MyArray2%
		Sleep, 1000
		Send, {tab}
		Sleep, 1000
		Send, %MyArray3%
		Sleep, 1000
		Send, {tab}
		Sleep, 1000
		Send, {tab}
		Sleep, 1000
		Send, {tab}
		Sleep, 1000
		Send, %MyArray4%
		Sleep, 1000
		Send, {tab}
		Sleep, 1000
		Send, {tab}
		Sleep, 1000
		Send, {Enter}
		Sleep, 2000

sleep, 1000
Loop, {
mouseclickdrag, left, 36, 278, 213, 278
Send, ^c
clipwait, 
If clipboard contains Search Inventory
	

		
				
		; Wait for New Request Window screen to disappear
		;Loop,
	;{
		;ImageSearch, FX, YX, 455, 234, 593, 293, %Imagelocation%BI New Request Window.bmp
		;If ErrorLevel = 1
		{
			;MsgBox, Request Response screen closed.
			Comment = Request Entered
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
SoundPlay *48
MsgBox, %Count% accounts resolved. `nTime started: %TimeBegin%`nTime completed: %TimeEnd%
ExitApp
ESC::Pause