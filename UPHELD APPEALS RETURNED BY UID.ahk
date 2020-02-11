#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

MsgBox, In your Excel workbook, select the cell in column E next to the account you want to start with. Column A should contain UID, column B should contain Response Reason "Additional Information Supplied", Column C should contain New Request Category of "Approval" and column D should contain New Request Comments. Now log into Direct Connect under UHC org.
InputBox, Count, How many accounts?, How many accounts do you want to resolve?

CHECKEXCEL:
IfWinNotExist, Microsoft Excel
	{
		MsgBox, Ensure Excel is open.
		GoTo, CHECKEXCEL	
	}

CHECKDCPAGE:
IfWinNotExist,  Inventory
	{
		MsgBox, Ensure the home Direct Connect page is open.
		GoTo, CHECKDCPAGE	
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
		 MouseMove, 976, 203 ; Inventory button
		 Sleep, 500
		 Click, 990, 247 ; Search button

			  sleep, 1000
			  Loop, {
					  mouseclickdrag, left, 34, 278, 209, 278
					 Send, ^c
					 clipwait, 
					 If clipboard contains Search Inventory

						{
						  ;MsgBox, Search Inventory page loaded.
						  Break
						}
					}
			;Click, 150,325 ; Click to select 
			Sleep, 500
			;Click, 150,325
			Send, {Tab}
			Sleep, 500
			Send, {Up} ; Go up to Unique ID
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
					MouseMove, 96, 424
					Sleep, 100
					Click, 96, 424
					Click, 96, 424
					Click, 96, 424

								
			Sleep, 1000
			
				;Looks to see if the account screen has loaded before continuing. 
				NEWSCREEN:
							mouseclickdrag, left, 440, 428, 506, 428
							Send, ^c
							;clipwait, 
							If clipboard <> In Process
								{
								If clipboard <> Complete
									Loop
										{
										Sleep, 1000
										GoTo, NEWSCREEN
										}
									else
										{
										GoTo, SCREENFOUND
										}
								{
								GoTo, SCREENFOUND
								}
								}
			SCREENFOUND:
			
			Clipboard = 

		Loop, {
							mouseclickdrag, left, 458, 428, 506, 428
							Send, ^c
							clipwait, 
							If clipboard contains Complete
					
						    {
								;MsgBox, Image found.
								comment = Complete
								
								Sleep, 2000
			Clipboard =
			
							mouseclickdrag, left, 46, 279, 185, 279
							Send, ^c
							clipwait, 
							If clipboard contains Account View
								
		
		Sleep, 500
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
			
							mouseclickdrag, left, 46, 279, 185, 279
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
		Send, {TAB}
		Sleep, 100
		Send, {Enter}
		
		if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000


		Sleep, 3000
		Send, {Tab}
		Sleep, 100
		Send, {Tab}
		Sleep, 100
		Send, %MyArray2%
		Sleep, 100
		Send, {Tab}
		Sleep, 100
		Send, {Tab}
		Sleep, 100
		Send, {Enter}
		
		Sleep, 5000
		
					
							mouseclickdrag, left, 46, 279, 185, 279
							Sleep, 500
							Send, ^c
							clipwait, 
							If clipboard contains Account View
		
							
	if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000						
							
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

	;CATEGORY
		
		Sleep, 5000
		
		Send, {Tab}
		Sleep, 1000
		Send, %MyArray3%
		Sleep, 1000
		Send, {tab}
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

		Comment = Request Entered

	WinWaitActive, Inventory
	
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