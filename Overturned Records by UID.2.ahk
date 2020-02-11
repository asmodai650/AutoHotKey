#SingleInstance, Force
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO

MsgBox, In your Excel workbook, select the cell in column D next to the account you want to start with. Column A should contain UID, column B should contact Description of Resolution and column C should contain Resolved Reason of Created in Error. Now log into Direct Connect under UHC org.
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
Appeal_Comment =
Resolved_Reason =

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, +{Home}
Sleep, 500
Send, ^c
ClipWait
StringSplit, MyArray, clipboard, %A_Tab%
Account = %MyArray1% 
Appeal_Comment = %MyArray2%
Resolved_Reason = %MyArray3%




WinActivate, Windows Internet Explorer
WinWaitActive, Windows Internet Explorer
MouseMove, 70,268 ; Inventory button
Sleep, 500
Click, 64, 294 ; Search button
Sleep, 1000


sleep, 1000
			  Loop, {
					  mouseclickdrag, left, 30, 313, 158, 313
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

Sleep,500
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
MouseMove, 81,457
Sleep, 100
Click, 81,457
Click, 81,457
Click, 81,457

sleep, 1000
Clipboard =
			  Loop, {
					  mouseclickdrag, left, 42, 313, 145, 313
					 Send, ^c
					 clipwait, 
					 If clipboard contains Account View

						{
						  ;MsgBox, Search Inventory page loaded.
						  Break
						}
					}

Sleep, 2000
Clipboard = 

Loop, {
mouseclickdrag, left, 440, 456, 483, 456
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


sleep, 1000
Loop, {
mouseclickdrag, left, 42, 313, 145, 313
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}
Sleep, 500
Send, {End}
Sleep, 500

;this clicks manual reassign
MouseMove, 381,766
Sleep, 100
Click, 381,766
Sleep, 500

Send, {Home}
Sleep, 500

sleep, 1000
Loop, {
mouseclickdrag, left, 42, 313, 145, 313
Send, ^c
clipwait, 
If clipboard contains Account View

			{
;MsgBox, Account View page loaded.
Break
			}
		}


Sleep, 500
Send, {Tab}
Sleep, 100
Send, {Enter} ;Saves Edit
Sleep, 5000


Sleep, 1000
Clipboard = 
Loop, 
   {
        mouseclickdrag, left, 41, 314, 145, 314
        Send, ^c
        clipwait, 
	    If clipboard contains Account View

 {
        ;MsgBox, Account View page loaded.
        Break
  }
    }


Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {Enter}


Sleep, 2000
Clipboard =

sleep, 1000
Loop, {
mouseclickdrag, left, 371, 592, 470, 592
Send, ^c
clipwait, 
If clipboard contains Response Reason
	
{
;MsgBox, Image found.
Break
}
}


Sleep, 500
Send, {TAB}
Sleep, 500

Send, {Down}
Sleep, 500
Send, {Enter} ; Select Response reason
Sleep, 500

Send, {Tab}
Sleep, 500
Send, %MyArray2%
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}

Sleep, 2000
Clipboard =


sleep, 5000
Loop, {
mouseclickdrag, left, 42, 313, 145, 313
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

; RESOLVING CLAIM
Sleep, 100
Send, {TAB}
Sleep, 200
Send, {Enter}

Sleep, 500
Send, {TAB}
Sleep, 200
Send, %MyArray3%
Sleep, 200
Send, {TAB}
Sleep, 200
Send, {ENTER}


Sleep, 1000

sleep, 1000
Clipboard =
Loop, {
mouseclickdrag, left, 30, 314, 158, 314
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Comment = Resolved
goto , EXCEL
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