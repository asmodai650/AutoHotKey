
SetTitleMatchMode, 2

Imagelocation = C:\Users\vbeam\Desktop\BI MACRO\

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
	;MsgBox, Ensure the home Direct Connect page is open.
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
MouseMove, 78,247 ; Inventory button HOME VERSION
;MouseMove, 70,245 ; Inventory button
Sleep, 500
;Click, 70, 295 ; Search button HOME VERSION
Click, 67, 269 ; Search button


sleep, 1000
Loop, {
;mouseclickdrag, left, 30, 312, 159, 312	; HOME VERSION
mouseclickdrag, left, 30, 290, 159, 290
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}


Send, {Tab}
Sleep, 500
Send, {Up} ; Go down to Unique ID
Sleep, 500
Send, {Tab} ; Tab over to search field
Sleep, 500
Send, %Account%  ; Enters Unique ID
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
;MouseMove, 80,458 ; HOME VERSION
;Sleep, 100 ; HOME VERSION
;Click, 80,458 ; HOME VERSION
;Click, 80,458 ; HOME VERSION
;Click, 80,458 ; HOME VERSION

MouseMove, 59,444
Sleep, 100
Click, 59,444
Click, 59,444
Click, 59,444



Sleep, 2000
Clipboard = 


sleep, 1000
Loop, {
;mouseclickdrag, left, 42, 311, 146, 311 ; HOME VERSION
mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Sleep, 2000
Clipboard =

Loop, {
;mouseclickdrag, left, 442, 457, 485, 457 ; HOME VERSION
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



sleep, 1000
Loop, {
;mouseclickdrag, left, 42, 311, 146, 311 ; HOME VERSION
mouseclickdrag, left, 41, 291, 145, 291
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
Sleep, 500
Send, {Enter}
Sleep, 500


;Sleep, 2000
;Clipboard =

;sleep, 1000
;Loop, {
;mouseclickdrag, left, 414, 463, 513, 463 ; HOME VERSION
;mouseclickdrag, left, 371, 576, 469, 576
;Send, ^c
;clipwait, 
;If clipboard contains Response Reason
	
;{
;MsgBox, Image found.
;Break
;}
;}



Sleep, 500
Send, {TAB}
Sleep, 500
Send, {TAB}
Sleep, 500
Send, %MyArray2%
Sleep, 500
;Send, {Enter} ; Select Response reason
Sleep, 500

Send, {Tab}
Sleep, 500
Send, %MyArray3%
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}

Sleep, 2000
Clipboard =


sleep, 1000
Loop, {
;mouseclickdrag, left, 42, 311, 146, 311 ; HOME VERSION
mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Sleep, 1000
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}  ; Clicks Edit button
Sleep, 1000
Send, {Tab}
Sleep, 300
Send, {Tab}
Sleep, 300
Send, {Tab}
Sleep, 300
Send, {Tab}
Sleep, 300
Send, {Tab}
Sleep, 1000

Send, %MyArray2%  ; Reassigned worklist
Sleep, 500
Send, {Tab}
Sleep, 500

Send, {Home}
Sleep, 500

Sleep, 2000
Clipboard =


sleep, 1000
Loop, {
;mouseclickdrag, left, 42, 311, 146, 311 ; HOME VERSION
mouseclickdrag, left, 41, 291, 145, 291
Send, ^c
clipwait, 
If clipboard contains Account View

{
;MsgBox, Account View page loaded.
Break
}
}

Sleep, 100
Send, {Tab}
Sleep, 200
Send, {Enter}  ; Clicks Save Button


Sleep, 2000
Clipboard =


sleep, 1000
Loop, {
;mouseclickdrag, left, 42, 313, 146, 313 ; HOME VERSION
mouseclickdrag, left, 41, 291, 145, 291
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
Send, {Tab}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Enter}  ; Clicks Cancel Button


Sleep, 2000
Clipboard =


sleep, 1000
Loop, {
;mouseclickdrag, left, 30, 312, 159, 312	; HOME VERSION
mouseclickdrag, left, 30, 290, 159, 290
Send, ^c
clipwait, 
If clipboard contains Search Inventory

{
;MsgBox, Search Inventory page loaded.
Break
}
}

; Wait for Request/ Response Account screen to disappear
Loop,
	{
	ImageSearch, FX, YX, 335,310, 495,370, %Imagelocation%Request Response Image.bmp
	If ErrorLevel = 1
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


}

FormatTime, TimeEnd,, Time
SoundPlaywe incorrectly calculated the allowed amount                                                                                                                                                                                                                                                                                                                                                  we incorrectly calculated the allowed amount                     