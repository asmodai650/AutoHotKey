; this scripts copies check images from the unidentified claim to the claim the money is posted to.

inputbox,c,,This macro Copies images from Unidentifieds to the claim the funds were posted to.`n`nPlease open the Unidentified Search windos in Safari.`n`n`n 		 How many claims?,,500,200

Gui, Add, Picture, x6 y40 w340 h260 , \\Aim\client_services$\MACROS- Chris Taylor\Matt\Images\Move unidentified image Setup example.bmp
Gui, Add, Button, x36 y320 w100 h30 , Continue
Gui, Add, Button, x176 y320 w100 h30 , Cancel
Gui, Add, Text, x26 y10 w310 h30 +Center, This is how the spreadsheet should look.
Gui, Add, Edit, x766 y890 w50 h80 , Edit
; Generated using SmartGUI Creator for SciTE
Gui, Show, w355 h355, Move Unidentified Image GUI
return

ButtonCancel:
ExitApp
ButtonContinue:
Gui, Cancel


loop, %c%
{
TrayTip,, %a_index% of %c%.,1
WinWait, Microsoft Excel, 
	IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
	WinWaitActive, Microsoft Excel, 
	Sleep, 100
	send, {ShiftDown}{Space}{ShiftUp}
	Sleep, 100
	Send, ^c{down}{up}
	ClipWait

StringReplace, clipboard, clipboard, `r`n, , All
StringSplit, cell, Clipboard, %a_tab%

key=%cell1%
note=%cell2%
	Sleep, 100

TrayTip,, %a_index% of %c%.`nKey %key%`nNote %note%,1

WinWait, Search, 
	IfWinNotActive, Search, , WinActivate, Search, 
	WinWaitActive, Search, 
	Sleep, 300
	Send, %key%{Enter}
	sleep, 100
	send, {tab}{enter}

WinWaitActive, Unidentified Provider Claim, 
	sleep, 300
	Send, !t
	sleep, 100
	Send,{UP  2}{ENTER}

WinWaitActive, Document Image, 
	sleep, 1000
	Send, !e
	sleep, 500
	send,{UP  2}
	sleep, 200
	send, {ENTER}

WinWaitActive, Find, 
	sleep, 300
	WinWaitActive, Find, 
	MouseClick, left, 161, 45,
	Sleep, 100
	Send, c{TAB}c{TAB}%note%{TAB}{SPACE}{ENTER}

WinWaitActive, Copy/Move Image, 
;	MouseClick, left,  53,  21
	Sleep, 100
	Send, {ENTER}

WinWaitActive, Document Image, 
	Sleep, 100
	WinClose

WinWaitActive, Unidentified Provider Claim, 
	Send, ^f

WinWaitActive, Search, 

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
sleep, 100
Send, done{DOWN}
}
MsgBox, Done, %c% complete.`
return
Pause::Pause
esc::ExitApp
