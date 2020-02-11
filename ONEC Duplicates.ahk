#SingleInstance Force

MsgBox, In Excel, click on the Claim Number in column A. Open Safari Corp. and go to the Data Mining Duplicate Claims window. Click in the "Claim ID" search field.

msgBox, Click OK to start this macro.

Loop
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Clipboard = 
ClaimID =

;Send, {ShiftDown}{Space}{ShiftUp}
Sleep, 500
Send, ^c
ClipWait, 1
StringReplace, Clipboard, Clipboard, `r`n, , All
StringSplit, Cell, Clipboard, %A_Tab%

ClaimID = %Cell1%

	if clipboard =
	{
	break
	}

WinWait, Safari
IfWinNotActive, Safari
WinActivate, Safari
WinWaitActive, Safari

Send, ^f
Sleep, 250
Send, Rule Desc
Sleep, 250
;Send, {Tab 5}
;Sleep, 250
;Send, {Enter}
;Sleep, 250
Send, {Esc}
Sleep, 250

Send, +{Tab 2}
Sleep, 250
Send, %ClaimID%
Sleep, 250
Send, {Tab}
Sleep, 250
Send, {Enter}
Sleep, 1000

Send, ^{a}
Sleep, 100
Send, ^{c}
Sleep, 250
If clipboard contains No rows returned
{
	;msgbox, FOUND!
	Comment = No rows returned
	Goto EXCEL
}

Sleep, 500
Send, ^f
Sleep, 250
Send, %ClaimID%
Sleep, 250
Send, {Esc}
Sleep, 250
Send, +{Tab}
Sleep, 250
Send, {s}
Sleep, 250

Send, +{Tab 14}
Sleep, 250
Send, {s}
Sleep, 250
Send, {Enter}
;add 1000
Sleep, 4000

Send, ^{a}
Sleep, 100
Send, ^{c}
Sleep, 250
If clipboard contains Process
{
	;msgbox, FOUND!
	Sleep, 6000
	Send, ^{a}
	Sleep, 100
	Send, ^{c}
	Sleep, 250
	If clipboard contains Process
		{
			Send, ^f
			Sleep, 250
			Send, Claim ID
			Sleep, 250
			Send, {Esc}
			Sleep, 250
			Send, {Tab 2}
			Sleep, 10000
			Comment = Possible Failure. Rerun Claim.
			GoTo EXCEL
		}
}

Send, ^{a}
Sleep, 100
Send, ^{c}
Sleep, 250
If clipboard contains Please correct the following
{
	;msgbox, FOUND!
	Comment = Deadlock. Rerun claim.
	Goto EXCEL
}

Comment = Done
Goto EXCEL

EXCEL:
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 250
Send, {tab}
Sleep, 100
Sendraw, %Comment%
Sleep, 100
Send, {Down}
Sleep, 100
Send, {Home}
Sleep, 100

ClaimID = 
Clipboard = 

}

MsgBox, Done!!

ExitApp

Pause::Pause
;Esc::Pause