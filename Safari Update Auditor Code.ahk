#SingleInstance Force

MsgBox, "Column A = ClaimID, Column B = NewAuditor. SAFARI must be full screen."

Loop
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500

ClaimID =
NewAuditor =

Send, {ShiftDown}{Space}{ShiftUp}
Sleep, 500
Send, ^c
ClipWait, 1
StringReplace, Clipboard, Clipboard, `r`n, , All
StringSplit, Cell, Clipboard, %A_Tab%

ClaimID = %Cell1%
NewAuditor = %Cell2%

	if ClaimID =
	{
	break
	}

WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 500
Send, %ClaimID%
Sleep, 500
Send, {Enter}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}
Sleep, 500

WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 1000
MouseMove, 776, 123
Sleep, 250
MouseClick
Sleep, 500

WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 500
SendRaw, %NewAuditor%
Sleep, 500
Send, {Enter}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}
Sleep, 500

WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 1000
Send, !f
Sleep, 500
Send, s
Sleep, 1000
Send, ^f
Sleep, 500


;Sleep, 500
;ControlFocus, ThunderRT6CommandButton3, Change Auditor
;ControlClick, ThunderRT6CommandButton3, Assign a Provider to a Region, , Left, 1
;Sleep, 500


WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 500

WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500
Send, {Tab 2}
Sleep, 250
Send, x
Sleep, 500
Send, {Down}
Sleep, 500
Send, {Home}
Sleep, 500
}

MsgBox, Done!!

ExitApp

Esc::Pause