#SingleInstance Force

MsgBox, Column A = Old Territory, Column B = New Territory, Column C = COB/Prov Code.

Loop
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500

OldTerritory =
NewTerritory =
CBOCode =

Send, {ShiftDown}{Space}{ShiftUp}
Sleep, 500
Send, ^c
ClipWait, 1
StringReplace, Clipboard, Clipboard, `r`n, , All
StringSplit, Cell, Clipboard, %A_Tab%

OldTerritory = %Cell1%
NewTerritory = %Cell2%
CBOCode = %Cell3%

	if OldTerritory =
	{
	break
	}

WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 500
Send, %CBOCode%
Sleep, 500
Send, {Enter}
Sleep, 500
Send, {Tab}
Sleep, 500
Send, {Enter}
Sleep, 500
WinWait, Provider Master File
IfWinNotActive, Provider Master File
WinActivate, Provider Master File
WinWaitActive, Provider Master File
Sleep, 500
Send, ^t
Sleep, 500
WinWait, Assign a Provider to a Region
IfWinNotActive, Assign a Provider to a Region
WinActivate, Assign a Provider to a Region
WinWaitActive, Assign a Provider to a Region
Sleep, 500
Send, {Tab 4}
Sleep, 500
send, {enter}
Sleep, 500
Send, {Enter}
Sleep, 500
Send, {Tab 3}
Sleep, 500


;Send, {down 22}
SendInput, %NewTerritory%
Sleep, 1000
ControlFocus, ThunderRT6CommandButton3, Assign a Provider to a Region
;ControlClick, ThunderRT6CommandButton3, Assign a Provider to a Region, , Left, 1
Sleep, 750
Send, {Enter}
Sleep, 750
Send, {Enter}
Sleep, 750
Send, !c
Sleep, 750
Send, ^f
Sleep, 500


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
Send, {Down}
Sleep, 500
Send, {Home}
Sleep, 500
}

MsgBox, Done!!

ExitApp

Esc::Pause