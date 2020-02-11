#SingleInstance Force

MsgBox, Safari Provider Master File window must be open to a new provider! Excel Sheet coulmns must be in this format:, 1 Code, 2 Tax ID, 3 Provider Name, 4 Mailing Address, 5 Zip Code, 6 Zip Code Extra Digits, 7 City, 8 State, 9 FacType, 10 Rtn Chck, 11 Macro Result


Loop
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 500

Clipboard =

Send, {ShiftDown}{Space}{ShiftUp}
Sleep, 500
Send, ^c
ClipWait, 1
StringReplace, Clipboard, Clipboard, `r`n, , All
StringSplit, Cell, Clipboard, %A_Tab%

Code = %Cell1%
TaxID = %Cell2%
ProvName = %Cell3%
Address = %Cell4%
Zip = %Cell5%
Zip2 = %Cell6%
City = %Cell7%
State = %Cell8%
FacType = %Cell9%
ReturnCheck = %Cell10%



	if Code =
	{
	break
	}

;WinWait, Search
;IfWinNotActive, Search
;WinActivate, Search
;WinWaitActive, Search
;Sleep, 500
;Send, %Code%
;Sleep, 500
;Send, {Enter}
;Sleep, 500
;Send, {Tab}
;Sleep, 500
;Send, {Enter}
;Sleep, 500



WinWait, Provider Master File
IfWinNotActive, Provider Master File
WinActivate, Provider Master File
WinWaitActive, Provider Master File
Sleep, 500
Send, !f
Sleep, 500
Send, n
Sleep, 500


Send, p
Sleep, 500
Send, {tab}
Sleep, 500

SendRaw, %Code%
Sleep, 500
Send, {tab}
Sleep, 500

SendRaw, %TaxID%
Sleep, 500
Send, {tab 2}
Sleep, 500

SendRaw, %ProvName%
Sleep, 500
Send, {tab}
Sleep, 500

SendRaw, %Address%
Sleep, 500
Send, {tab 2}
Sleep, 500

SendRaw, %Zip%
Sleep, 1000

SendRaw, %Zip2%
Sleep, 500
Send, {tab 3}
Sleep, 500

SendRaw, %City%
Sleep, 500
Send, {tab}
Sleep, 500

SendRaw, %State%
Sleep, 500
Send, {tab 7}
Sleep, 500

SendRaw, %FacType%
Sleep, 500
Send, {tab 6}
Sleep, 500

SendRaw, %ReturnCheck%
Sleep, 500
Send, {tab 2}
Sleep, 500

Sleep, 500
Send, !f
Sleep, 500
Send, s
Sleep, 500

;Sleep, 500
;Send, !f
;Sleep, 500
;Send, n
;Sleep, 1000


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