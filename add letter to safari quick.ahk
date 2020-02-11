msgbox, This macro adds an image to safari.`n`nIn this script the "path" to the image must be the same as the Safari claim number. `nPlease complete the first image manualy to "set" the correct folder.

;#1::
loop, 
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}

StringReplace, clipboard, clipboard, `r`n, , All
claim= %Clipboard%
if claim =
	break
WinWait, Patient Claim, 
IfWinNotActive, Patient Claim, , WinActivate, Patient Claim, 
WinWaitActive, Patient Claim
send, ^f

WinWaitActive, Search, 
Sleep, 100
Send, %claim%{ENTER}
Sleep, 100
Send, {tab}{Enter}

WinWaitActive, Patient Claim
Sleep, 100
Send, {F3} ; images
Sleep, 100


WinWaitActive, Document Image, 
sleep, 3000
;MsgBox, 4, Does the letter already exists?, Add Letter?
;IfMsgBox No
;	{
;	WinClose
 ;   goto, result
	;}
WinWaitActive, Document Image, 
Sleep, 200
send, !f
sleep, 200
send, i

WinWaitActive, Open, 
Sleep, 200
send, %claim%.tif{enter}

WinWaitActive, Enter Image Name, 
Sleep, 100
Send, {Tab 2}
;MsgBox, correct?
Sleep, 100
Send, {Enter}
Sleep, 300

WinWaitActive, Document Image, 
Sleep, 200
;*******send, c
sleep, 200
;send, ^a
;send, !e
;sleep, 200
;send, {down 2}
;sleep, 200,
;send, {enter}
;MsgBox, correct?
WinClose
WinWaitActive, Patient Claim
Sleep, 300

IfWinExist, Safari - Application Error Report
	WinClose
WinWaitActive, Patient Claim
sleep, 300
Clipboard=
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep, 100
Send, {right}added{down}{left}
/*
continue
result:
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep, 100
Send, {right}letter already exsisted{down}{left}
*/
}
MsgBox, done
return
Pause::Pause