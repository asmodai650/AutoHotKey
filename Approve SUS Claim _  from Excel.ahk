MsgBox, This will Pass QA suspense claim #'s from a spreadsheet. Select the 1st claim in Excel to be resolved. Open AHK Window Spy.
InputBox,a,, What comment do you want in the Suspense notes?,,

c=0
inputbox,c,,How many claims?,,200,175
loop, %c%

{

WinActivate, Microsoft Excel -
WinWaitActive, Microsoft Excel - 
Send, {CtrlDown}c{CtrlUp}

WinActivate, Safari
WinWaitActive, Safari
Sleep, 500
Click, 103,157

WinWaitActive, Search
sleep, 100
Send, {CtrlDown}v{CtrlUp}
Sleep, 100
Send, {Tab 5}{Down 0}
Sleep, 100
Send, {Enter}

WinWaitActive, Patient Claim
Sleep, 500
Click, 390,65

WinWaitActive, Claim Request
Sleep, 500
Click, 334,135
Sleep, 500
Send, {Down 10}
Sleep, 1000
Send, {Enter}

WinWaitActive, Suspense Request
Sleep, 500
Click Right, 335,169
Sleep, 100
Click, 393,203

WinWaitActive, Suspense Request Tracking
Sleep, 500
Send, %a%
Sleep, 500
Click, 49,65

WinWaitActive, Close Request
Send, {Enter}
Sleep, 100

WinWaitClose, Suspense Request Tracking
Sleep, 100

WinClose, Suspense Request
WinWaitClose, Suspense Request
Sleep, 100

WinClose, Claim Request
WinWaitClose, Claim Request
Sleep, 100

WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 100

WinClose, Patient Claim
WinWaitClose, Patient Claim
Sleep, 100

WinActivate, Microsoft Excel -
WinWaitActive, Microsoft Excel - 
Send, {Down}

}

MsgBox, %c% claims resolved and approved.

