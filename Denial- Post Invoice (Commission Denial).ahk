IfWinNotExist, Post-Invoice Denial Reasons,
{

;run \\aim.aimhealth.com\client_services$\Chris Taylor\All Public Macros\General Use Macros\Commission Denial- Post Inv\Post-Invoice Denial Reasons.xlsx
;WinWait, Password, 
IfWinNotActive, Password, , WinActivate, Password, 
;WinWaitActive, Password, 
;Sleep, 500
;Send, {ENTER}
;WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
;IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
;WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
;Sleep 500
}

MsgBox, 
(
Description: 	This macro enters commission denials for Post-Invoice claims with no money on the claim.

Requirements: 	This macro requires you to use the Post-Invoice Denial Reasons spreadsheet. It requires the AIM Claim ID in
		Column A, the Commission Denial Amount in Column B, the Denial Reason in Column C and the Denial Comments in
		Column D. If you do not want to enter a note, just leave column D blank. The reasons MUST match the reasons 
		shown in Column H of the spreadsheet. 
	
		DO NOT paste anything in column E.

Starting Inst:	Make sure that Safari is running only on claim is open. Safari should be maximized on one
		screen and Post-Invoice Denial Reasons.xlsx should be maximized on your other screen. Make sure that you
		have the cell selected for the 1st claim you need to deny from Column A.
)





WinWait, Safari, 
IfWinNotActive, Safari, , WinActivate, Safari, 
WinWaitActive, Safari, 
WinMaximize
;pause
Sleep, 100
Send, {ALTDOWN}i{ALTUP}{ENTER}
Sleep, 100
WinMaximize
Sleep, 100

Beginning:
Loop,
{
ClipBoard =
WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 

Sleep, 100
Send, ^{LEFT}
Sleep, 100
Send, ^c
ClipWait
ClaimID := ClipBoard
ClipBoard =
Send {RIGHT}

Send, ^c
ClipWait
DenialAmt := ClipBoard
ClipBoard =
Send {RIGHT 2}


Send, ^c
ClipWait
StringReplace, clipboard, clipboard, `r`n, , All
DenialComments := ClipBoard
ClipBoard =
Send {RIGHT}

Send, ^c
ClipWait
StringReplace, clipboard, clipboard, `r`n, , All
DenialReason := ClipBoard
ClipBoard =
Send {RIGHT}




WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
WinMaximize
Sleep, 100
Send, {CTRLDOWN}f{CTRLUP}

;Identify Claims w/ Multiple Refunds
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Sleep, 300
Send, %ClaimID%
Sleep 300
Send {ENTER}{TAB}
Sleep 300
Send {DOWN}
Sleep 200

PixelSearch, Px, Py, 337, 203, 350, 203, 0xC56A31, 3, Fast
if ErrorLevel = 0
{
Sleep 300
WinClose,  Search
Sleep 300
WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
Sleep 300
Send Review- Multiple Refunds
Sleep 200
Send {DOWN}{HOME}
Sleep 300
GoTo Beginning
}



Sleep, 100
Send, {ENTER}
NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}

Sleep, 1000

OpenRequest:
Send, {CTRLDOWN}r{CTRLUP}
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}

Sleep, 1000



WinWaitActive, Claim Requests for Claim ID:  , , 3
IfWinNotActive Claim Requests for Claim ID:
{
IfWinActive Provider State
  Send {ENTER}
GoTo OpenRequest
}
Sleep, 1000
Send, {ALTDOWN}ec{ALTUP}
NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 1000

;If statement so that if claim is already in research it will skip to the next claim.

IfWinNotExist, Commission Denial,
{
Sleep, 2000
}

IfWinNotExist, Commission Denial,
{
WinWaitActive, Claim Requests for Claim ID:  , 
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100
WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
Sleep, 100
Send, Claim Not Allowing Commission Denial
Sleep, 100
Send, {DOWN}{LEFT 5}
Continue
}



WinWait, Commission Denial, 
IfWinNotActive, Commission Denial, , WinActivate, Commission Denial, 
WinWaitActive, Commission Denial, 
Sleep, 100
Send, {TAB}
Sleep, 100
Send, %DenialAmt%
Sleep, 100
MouseClick, left,  402,  115
Sleep, 1000

IfWinExist, Save Record,
{
Sleep 300
WinWait, Save Record, 
IfWinNotActive, Save Record, , WinActivate, Save Record, 
WinWaitActive, Save Record, 
Sleep, 500
Send {ENTER}
Sleep, 500

IfWinExist, Save Record,
{
Pause
Sleep, 500
Send {ENTER}
Sleep, 1000
}

Send {DEL}
Sleep 500

WinWait, Commission Denial, 
IfWinNotActive, Commission Denial, , WinActivate, Commission Denial, 
WinWaitActive, Commission Denial, 
Sleep 500
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 500

WinWait, Save Changes, 
IfWinNotActive, Save Changes, , WinActivate, Save Changes, 
WinWaitActive, Save Changes, 
Sleep, 1000
Send, {TAB}{ENTER}
Sleep, 500
WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  , 
Sleep 500
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 500
WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
Sleep, 100
Send, Denial Amt Exceeds Commission Due
Sleep, 100
Send, {DOWN}{HOME}
Continue
}


;MouseClick, left,  222,  16
;Sleep, 100
;Send, %UserInput1%
;MouseClick, left,  45,  169
;Sleep, 100



			WinWait, Transaction Reasons, 
			IfWinNotActive, Transaction Reasons, , WinActivate, Transaction Reasons, 
			WinWaitActive, Transaction Reasons, 

			Sleep, 300
			Send, {TAB 2}

			If (1) = DenialReason 
			{
			Send, {DOWN}{RIGHT}{DOWN}
			}
			If (2) = DenialReason
			{
			Send, {DOWN}{RIGHT}{DOWN 2}
			}
			If (3) = DenialReason
			{
			Send, {DOWN}{RIGHT}{DOWN 3}
			}
			If (4) = DenialReason
			{
			Send, {DOWN 2}{RIGHT}{DOWN}
			}
			If (5) = DenialReason
			{
			Send, {DOWN 2}{RIGHT}{DOWN 2}
			}
			If (6) = DenialReason
			{
			Send, {DOWN 2}{RIGHT}{DOWN 3}
			}
			If (7) = DenialReason
			{
			Send, {DOWN 2}{RIGHT}{DOWN 4}
			}
			If (8) = DenialReason
			{
			Send, {DOWN 3}{RIGHT}{DOWN}
			}
			If (9) = DenialReason
			{
			Send, {DOWN 4}{RIGHT}{DOWN}
			}
			If (10) = DenialReason
			{
			Send, {DOWN 5}{RIGHT}{DOWN}
			}
			If (11) = DenialReason
			{
			Send, {DOWN 5}{RIGHT}{DOWN 2}
			}
			If (12) = DenialReason
			{
			Send, {DOWN 5}{RIGHT}{DOWN 3}
			}
			If (13) = DenialReason
			{
			Send, {DOWN 6}{RIGHT}{DOWN}
			}
			If (14) = DenialReason
			{
			Send, {DOWN 6}{RIGHT}{DOWN 2}
			}
			If (15) = DenialReason
			{
			Send, {DOWN 6}{RIGHT}{DOWN 3}
			}
			If (16) = DenialReason
			{
			Send, {DOWN 6}{RIGHT}{DOWN 4}
			}
			If (17) = DenialReason
			{
			Send, {DOWN 6}{RIGHT}{DOWN 5}
			}
			If (18) = DenialReason
			{
			Send, {DOWN 7}{RIGHT}{DOWN}
			}
			If (19) = DenialReason
			{
			Send, {DOWN 7}{RIGHT}{DOWN 2}
			}
			If (20) = DenialReason
			{
			Send, {DOWN 8}{RIGHT}{DOWN}
			}
			If (21) = DenialReason
			{
			Send, {DOWN 8}{RIGHT}{DOWN 2}
			}
			If (22) = DenialReason
			{
			Send, {DOWN 8}{RIGHT}{DOWN 3}
			}
			If (23) = DenialReason
			{
			Send, {DOWN 8}{RIGHT}{DOWN 4}
			}
			If (24) = DenialReason
			{
			Send, {DOWN 8}{RIGHT}{DOWN 5}
			}
			Send, {Enter}












WinWait, Commission Denial, 
IfWinNotActive, Commission Denial, , WinActivate, Commission Denial, 
WinWaitActive, Commission Denial, 

If DenialComments <> 
{
Sleep, 100
Send, {F2}
WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, 
Sleep, 100
Send, {ALTDOWN}n{ALTUP}
Sleep, 500
NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 500

WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, 
Sleep, 100

Send, {TAB}
Sleep, 100
ClipBoard := DenialComments
Send ^v

CaptureNote:
ClipBoard =
Send {CTRLDOWN}{SHIFTDOWN}{HOME}{CTRLUP}{SHIFTUP}
Sleep 100
Send ^c
ClipWait, ,2

If clipboard <> %DenialComments%
  GoTo CaptureNote


Sleep, 500
WinWaitActive, Notes, 
IfWinExist, Notes, , WinClose
Sleep, 100


WinWait, Save Changes, 
IfWinNotActive, Save Changes, , WinActivate, Save Changes, 
WinWaitActive, Save Changes, 
Sleep, 100
Send, {ENTER}
Sleep, 500


WinWait, Commission Denial, 
IfWinNotActive, Commission Denial, , WinActivate, Commission Denial, 
WinWaitActive, Commission Denial, 
Sleep, 100
}

;Send, {ALTDOWN}fv{ALTUP}

;Send, {CTRLDOWN}{F4}{CTRLUP}




WinClose, Commission Denial, 
IfWinExist, Save Changes,
	Send, {Enter}
WinWaitClose Commission Denial
WinActivate Claim Requests for Claim ID
WinWaitActive, Claim Requests for Claim ID:  , 
IfWinExist, Claim Requests for Claim ID:  , , WinClose

WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
Sleep, 100
Send, x{ENTER}{LEFT 5}


}



]::
Reload
Pause::
Pause