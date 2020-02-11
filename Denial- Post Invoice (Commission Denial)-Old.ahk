IfWinNotExist, Post-Invoice Denial Reasons,
{

run \\aim\client_services$\Chris Taylor\All Public Macros\General Use Macros\Commission Denial- Post Inv\Post-Invoice Denial Reasons.xlsx
WinWait, Password, 
IfWinNotActive, Password, , WinActivate, Password, 
WinWaitActive, Password, 
Sleep, 500
Send, {ENTER}
WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
Sleep 500
}

MsgBox, 
(
Description: 	This macro enters commission denials for Post-Invoice claims with no money on the claim.

Requirements: 	This macro requires you to use the Post-Invoice Denial Reasons spreadsheet. It requires the AIM Claim ID in
		Column A, the Commission Denial Amount in Column B, the Denial Reason in Column C and the Denial Comments in
		Column D.The reasons MUST match the reasons shown in Column H of the spreadsheet. 
	
		DO NOT paste anything in column E.

Starting Inst:	Make sure that Safari is running but all claims are closed. Safari should be maximized on one
		screen and Post-Invoice Denial Reasons.xlsx should be maximized on your other screen. Make sure that you
		have the cell selected for the 1st claim you need to deny from Column A.
)


;InputBox, UserInput1, Requestor's Name, Please enter the name of the Requestor (Only as much as is need to select the correct name on the denial.), , 300, 150
;if ErrorLevel
;{
;    MsgBox, CANCEL was pressed.
;	Return
;}
;else



;Loop
;{
;IfWinExist, Patient Claim,
;WinWaitActive, Patient Claim, , WinClose
;}

;Loop
;{
;IfWinExist, Claim Requests for Claim ID:  , , WinClose
;}


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
Send {RIGHT}

Send, ^c
ClipWait
StringReplace, clipboard, clipboard, `r`n, , All
DenialReason := ClipBoard
ClipBoard =
Send {RIGHT}

Send, ^c
ClipWait
StringReplace, clipboard, clipboard, `r`n, , All
DenialComments := ClipBoard
ClipBoard =
Send {RIGHT }



WinWait, Patient Claim , 
IfWinNotActive, Patient Claim , , WinActivate, Patient Claim , 
WinWaitActive, Patient Claim , 
WinMaximize
Sleep, 100
Send, {CTRLDOWN}f{CTRLUP}
WinWait, Search, 
IfWinNotActive, Search, , WinActivate, Search, 
WinWaitActive, Search, 
Send, %ClaimID%{ENTER}
Sleep, 100
Send, {TAB}{ENTER}
NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}

Sleep, 1000
Send, {CTRLDOWN}r{CTRLUP}
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}

Sleep, 1000
WinWait, Claim Requests for Claim ID:  , 
IfWinNotActive, Claim Requests for Claim ID:  , , WinActivate, Claim Requests for Claim ID:  , 
WinWaitActive, Claim Requests for Claim ID:  , 
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
Send, {RIGHT 5}
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
Send, {RIGHT 4}
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

			If (1) = clipboard 
			{
			Send, {DOWN}{RIGHT}{DOWN}
			}
			If (2) = clipboard
			{
			Send, {DOWN}{RIGHT}{DOWN 2}
			}
			If (3) = clipboard
			{
			Send, {DOWN}{RIGHT}{DOWN 3}
			}
			If (4) = clipboard
			{
			Send, {DOWN 2}{RIGHT}{DOWN}
			}
			If (5) = clipboard
			{
			Send, {DOWN 2}{RIGHT}{DOWN 2}
			}
			If (6) = clipboard
			{
			Send, {DOWN 2}{RIGHT}{DOWN 3}
			}
			If (7) = clipboard
			{
			Send, {DOWN 2}{RIGHT}{DOWN 4}
			}
			If (8) = clipboard
			{
			Send, {DOWN 3}{RIGHT}{DOWN}
			}
			If (9) = clipboard
			{
			Send, {DOWN 4}{RIGHT}{DOWN}
			}
			If (10) = clipboard
			{
			Send, {DOWN 5}{RIGHT}{DOWN}
			}
			If (11) = clipboard
			{
			Send, {DOWN 5}{RIGHT}{DOWN 2}
			}
			If (12) = clipboard
			{
			Send, {DOWN 5}{RIGHT}{DOWN 3}
			}
			If (13) = clipboard
			{
			Send, {DOWN 6}{RIGHT}{DOWN}
			}
			If (14) = clipboard
			{
			Send, {DOWN 6}{RIGHT}{DOWN 2}
			}
			If (15) = clipboard
			{
			Send, {DOWN 6}{RIGHT}{DOWN 3}
			}
			If (16) = clipboard
			{
			Send, {DOWN 6}{RIGHT}{DOWN 4}
			}
			If (17) = clipboard
			{
			Send, {DOWN 6}{RIGHT}{DOWN 5}
			}
			If (18) = clipboard
			{
			Send, {DOWN 7}{RIGHT}{DOWN}
			}
			If (19) = clipboard
			{
			Send, {DOWN 7}{RIGHT}{DOWN 2}
			}
			If (20) = clipboard
			{
			Send, {DOWN 8}{RIGHT}{DOWN}
			}
			If (21) = clipboard
			{
			Send, {DOWN 8}{RIGHT}{DOWN 2}
			}
			If (22) = clipboard
			{
			Send, {DOWN 8}{RIGHT}{DOWN 3}
			}
			If (23) = clipboard
			{
			Send, {DOWN 8}{RIGHT}{DOWN 4}
			}
			If (24) = clipboard
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
Send ClipBoard =


Sleep, 500
WinWaitActive, Notes, 
IfWinExist, Notes, , WinClose
Sleep, 100


WinWait, Save Changes, 
IfWinNotActive, Save Changes, , WinActivate, Save Changes, 
WinWaitActive, Save Changes, 

Sleep, 100
Send, {ENTER}
Sleep, 100
WinWait, Commission Denial, 
IfWinNotActive, Commission Denial, , WinActivate, Commission Denial, 
WinWaitActive, Commission Denial, 
Sleep, 100
}

Send, {ALTDOWN}fv{ALTUP}

;Send, {CTRLDOWN}{F4}{CTRLUP}




 

WinWaitActive, Claim Requests for Claim ID:  , 
IfWinExist, Claim Requests for Claim ID:  , , WinClose

WinWait, Microsoft Excel - Post-Invoice Denial Reasons, 
IfWinNotActive, Microsoft Excel - Post-Invoice Denial Reasons, , WinActivate, Microsoft Excel - Post-Invoice Denial Reasons, 
WinWaitActive, Microsoft Excel - Post-Invoice Denial Reasons, 
Sleep, 100
Send, {RIGHT}x{ENTER}{LEFT 5}


}

End:
run \\aim\client_services$\Chris Taylor\All Public Macros\Macro Log.xls
WinWait, Microsoft Excel - Macro Log.xls, 
IfWinNotActive, Microsoft Excel - Macro Log.xls, , WinActivate, Microsoft Excel - Macro Log.xls, 
WinWaitActive, Microsoft Excel - Macro Log.xls, 
Sleep 300
Send %A_AhkPath%{RIGHT}
Send %WinTitle%
Sleep 300
Send %A_UserName%{RIGHT}
Sleep 300
Send, %A_MM%/%A_DD%/%A_YYYY%{RIGHT}
Sleep 300
FormatTime, TimeString,, hh:mm:ss tt
Send  %TimeString%{RIGHT}
Sleep 300
Send %LoopCount%{DOWN}{HOME}
Sleep 300
Send ^s
Sleep 3000
Send ^{F4}
Sleep 300
ExitApp



]::
Reload
Pause::
Pause