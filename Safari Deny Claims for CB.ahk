#SingleInstance force
SetTitleMatchMode, 2
InputBox,a,, Who is the requestor? (Enter the 1st name as it appears in Safari)
InputBox,b,, What is the 1st number for the denial reason?`nCB REASONS`n1 COB >> 1 COB - Commercial`n1 COB >> 2 COB - Medicare`n1 COB >> 3 Third Party Liability`n2 Contractual >> 1 Refund Written Incorrectly`n2 Contractual >> 2 Patient Refund - No refund due Payer`n2 Contractual >> 3 Refund Written According to Incorrect Information`n2 Contractual >> 4 Retroactive Contract`n3 Duplicate Issue >> 1 Duplicate Refund - Must cite origininal claim`n4 Other >> 1 Denied - PPN Claim`n5 Provider/Patient Account >> 1 Payment Posted to Incorrect Account`n5 Provider/Patient Account >> 2 Increase in total charges`n5 Provider/Patient Account >> 3 Documentation Not Available From Provider`n6 Pursued Prior >> 1 Payer Pursued Prior`n6 Pursued Prior >> 2 LOB Paid Prior`n6 Pursued Prior >> 3 Retracted or Paid Prior`n6 Pursued Prior >> 4 Vendor Pursued Prior`n6 Pursued Prior >> 5 Vendor Paid Prior`n7 Restriction >> 1 Payer Restriction (Must explain restriction violation)`n7 Restriction >> 2 Restricted Pay Type`n8 Uncollectible >> 1 Settlement Account`n8 Uncollectible >> 2 Payer cannot verify direct pay`n8 Uncollectible >> 3 Refund correct - Payer is not a client`n8 Uncollectible >> 4 Uncollectible Provider,,500,550

InputBox,c,, What is the 2nd number for the denial reason?`N1 COB >> 1 COB - Commercial`n1 COB >> 2 COB - Medicare`n1 COB >> 3 Third Party Liability`n2 Contractual >> 1 Refund Written Incorrectly`n2 Contractual >> 2 Patient Refund - No refund due Payer`n2 Contractual >> 3 Refund Written According to Incorrect Information`n2 Contractual >> 4 Retroactive Contract`n3 Duplicate Issue >> 1 Duplicate Refund - Must cite origininal claim`n4 Other >> 1 Denied - PPN Claim`n5 Provider/Patient Account >> 1 Payment Posted to Incorrect Account`n5 Provider/Patient Account >> 2 Increase in total charges`n5 Provider/Patient Account >> 3 Documentation Not Available From Provider`n6 Pursued Prior >> 1 Payer Pursued Prior`n6 Pursued Prior >> 2 LOB Paid Prior`n6 Pursued Prior >> 3 Retracted or Paid Prior`n6 Pursued Prior >> 4 Vendor Pursued Prior`n6 Pursued Prior >> 5 Vendor Paid Prior`n7 Restriction >> 1 Payer Restriction (Must explain restriction violation)`n7 Restriction >> 2 Restricted Pay Type`n8 Uncollectible >> 1 Settlement Account`n8 Uncollectible >> 2 Payer cannot verify direct pay`n8 Uncollectible >> 3 Refund correct - Payer is not a client`n8 Uncollectible >> 4 Uncollectible Provider,,500,550
InputBox,d,, How many claims?
MsgBox, Select the first claim on the spreadsheet to be moved to DREQ.

#Persistent
{
IfWinExist, Provider State
Send, {Enter}

IfWinExist, Safari Application
Send, !s

IfWinExist, Save Changes
Send, {Enter}
}

Loop, %d%
{
WinWait, Microsoft Excel
IfWinNotActive, Microsoft Excel
WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 100
Clipboard =
Sleep, 100
Send {HOME}
Sleep 300
Send, ^c
Sleep, 100
ClipWait
Sleep, 100
ClaimID = %Clipboard%
Sleep, 100
Send, {RIGHT}
Sleep, 100
Clipboard =
Sleep, 100
Send, ^c
Sleep, 100
ClipWait
Sleep, 100
Note = %Clipboard%
Sleep, 100
Send, {RIGHT}
Sleep, 100
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 100
Send, ^f
Sleep, 100
WinWait, Search
IfWinNotActive, Search
WinActivate, Search
WinWaitActive, Search
Sleep, 100
Send, %ClaimID%
Sleep, 100
Send, {ENTER}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {END}
Sleep, 100
Send, {ENTER}
Sleep, 1000
WinWait, Patient Claim
IfWinNotActive, Patient Claim
WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 100
ImageSearch, FoundX, FoundY, 1, 550, 350, 610, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\Invoiced.PNG
	If Errorlevel = 0
	{
	WinActivate, Patient Claim
	WinClose, Patient Claim
	WinWaitClose, Patient Claim
	Sleep, 100
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	Sleep, 100
	Send, Claim Invoiced, cannot deny
	Sleep, 100
	Send, {Down}{Left 2}
	Continue
	}
ImageSearch, FoundX, FoundY, 1, 550, 350, 610, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\RTI.PNG
	If Errorlevel = 0
	{
	WinActivate, Patient Claim
	WinClose, Patient Claim
	WinWaitClose, Patient Claim
	Sleep, 100
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	Sleep, 100
	Send, Claim is in RTI
	Sleep, 100
	Send, {Down}{Left 2}
	Continue
	}
ImageSearch, FoundX, FoundY, 1, 550, 350, 610, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\Provider Denied.PNG
	If Errorlevel = 0
	{
	WinActivate, Patient Claim
	;WinClose, Patient Claim
	;WinWaitClose, Patient Claim
	Sleep, 100
	WinActivate, Microsoft Excel
	WinWaitActive, Microsoft Excel
	Sleep, 100
	Send, Already denied
	Sleep, 100
	Send, {Down}{Left 2}
	Continue
	}
ImageSearch, FoundX, FoundY, 1, 550, 350, 610, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\Suspense.PNG
	If Errorlevel = 0
	{
	WinActivate, Patient Claim
	WinWaitActive, Patient Claim
	Sleep, 500
	Send, ^r  ; Request
	Sleep, 1000
	
	loop
	{
		IfWinExist, Save Changes
			{
			WinActivate, Save Changes
				Send, n
			WinWaitClose, Save Changes
			}
		IfWinExist, Claim Request
			break
	}
	
	WinActivate, Claim Requests
	WinWaitActive, Claim Requests
	Sleep, 500
	Send, {End}
	SLeep, 100
	Send, {Enter}
	Sleep, 100

	WinWaitActive, Suspense Request
	Sleep, 500
	Click, right, 306,167  ; right click on the sus entry
	Sleep, 500
	;Click, left, 372,180  ; select deny from drop down menu
	Send, {Down}{Enter}
	Sleep, 100

	WinWaitActive, Transaction Reasons
	Sleep, 100
	Send, {Tab 2}
	Sleep, 100
	Send, {Down %b%}{Enter} ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	Sleep, 100
	Send, {Down %c%}{Enter} ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	Sleep, 100

	WinWaitActive, Save Changes
	Sleep, 100
	Send, {Enter}
	WinWaitClose, Save Changes

	WinWait, Suspense Request Tracking - Claim
	IfWinNotActive, Suspense Request Tracking - Claim
	WinActivate, Suspense Request Tracking - Claim
	WinWaitActive, Suspense Request Tracking - Claim
	Sleep, 100
	SendInput, %Note%
	Sleep, 100
	MouseClick, Left, 65, 65, 1
	Sleep, 1000

;WinWaitClose, Suspense Request Tracking - Claim

		IfWinExist, Close Suspense
		{
		WinActivate, Close Suspense
		WinClose, Close Suspense
		WinWaitClose, Close Suspense
		WinActivate, Suspense Request Tracking
		WinClose, Suspense Request Tracking
		WinWaitClose, Suspense Request Tracking
		WinWaitActive, Suspense Request
		Sleep, 100
		Click, right, 306,167
		Sleep, 100
		Click, left, 363,231
		Sleep, 100
		WinWaitActive, Suspense Request Tracking
		Sleep, 100
		Send, Payment made on Claim. Resolving suspense.
		Sleep, 100
		Click, 63,67
		WinWaitActive, Close Request
		Sleep, 100
		Click, 273,132
		WinWaitClose, Suspense Request Tracking
		Sleep, 500
		WinClose, Suspense Request
		WinWaitclose, Suspense Request
		Sleep, 500
		WinClose, Claim Requests
		WinWaitClose, Claim Requests
		Sleep, 500
		;WinClose, Patient Claim
		;WinWaitClose, Patient Claim
		;Sleep, 500
		WinActivate, Microsoft Excel
		WinWaitActive, Microsoft Excel
		Sleep, 100
		Send, Money on claim, cannot deny
		Sleep, 100
		Send, {Down}{Left 2}
		Continue
		}

	
	WinWaitClose, Suspense Request Tracking - Claim

	WinClose, Suspense Request
	WinWaitClose, Suspense Request
	
	goto, Selectrequest 
	}

; Check if the claim is in DREQ
ImageSearch, FoundX, FoundY, 1, 550, 350, 610, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Aetna\Special Projects\Ben\Macros\Macro Images\DREQ.PNG
	If Errorlevel = 0
	{
	WinActivate, Patient Claim
	WinWaitActive, Patient Claim
	Sleep, 500
	Send, ^r  ; Request
	Sleep, 1000
loop
	{
		IfWinExist, Save Changes
			{
			WinActivate, Save Changes
				Send, n
			WinWaitClose, Save Changes
			}
		IfWinExist, Claim Request
			break
	}
	
	WinWaitActive, Claim Requests
	Sleep, 100

	goto, Selectrequest
	}

WinActivate, Patient Claim
WinWaitActive, Patient Claim
Sleep, 500
Send, ^r  ; Request
Sleep, 1000
loop
	{
		IfWinExist, Save Changes
			{
			WinActivate, Save Changes
				Send, n
			WinWaitClose, Save Changes
			}
		IfWinExist, Claim Request
			break
	}

WinActivate, Claim Requests
WinWaitActive, Claim Requests
Sleep, 500
Send, !ed  ; Select Deny Claim from Edit menu.

WinWaitActive, Claim Denial
Sleep, 100
SendInput, %a%
Sleep, 100
Send, {Enter}
Sleep, 500
Click, 50,170  ; Select a Reason

WinWaitActive, Transaction Reasons
Send, {Tab 2}
Sleep, 100
Send, {Down %b%} {Enter} ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Sleep, 100
Send, {Down %c%} {Enter} ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Sleep, 100
WinWaitClose, Transaction Reasons

WinActivate, Claim Denial
WinWaitActive, Claim Denial
Sleep, 500
Send, !fv  ; Save and Close
Sleep, 500
	IfWinExist, Provider State
	{
	WinActivate, Provider State
	Click, 284,99  ; OK
	WinWaitClose, Provider State
	}
WinWaitClose, Claim Denial

Selectrequest:

WinActivate, Claim Requests for Claim
WinWaitActive, Claim Requests for Claim
Sleep, 500
Send, {End}
Sleep, 500
Send, {Enter}

WinWaitActive, Claim Denial Request,,5
	If Errorlevel  ; if window above does not open after 5 seconds
	{
	WinClose, Suspense Request	
	WinClose, Claim Request
	WinWaitClose, Claim Request
	WinActivate, Patient Claim
	WinWaitActive, Patient Claim
	Sleep, 5000
	Send, ^r  ; Request
	WinWaitActive, Claim Requests for Claim
	Sleep, 500
	Send, {End}
	Sleep, 500
	Send, {Enter}
	}
		
WinWaitActive, Claim Denial Request
Sleep, 100
Send, %a%{Enter}
Sleep, 500
MouseClick, Left, 213, 325, 2  ; Double click on Needs Approval
;Click, 211,329

WinWaitActive, DM Account Manager Approval,,5
	
	If Errorlevel  ; if window above does not open after 5 seconds
	{
	ApprovalError:
	SplashTextOn,,,Safari is still processing your request - resetting request window
	Sleep, 2000
	SplashTextOff
	WinClose, Claim Denial Request
	Sleep, 1000
		IfWinExist, Save Changes
		{
		WinActivate, Save Changes
		Send, {Enter}
		WinWaitClose, Save Changes
		}
	WinWaitClose, Claim Denial Request
	Sleep, 100
	WinKill, Claim Requests for Claim
	WinWaitClose, Claim Requests for Claim
	Sleep, 100
	WinActivate, Patient Claim
	WinWaitActive, Patient Claim
	Sleep, 100
	Send, ^r
	WinActivate, Claim Requests
	WinWaitActive, Claim Requests
	Sleep, 500
	Send, {End}
	Sleep, 500
	Send, {Enter}
	WinWaitActive, Claim Denial Request
	Sleep, 500
	Click, 211,329  ; Double click on Needs Approval
	Click, 211,329
	WinWaitActive, DM Account Manager Approval,,5
	If Errorlevel
	GoTo, ApprovalError
	}
	Else
		
WinWaitActive, DM Account Manager Approval	
Sleep, 100
Send, A
Sleep, 100
Send, {Tab}
Sleep, 100
SendInput, %Note%
Send, {Tab}
Sleep, 500
Send, {Enter}
WinWaitClose, DM Account Manager Approval

WinActivate, Claim Denial Request
WinWaitActive, Claim Denial Request
Sleep, 500
Send, !fv  ; Save and Close
Sleep, 1000
	IfWinExist, Provider State
	{
	WinActivate, Provider State
	Click, 284,99  ; OK
	WinWaitClose, Provider State
	}

loop
	{
		IfWinExist, Claim Denial Request
			{
			winclose, Claim Denial Request
			sleep, 1000
			;WinWaitClose, Claim Denial Request
			}
		else
			break
	}


WinClose, Claim Requests
WinWaitClose, Claim Requests
Sleep, 100
;WinClose, Patient Claim
Send, ^s
Sleep, 100

loop
	{
		IfWinExist, Save Changes
			{
			WinActivate, Save Changes
				Send, n
			WinWaitClose, Save Changes
			}
		IfWinActive, Patient Claim
			break
	}

WinActivate, Microsoft Excel
WinWaitActive, Microsoft Excel
Sleep, 100
Send, Denied
Sleep, 100
Send, {Down}{Left 2}

updatetext =  %A_ComputerName%|%A_UserName%|%A_Now%|%A_ScriptName%

loop 3
{
FileAppend, %updatetext%, \\Aim.aimhealth.com\client_services$\Center of Excellence (COE)\Project Management\Macros\Macro Tracker\%A_ComputerName%.%A_UserName%.txt
                if errorlevel
                {
                                sleep, 1000
                                continue
                }
else
                break
}

}

Esc::ExitApp
Pause::Pause