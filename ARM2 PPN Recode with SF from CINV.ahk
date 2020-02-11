;WinActivate PPN recode
;WinWaitActive PPN recode
#SingleInstance Force

Beginning:
WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, 

Status =

ClipBoard =
Sleep 300
Send {HOME}
Send {SHIFTDOWN}{RIGHT 69}{SHIFTUP}
Sleep 300
Send ^c
ClipWait

Send {HOME}

StringSplit, cell, clipboard, %A_Tab%,
FacsAcct# := cell1
StringReplace, cell1, cell1, `r`n, , All
Client := cell2
SafariClaim# := cell3
RecodeContract := cell4

StringReplace, cell54, cell54, `#, {`#}, All


StringReplace, cell20, cell20, `r`n, , All
If cell20 = 
  cell20 = .01

If cell1 =
{
ClipBoard =
Send ^{HOME}
Sleep 100
Send ^s
Sleep 2000
msgbox The ARM2 CINV recodes have been completed.
;Send !fde
;SetTitleMatchMode 2
;WinWaitActive Message
;SetTitleMatchMode 1
;Sleep 500
;Send vonna{ENTER}
;Sleep 500
;Send ^{ENTER}
ExitApp
}

StringReplace, cell14, cell14, `r`n, , All

WinActivate Patient Claim
WinWaitActive Patient Claim



PixelGetColor, SafariBlue, 226, 92

Send ^f
WinWaitNotActive Patient Claim

IfWinActive Save Changes
   Send !n

WinWaitActive Search

Send %SafariClaim#%{ENTER}{TAB}{DOWN}

PixelSearch, Px, Py, 337, 203, 350, 203, %SafariBlue%, 3, Fast
if ErrorLevel = 0
{
Sleep 300
WinClose,  Search
Sleep 300
WinWaitActive, Microsoft Excel - HMO Results for, 
Sleep 300
Send {RIGHT 7}Review- Multiple Refunds
Sleep 200
Send {DOWN}{HOME}
Sleep 300
GoTo Beginning
}


Send {ENTER}

NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 500

Loop
{
PixelSearch, Px, Py, 14,450,139,475, %SafariBlue%, 3, Fast
if ErrorLevel = 0
{
Break
}
if ErrorLevel = 1
{
Sleep 300
Click 290, 515
Sleep, 100
Send, +{TAB}
Sleep 300
}
}

;1.24.14 update
If cell1 = UND
  {
  Click 122, 380
  Sleep 200
  Send {TAB}.
  Sleep 200
  MouseClick, left,  685,  269
  MouseClick, left,  685,  269
  Sleep 300
  Send ARM2- Provider Unidentified{ENTER}
  Sleep 200
  Send !o
  Sleep 200
  Send ^s

  NoBreakOnTheseCursors=AppStarting, Wait 
  Loop 
  { 
    Sleep, 100 
    IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
      Break 
  }
  Sleep, 500

  WinWaitActive Patient Claim
  WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
  IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
  WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
  Send {RIGHT 7}Saved as Unidentified{ENTER}{HOME}
  GoTo Beginning
  }

;2.3.14 update
If cell2 = MR
  {
  Click 122, 380
  Sleep 200
  Send {TAB}.
  Sleep 200
  MouseClick, left,  685,  269
  MouseClick, left,  685,  269
  Sleep 300
  Send ARM2- Manual Review{ENTER}%cell1%{ENTER}
  Sleep 200
  Send !o
  Sleep 200
  Send ^s

  NoBreakOnTheseCursors=AppStarting, Wait 
  Loop 
  { 
    Sleep, 100 
    IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
      Break 
  }
  Sleep, 500

  WinWaitActive Patient Claim
  WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
  IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
  WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
  Send {RIGHT 7}Saved for manual review.{ENTER}{HOME}
  GoTo Beginning
  }


WinWaitActive Patient Claim
Sleep 500

/*
PixelSearch, Px, Py, 7,  236,  104,  278, 0x0000FF, 3, Fast
if ErrorLevel = 0
{
Send ^e

WinWaitNotActive Patient Claim
Sleep 500



IfWinActive Recode Policy
{
Send {ENTER}
WinWaitActive Patient Claim

ControlGetText, RefundAmt, ThunderRT6TextBox14, Patient Claim
Sleep 300
If RefundAmt = 1.00
{
ControlFocus, ThunderRT6TextBox15, Patient Claim
Send {TAB}.{TAB 2}%cell20%
Sleep 500
Send ^s
WinWaitNotActive Patient Claim

NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}

Sleep, 500

Sleep 500
GoTo Recode
}


WinWaitActive Patient Claim
Send ^e

WinWaitNotActive Patient Claim
Sleep 500

IfWinActive Recode Policy
{
Send {ENTER}
WinWaitActive Patient Claim
WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template,  
Send {RIGHT 7}Amt Diff{ENTER}{HOME}
GoTo Beginning
}
}

IfWinActive Invalid Recode
{
Send {ENTER}
WinWaitActive Patient Claim
GoTo InfoEntry
}

WinWaitActive Lookup Payor Contract

	Sleep 500
 	ControlGetText, PayorCode,  ThunderRT6TextBox2,  Lookup Payor Contract
	
	If PayorCode = 80000
	  GoTo Recode


	If PayorCode <> %RecodeContract%
	  GoTo Recode


	Recoded = Yes
	WinClose Lookup Payor Contract
 	WinWaitClose Lookup Payor Contract
	GoTo InfoEntry
}

*/
Send ^e

WinWaitNotActive Patient Claim
Sleep 500


IfWinActive Recode Policy
{
Send {ENTER}
WinWaitActive Patient Claim

ControlGetText, RefundAmt, ThunderRT6TextBox14, Patient Claim

;If RefundAmt = 1.00
{
;ControlFocus, ThunderRT6TextBox15, Patient Claim
;Send {TAB}.{TAB 2}%cell20%
;Sleep 500
;Send ^s
;WinWaitNotActive Patient Claim

;NoBreakOnTheseCursors=AppStarting, Wait 
;Loop 
;{ 
;  Sleep, 100 
;  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
;    Break 
;}
;Sleep, 500

Send {F6}
WinWaitNotActive Patient Claim
IfWinActive Save Changes
  Send !n
WinWaitActive Provider AR Summary for Claim
Sleep 200
Click 614, 333
Sleep 300
Send {HOME}{ENTER}
WinWaitActive Provider AR Detail
Sleep 200
Send !er
WinWaitActive Edit Amount
Sleep 200
mouseclick, right, 128, 55
Sleep 200
Send {DOWN 3}{ENTER}
ClipWait
WinClose Edit Amount
WinWaitActive Provider AR Detail
WinClose Provider AR Detail
WinWaitActive Provider AR Summary for Claim
WinClose Provider AR Summary for Claim
WinWaitActive Patient Claim
RefundMatch := ClipBoard
ClipBoard =

ControlFocus, ThunderRT6TextBox15, Patient Claim
Send {TAB}.{TAB 2}%RefundMatch% ;cell20
Sleep 500
Send ^s
WinWaitNotActive Patient Claim
WinWaitActive Patient Claim
Send ^e

WinWaitNotActive Patient Claim
Sleep 500

IfWinActive Recode Policy
{
Send {ENTER}
WinWaitActive Patient Claim
WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template,  
Send {RIGHT 7}Amt Diff{ENTER}{HOME}
GoTo Beginning
}

Sleep 500
GoTo Recode
}

WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
Send {RIGHT 7}Claim In Suspense{ENTER}{HOME}
GoTo Beginning
}

WinWaitActive Lookup Payor Contract

	Sleep 500
 	ControlGetText, PayorCode,  ThunderRT6TextBox7,  Lookup Payor Contract
 	ControlGetText, PayorName,  ThunderRT6TextBox6,  Lookup Payor Contract
	
	If PayorCode = 80000
	  GoTo Recode

	If PayorCode = 21291
	  GoTo Recode

	Recoded = Yes
	WinClose Lookup Payor Contract
 	WinWaitClose Lookup Payor Contract
	GoTo InfoEntry

Recode:
WinWaitActive Lookup Payor Contract

Send !fn
Sleep 300

ControlSend, ThunderRT6TextBox2, %RecodeContract%, Lookup Payor Contract

Sleep 300
Send {ENTER}{ENTER}


WinWaitNotActive Patient Claim


If cell3 contains PRS
{
WinWaitActive Report Code
Send %cell2%{DOWN}{TAB}{ENTER}
WinWaitClose Report Code
}

WinWait Recode Policy, ,2

IfWinExist Recode Policy
{
WinWaitActive Recode Policy

Sleep 300
Send !y

WinWaitClose Recode Policy
}



WinWaitActive Patient Claim


MouseClick, left,  281,  465
Sleep, 500


Loop
{
PixelSearch, Px, Py, 14,450,139,475, %SafariBlue%, 3, Fast
if ErrorLevel = 0
{
Break
}
if ErrorLevel = 1
{
Sleep 300
MouseClick, left,  221,  428
Sleep, 100
Send, {TAB 3}
Sleep 300
}
}

;Added 1.14.14 for required special fields
MouseClick, left,  571,  253
Sleep, 1000
MouseClick, left,  571,  253
Sleep, 100
MouseClick, left,  685,  269
MouseClick, left,  685,  269
Sleep, 300
Send, %cell31%{ENTER}%cell32%{ENTER}%cell33%{ENTER 5}1/1/11{ENTER 14}1{ENTER 9}1!o
Sleep 1000


ControlFocus, ThunderRT6TextBox15, Patient Claim
Send {TAB}.
;*************TESTING
;Send ^s
Send !fs

WinWaitNotActive Patient Claim


IfWinActive Missing Fields
{
	Send .
	WinClose Missing Fields
	WinWaitClose Missing Fields
}

IfWinActive Failed Validation
{
	Send {ENTER}
	WinWaitClose Failed Validation
	WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
	IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
	WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
	Send {RIGHT 7}Failed Validation{ENTER}{HOME}
	GoTo Beginning	
}

IfWinActive Save Changes
Send !y

;WinWaitActive Claim ID

Loop
{
IfWinExist Patient Claim
{
  WinActivate Patient Claim
  break
}
Sleep 500
}


Refresh:
WinWaitActive Patient Claim

IfWinNotActive Patient Claim
{
  Send {F5}
  GoTo Refresh
}
;SetKeyDelay 50
InfoEntry:
NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 500

;checks to see the status of claims
ImageSearch, FoundX, FoundY, 10,850,300,1030,C:\Users\%A_UserName%\Desktop\Recode Images\Status RTI.PNG
If ErrorLevel = 0
  Status = RTI

ImageSearch, FoundX, FoundY, 10,850,300,1030,C:\Users\%A_UserName%\Desktop\Recode Images\Status PQA.PNG
If ErrorLevel = 0
  Status = PQAd

ControlFocus, ThunderRT6TextBox9, Patient Claim
Sleep 300
Send +{TAB}
Sleep 200
Send %cell5%{TAB}%cell6%{TAB}%cell7%{TAB 3}%cell9%{TAB}

If cell10 = 
	Send %cell9%{TAB}
else
	Send %cell10%{TAB}

If cell11 = 
	Send 0{TAB}
else
	Send %cell11%{TAB}

Send {TAB}04{TAB 2}

If cell15 = 
	Send %cell7% %cell6%{TAB}
else
	Send %cell15% %cell16%{TAB}

If cell3 contains PRS ;1.8.14
  Send %cell17%{TAB}%cell18%{TAB}%cell19%{TAB}%cell37%{TAB 2}other{TAB}

Else
Send %cell17%{TAB}%cell18%{TAB}%cell19%{TAB 3}o{TAB}

ClipBoard := cell27
Sleep 300
Send {CTRLDOWN}{SHIFTDOWN}{HOME}{DELETE}{END}{DELETE}{SHIFTUP}{CTRLUP}
Send ^v
ClipBoard =
Sleep 500

;Enters UID
If cell3 contains PRS ;1.8.14
{
Send !vu
Sleep 200
Send ARM2`-%cell1%{ENTER}
Sleep 500
}

;Gets Provider Info & puts into special fields
If cell3 not contains PRS ;1.8.14
{
ProviderMasterFileClick:
ControlClick, ThunderRT6CommandButton3,  Patient Claim
Sleep 500
Send {ENTER}
WinWaitActive Provider Master File, ,2
IfWinNotActive Provider Master File
  GoTo ProviderMasterFileClick
ControlGetText, cell54,  ThunderRT6TextBox16,  Provider Master File
ControlGetText, cell55,  ThunderRT6TextBox7,  Provider Master File
WinClose Provider Master File
Sleep 500
IfWinExist Provider Master File
  Send ^{F4}
WinWaitActive Patient Claim
}

If cell3 contains PRS ;1.8.14
{
cell54 := cell63
cell55 := cell65
}


;Loop 31
;{
;    SF_Info := cell%A_Index%
;    Send %SF_Info%{ENTER}
;}

Click,  699,  266, 2

    LoopCount = 31
Loop 31
{
    SF_Info := cell%LoopCount%
	IF SF_Info =
		SF_Info = {DELETE}
    Send %SF_Info%{ENTER}
    LoopCount ++
}

If Cell18 <>
    ControlSend, ThunderRT6TextBox17, %Cell18%{TAB}, Patient Claim

If cell3 not contains PRS
{
If cell2 contains ASRC
{

;LINE ITEM ENTRY
ControlGetText, value2, ThunderRT6TextBox14, Patient Claim
StringReplace, value2, value2, `,,, ALL 
Sleep 100
value2 := RegExReplace(value2, "[`r`n`t]+$") 
Sleep 100
MouseClick, left,  376,  385
Sleep, 100
WinWait, Line Item Details, 
IfWinNotActive, Line Item Details, , WinActivate, Line Item Details, 
WinWaitActive, Line Item Details, 
Sleep, 500

MouseClick, left,  809,  260
Sleep, 1000


MouseClick, left,  55,  133
Sleep, 300
Send, %ID2%
Sleep 300
Send {TAB}00{TAB}
Sleep 300
Send %Value2%{TAB}
Sleep 300
Send 1111111
Sleep 300
WinClose Line Item Details
WinWaitClose Line Item Details
}
}

;Send ^s



;WinWaitActive Claim ID
WinWaitActive Patient Claim

If cell14 =
	GoTo PQA

If cell14 = 0
	GoTo PQA

If (cell14 <> "PPN Submitted to Provider- Awaiting Approvals" and cell14 <> "Pre-Collect")
{
Send ^d
WinWaitNotActive Patient claim

IfWinActive Save Changes
Send !y

WinWaitActive Patient Claim Detail
Sleep 300
ControlClick, edit8, Patient Claim Detail
Send {TAB}%cell14%{TAB}
Send !fv
WinWaitClose Patient Claim Detail
Sleep 300
WinWaitActive Patient Claim
}
;SetKeyDelay 10

PQA:
Sleep 500
;change payment date
Click  419,  268, 2
WinActivate, Provider System Transactions, 
WinWaitActive, Provider System Transactions, 
Sleep, 300
Send, {TAB}%cell9%
Sleep 200
Send !k
WinWaitClose Provider System Transactions
Sleep 1000


Sleep, 500

;If (Status <> "PQAd" and Status <> "RTI")
;If Status <> PQAd
Click  229,  244
Sleep 1000

;If Status = RTI
;  Send ^s

;If Status = PQAd
;  Send ^s

;If Status = Suspense
;  Send ^s

Sleep 1000

NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 1000


IfWinExist Failed Validation
{
Send {ENTER}
WinWaitClose Failed Validation
Sleep 300
Click  419,  268, 2
WinActivate, Provider System Transactions, 
WinWaitActive, Provider System Transactions, 
Sleep, 300
Send, {TAB}%cell9%
Sleep 200
Send !k
WinWaitClose Provider System Transactions
Send ^s
Sleep 1000
}

NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 500

IfWinActive Claim ID
  Send {F5}

;WinWaitActive Claim ID

Loop
{
IfWinExist Patient Claim
{
  WinActivate Patient Claim
  break
}
Sleep 500
}

WinWaitActive Patient Claim
WinWait, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
IfWinNotActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, , WinActivate, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
WinWaitActive, Microsoft Excel - AAA ARM2 Month End Live Report Template, 
Send {RIGHT 7}PQA'd{ENTER}{HOME}
GoTo Beginning

|::ExitApp

pause::pause