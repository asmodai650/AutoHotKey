#SingleInstance, Force



MsgBox, Open OZ and sort your tickets by subject. For each Humana TPL Pre-Approval ticket, copy the ticket number and the Requestor name and paste them into Columns M and L in Excel. When finished, select the cell in Column M for the OZ Ticket you want to start with. When ready, click "OK" in this box to begin.

CHECK:
Loop,
{
IfWinExist,  « RAM Helpdesk » ,
	{
		;msgbox, Window found
		WinActivate
		break
	}

IfWinNotExist,  « RAM Helpdesk » ,
	{
		MsgBox, Make sure that the RAM Helpdesk webpage is open!
		Sleep, 2000
		Goto, Check
	}
}
WinWaitActive,  « RAM Helpdesk » ,
Sleep, 1000
Send, {Browser_Refresh}
Sleep, 1000
Send, {Enter}
Sleep, 1000


loop
{

WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 1000
;ClipBoard =

Sleep, 1000
Send, {CTRLDOWN}c{CTRLUP}
ClipWait
StringReplace, ClipBoard, ClipBoard, `r`n, ,All
;msgbox, %clipboard%

If ClipBoard =
{
Msgbox Macro is complete.
ExitApp
}


Sleep, 1000
WinWait,  « RAM Helpdesk » , 

Loop,
{
IfWinExist,  « RAM Helpdesk » ,
	{
		;msgbox, Window found
		WinActivate
		break
	}

IfWinNotExist,  « RAM Helpdesk » ,
	{
		MsgBox, Make sure that the RAM Helpdesk webpage is open!
		Pause
	}
}


; IfWinNotActive, « RAM Helpdesk » Optum, Inc. , , WinActivate, « RAM Helpdesk » Optum, Inc. , 
; WinWaitActive, « RAM Helpdesk » , 
Sleep, 1000
Send, ^v
sleep, 200
Send, {tab}{enter}
Sleep, 2000

Send, ^a
Sleep, 1000
Send, ^c
ClipWait, 10, 1

Loop,
	{
		If clipboard contains Humana GB-TPL Pre-Approval
		{
			Sleep, 500
			mousemove, 750, 800
			click
			goto, FINDCLAIMINFORMATION
		}
		else
		{
			Sleep, 1000
			Send, ^f
			Sleep, 1000
			Send, Issue Details
			Sleep, 500
			Send, {Tab 2}
			Sleep, 500
			Send, +{Tab 6}
			Sleep, 200
			Send, {Enter}
			Sleep, 1000
			;Send, {Delete 8}
			;Sleep, 500
			comment = Double Check OZ Ticket
			WinWait, Microsoft Excel, 
			IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
			WinWaitActive, Microsoft Excel, 
			Sleep, 1000
			Clipboard = 
			Goto, EXCEL
		}
	}

FINDCLAIMINFORMATION:

Sleep, 1000

Send, ^f
sleep, 500
clipboard = 
Comment = Claim Information (required)
Send, %Comment%
Sleep, 1000

; patient name
Clipboard = 
Send, {Tab 2}
Sleep, 500
Sleep, 200
send, ^a
Sleep, 200
send, ^c
Sleep, 1000
ClipWait, 
Clip1 = %Clipboard%
;MsgBox, %Clip1%

;humana claim ID
Clipboard = 
Send, {Tab 13}
Sleep, 500
send, ^c
ClipWait,
Clip2 = %clipboard%
;MsgBox, %Clip2%

;member ID
Clipboard = 
Send, +{Tab}
Sleep, 500
send, ^c
ClipWait,
Clip3 = %clipboard%
;MsgBox, %Clip3%

;DOS
Clipboard = 
Send, +{Tab 9}
Sleep, 500
send, ^c
ClipWait,
Clip4 = %clipboard%
;MsgBox, %Clip4%

;Date of Loss
Clipboard = 
Send, {Tab 5}
Sleep, 500
send, ^c
ClipWait,
Clip5 = %clipboard%
;MsgBox, %Clip5%

;Provider Name
Clipboard = 
Send, +{Tab 4}
Sleep, 500
send, ^c
ClipWait,
Clip6 = %clipboard%
;MsgBox, %Clip6%

;TPL Carrier
Clipboard = 
Send, {Tab 5}
Sleep, 500
send, ^c
ClipWait,
Clip7 = %clipboard%
;MsgBox, %Clip7%

;Auto or Worker's Claim?
InputBox, Clip8, Claim Type, Please enter the claim type: AUTO or WORKER., , 640, 480
if ErrorLevel
	{
	MsgBox, CANCEL was pressed. Macro will now stop. Please Restart the macro!
		WinWaitActive,  « RAM Helpdesk » ,
		Sleep, 1000
		Send, {Browser_Back}
		Sleep, 1000
		Send, {Delete 8}
		Sleep, 500
		Send, {Backspace 8}
		Sleep, 1000
		Clipboard = 
		ExitApp
	}
%Clip8% = %clipboard%
;MsgBox, %Clip8%

;Auto Policy Holder's Name
Clipboard = 
send, {tab 2}
Sleep, 500
send, ^c
ClipWait,
Clip9 = %clipboard%
;MsgBox, %Clip9%

;Refund Amount
Clipboard = 
Send, +{Tab 6}
Sleep, 500
send, ^c
ClipWait,
Clip10 = %clipboard%
;MsgBox, %Clip10%

Sleep, 1000
Goto, PASTEDATA

PASTEDATA:
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 

Sleep, 100
Send, {home}
Sleep, 100
Send, %Clip1%
Sleep, 100
Send, {Tab}
Sleep, 100
Send, %Clip2%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip3%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip4%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip5%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip6%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip7%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip8%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip9%
Sleep, 100
send, {tab}
Sleep, 100
Send, %Clip10%
Sleep, 100

Goto, UPDATETICKET

UPDATETICKET:
;MsgBox, Data Entered

Sleep, 1000
WinWait,  « RAM Helpdesk » , 

Loop,
{
IfWinExist,  « RAM Helpdesk » ,
	{
		;msgbox, Window found
		WinActivate
		break
	}

IfWinNotExist,  « RAM Helpdesk » ,
	{
		MsgBox, Cannot Find Window
		Pause
	}
}

;Sleep, 500
;send, {Home 2}

sleep, 1000
mousemove, 1000, 800
click
send, ^f
sleep, 200
comment = tatus:
send, %comment%
sleep, 2000
send, {tab 2}
sleep, 2000
send, i
sleep, 2000
send, {tab 20}
sleep, 2000
SendRaw, Information copied to the spreadsheet.
sleep, 2000
send, {tab 6}
sleep, 5000
send, {enter}
sleep, 5000
send, {enter}
sleep, 5000
Comment = OZ Ticket Updated
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Sleep, 1000
Send, {Tab 3}
Sleep, 500
Goto, EXCEL



EXCEL:
;WinWait, Microsoft Excel, 
;IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
;WinWaitActive, Microsoft Excel, 

Sleep, 200
Send, {tab}
Sleep, 200
Send, %Comment%
Sleep, 200
Send, {DOWN}
Sleep, 200
Send, {LEFT}
Sleep, 1000
Clipboard = 

}
Pause::pause

Return