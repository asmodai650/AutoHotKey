;Optum Direct Connect
;Select United Healthcare (Payer)>Inventory>Search



#SingleInstance Force
;WinWait Search Inventory

IfWinNotExist Microsoft Excel - OPD Appeal and Inquiry Resolution Report
  {
   Msgbox The spreadsheet must be named "OPD Appeal and Inquiry Resolution Report" in order for the macro to work. Thanks.
   ExitApp
  }

Beginning:
Loop
{
;ADDED LINES 20 - 27 TO KEEP MACRO FROM STALLING IF TIME EQUALS :00, :15, :30, :45, - DCURTIS 20160719
;If Time Equals :00, :15, :30, :45, pause script for 60 seconds
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

WinActivate Microsoft Excel - OPD Appeal and Inquiry Resolution Report
WinWaitActive Microsoft Excel - OPD Appeal and Inquiry Resolution Report
Sleep 300
Send {HOME}{SHIFTDOWN}{RIGHT}{SHIFTUP}
ClipBoard =
ResolutionDescription =
Send ^c
ClipWait

StringSplit, cell, clipboard, %A_Tab%,

ClaimID = %cell1%
ResolutionDescription = %cell2%
StringReplace, ResolutionDescription, ResolutionDescription, `r`n, , All
RegExReplace(ResolutionDescription, "[`r`n`t]+$")
ClipBoard =

If ResolutionDescription =
{
Msgbox The macro is complete.
ExitApp
}

SearchInventory:
IfWinNotExist Search Inventory
  {

IfWinExist My Inventory
{
WinActivate My Inventory
WinWaitActive My Inventory
}

IfWinExist Partner Account View
{
WinActivate Partner Account View
WinWaitActive Partner Account View
}


IfWinExist Frontier
{
WinActivate Frontier
WinWaitActive Frontier
Sleep 500
Send {BACKSPACE}
Sleep 5000
}

IfWinExist Error
{
WinActivate Error
WinWaitActive Error
}

Sleep 300
If A_UserName contains abaugh1,cmoor61
MouseMove 1234, 200
Else
MouseMove 980, 200
Sleep 500
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
Click 985, 250
Sleep 2000
  }

WinWait, Search Inventory, ,3
IfWinNotExist Search Inventory
  GoTo SearchInventory

SearchInventory2:
WinWait, Search Inventory, ,3

IfWinNotExist Search Inventory
{
Send {BACKSPACE}
GoTo SearchInventory
}

WinActivate Search Inventory
WinWaitActive Search Inventory
Sleep 300

Send ^f
Sleep 300
Send claim number{ESC}
Sleep 300

Send u{TAB}%ClaimID%{ENTER}
Sleep 2000

Click 645, 175
Sleep 1000

Send ^a
Sleep 300

Send ^c
ClipWait, 2

If ClipBoard contains No items to display
{
Sleep 300
If A_UserName contains abaugh1,cmoor61
MouseMove 1234, 200
Else
MouseMove 980, 200
Sleep 500
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
Click 985, 250
Sleep 2000
Sleep 4000
  GoTo SearchInventory2
}

If ClipBoard not contains In Process
  If ClipBoard contains %ClaimID%
   {
   Send {HOME}
   Sleep 300
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   MouseMove 980, 200
   Sleep 500
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   Click 985, 250
   Sleep 1000
   WinActivate Microsoft Excel - OPD Appeal and Inquiry Resolution Report
   WinWaitActive Microsoft Excel - OPD Appeal and Inquiry Resolution Report
   Sleep 300
   Send {RIGHT 2}Not {"}In Process{"}{DOWN}
   ClipBoard =
   GoTo Beginning
   }
  Else
   GoTo SearchInventory

ClipBoard =
click 65, 438, 2
Sleep 300

;Send ^f
;Sleep 300
;Send in process{ESC}
;Sleep 1000
;msgbox %A_CaretX%, %A_CaretY%
;ControlGetPos , X, Y, Width, Height, ListView20WndClass1, Claim ID

;Send +{TAB}

;Sleep 500
;Send {ENTER}

WinWaitClose Search Inventory
WinWait, Partner Account View, ,3

IfWinNotExist Partner Account View
{
Send {BACKSPACE}
Sleep 2000
GoTo SearchInventory2
}


;Checking for Window
OpenRequests:
{
ClipBoard =

Send ^a
Sleep 300

Send ^c
ClipWait, 2
}

;ADDED 204-227 TO KEEP MACRO FROM STALLING IF THERE IS NO OPEN REQUEST FOR CLAIM - DCURTIS 20160617
;If ClipBoard not contains Open Requests
  ;GoTo OpenRequests
If ClipBoard not contains Open Requests
  If ClipBoard contains %ClaimID%
   {
   Send {HOME}
   Sleep 300
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   MouseMove 980, 200
   Sleep 500
If A_UserName contains abaugh1,cmoor61
Click 1250, 250
Else
   Click 985, 250
   Sleep 1000
   WinActivate Microsoft Excel - OPD Appeal and Inquiry Resolution Report
   WinWaitActive Microsoft Excel - OPD Appeal and Inquiry Resolution Report
   Sleep 300
   Send {RIGHT 2}No {"}Open Requests{"}{DOWN}
   ClipBoard =
   GoTo Beginning
   }
  Else
   GoTo SearchInventory

ClipBoard =
click 65, 438, 2
Sleep 300

Send ^f
Sleep 300
Send open requests{ESC}
Sleep 300

Send {TAB}{ENTER}
Sleep 500

;Checking for Window
RequestResponse:
ClipBoard =

Send ^a
Sleep 300

Send ^c
ClipWait
RequestResponseData :=ClipBoard
ClipBoard =

If RequestResponseData not contains Response Comments
  GoTo RequestResponse

;ADDED LINES 269 - 276 TO KEEP MACRO FROM STALLING IF TIME EQUALS :00, :15, :30, :45, - DCURTIS 20160719
;If Time Equals :00, :15, :30, :45, pause script for 60 seconds
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

ClipBoard =
ClipBoard := ResolutionDescription
Send ^f
Sleep 300
Send request comments{ESC}
Sleep 300
Send {TAB 2}n{tab}^v{TAB}{ENTER}
ClipBoard =

If RequestResponseData contains Response Date
{
Sleep 5000
Send ^f
Sleep 300
Send dialog{TAB}{ENTER}{ESC}
Sleep 300
Send {TAB}{ENTER}
Sleep 1000
Send ^f
Sleep 300
Send category{ESC}
Sleep 300
Send {TAB 2}{DOWN 6}{TAB 3}
Sleep 300

;ADDED LINES 305 - 312 TO KEEP MACRO FROM STALLING IF TIME EQUALS :00, :15, :30, :45, - DCURTIS 20160719
;If Time Equals :00, :15, :30, :45, pause script for 60 seconds
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

ClipBoard =
ClipBoard := ResolutionDescription

Send ^v{TAB}{ENTER}
Sleep 500
ClipBoard =
}
Sleep 3000

;Checking for Window
AccountView:
ClipBoard =

Send ^a
Sleep 300

Send ^c
ClipWait

If ClipBoard not contains Account View
  GoTo AccountView

ResolveAccount:
Send ^f
Sleep 300
Send edit
Sleep 300
Send {ESC}
Sleep 300
Send +{TAB}
Sleep 300

;ADDED LINES 348 - 355 TO KEEP MACRO FROM STALLING IF TIME EQUALS :00, :15, :30, :45, - DCURTIS 20160719
;If Time Equals :00, :15, :30, :45, pause script for 60 seconds
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

Send {ENTER}
Sleep 500

;Checking for Window

ClipBoard =

Send ^a
Sleep 300

Send ^c
ClipWait, 2

If ClipBoard not contains Resolve Reason
  GoTo ResolveAccount

Sleep 300

Send ^f
Sleep 300
Send resolve reason{ESC}
Sleep 500
Send {TAB}c
Sleep 300

;ADDED LINES 384 - 391 TO KEEP MACRO FROM STALLING IF TIME EQUALS :00, :15, :30, :45, - DCURTIS 20160719
;If Time Equals :00, :15, :30, :45, pause script for 60 seconds
if A_Min between 00 and 01
	sleep, 60000
if A_Min between 15 and 16
	sleep, 60000
if A_Min between 30 and 31
	sleep, 60000
if A_Min between 45 and 46
	sleep, 60000

Send {TAB}{ENTER}
Sleep 7000

WinActivate Microsoft Excel - OPD Appeal and Inquiry Resolution Report
WinWaitActive Microsoft Excel - OPD Appeal and Inquiry Resolution Report
Sleep 300
Send {RIGHT 2}x{DOWN}

}

pause::pause