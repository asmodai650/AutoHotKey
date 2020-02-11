#SingleInstance Force

IfWinNotExist Microsoft Excel - WellCare No Touch List Database
RunWait \\aim.aimhealth.com\client_services$\CS- Client Management Team\Account Specialist\CAS Specific\Brad Bell\All Others\WellCare\Wellcare No Touch\Macro\WellCare No Touch List Database.xlsx

msgbox Make sure the Wellcare information in notepad is named "WellCare List" and then open it.


WinActivate Microsoft Excel - WellCare No Touch List Database
WinWaitActive Microsoft Excel - WellCare No Touch List Database
Sleep 300

Msgbox Please deliminate the spreadsheet by "^" only and undo Word wrap. Once you have done this, select "OK" to continue.


WinActivate Microsoft Excel - WellCare No Touch List Database
WinWaitActive Microsoft Excel - WellCare No Touch List Database
Sleep 300
Send ^{HOME}





;WinActivate WellCare List
;WinWaitActive WellCare List
Sleep 300
;Send ^{HOME}+{END}{DEL 2}

Loop
{
WinActivate WellCare List
WinWaitActive WellCare List
Sleep 2000
Click 200, 200
Sleep 5000


NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 500


Send, ^g
WinWaitActive Go To Line
Sleep 300
Send 1000001{ENTER}
WinWaitClose Go To Line, ,2

IfWinActive Notepad - Goto Line
{
Send {ENTER}
WinWaitClose Notepad - Goto Line
Send {TAB 2}{ENTER}
WinWaitActive WellCare List
Send ^{END}
LastLoop = Yes
}


Sleep 300
Send ^+{HOME}
Sleep 300
ClipBoard =
Send ^x
ClipWait

WinActivate Microsoft Excel - WellCare No Touch List Database
WinWaitActive Microsoft Excel - WellCare No Touch List Database
Sleep 2000
WinGetTitle, ExcelTitle

Sleep 500
Send ^v
Sleep 500

;Send !ae
;WinWaitActive Convert Text
;Sleep 200
;Send dn
;Sleep 300
;ControlClick, EDTBX2
;Sleep 300
;Send {^}
;Sleep 300
;Send f
;WinWaitClose Convert Text

;Send ^{F3}
;Sleep 300
;Send A:B,D:D{ENTER}

Loop 3
{
If A_Index = 3
Send {RIGHT}
Sleep 200
Send {UP}
Sleep 200
Send ^{SPACE}
Sleep 300
Send ^{-}
Sleep 1000

NoBreakOnTheseCursors=AppStarting, Wait 
Loop 
{ 
  Sleep, 100 
  IfNotInString, NoBreakOnTheseCursors, %A_Cursor% 
    Break 
}
Sleep, 500
}

Send {UP}


;New Logic to save separate files 11.8.16
Send ^+s
WinWaitClose %ExcelTitle%
Sleep 9000


If LastLoop = Yes
  Break
}

Send ^{HOME}
Sleep 300
Send ^+{END}
Sleep 300
Send {SHIFTDOWN}{LEFT 3}{SHIFTUP}
Sleep 300
Send ^c
ClipWait

msgbox The macro is finished. The data is now ready to be pasted into a new notepad document.
ExitApp
pause::pause