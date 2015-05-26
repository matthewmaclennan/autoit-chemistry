#RequireAdmin
#include <Constants.au3>
#include <MsgBoxConstants.au3>

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         Matthew S. MacLennan

 Script Function:
	Fill in form in SimDCCE2.exe

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

Local $iPID = ShellExecute("C:\Users\Matthew S. MacLennan\Documents\CE Simulator\SimDCCE2\SimDCCE2.exe")
Local $wait = WinWait($iPID,"",4)
; Unfortunately, the form to fill is a BCGToolBar object, so I cannot fill it as per usual. I must resort to mouse clicking. Make sure the SimDCCE2 window fits the screen
; but is not maximized. Check the mouse coordinates listed.
MouseClick("left",20,431,1,0)
MouseClick("left",20,158,1,0)
MouseClick("left",99,168,2,0)
Send("+{END}{BS}55.5")
MouseClick("left",99,185,2,0)
Send("+{END}{BS}55.6")
MouseClick("left",99,202,2,0)
Send("+{END}{BS}55.7")
Send("^s")
$hWnd = WinWait("Save As", "")
ControlSetText($hWnd, "", "[CLASSNN:Edit1]", "Report_" & @YEAR & "-" & @MON & "-" & @MDAY & @HOUR & @MIN & @SEC & ".par")
ControlClick($hWnd,"","[CLASSNN:Button2]")
$mainW = WinWait("Report","")
; MouseClick("left",514, 688,1,0)
; Click &START button
ControlClick($mainW,"","[CLASSNN:Button14]")
; Wait for run to progress
Sleep(10000)
; stop run
ControlClick($mainW,"","[CLASSNN:Button15]")
; name dialogue box
$term = WinWait("Termination?","")
; Click &Yes
ControlClick($term,"","&Yes")
; Open new file (for end of loop)
Send("^n")
; Close program
;ProcessClose($iPID)
