#include <Array.au3>
#include <Constants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#include <Timers.au3>
#include <WindowsConstants.au3>
#NoTrayIcon

;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win10
; Author:         Leone Tolesano (0xleone | leone@lone.sh)
;
; Description:
;    Emulate user activity randomly opening (via Explorer - like a user) and using (currently):
;    Edge, Outlook, Excel, Word, Notepad and Calc
;
; Usage:
;    Run the app normally. When you want it to finish, hit the close button while "Idle" or
;    wait the random chance of auto-closing (10% each interaction).
;    You can also force kill with SHIFT+ESC combination (The "trap" is unreliable, so prefer
;    the UI close button via mouse click)
;
; Requirements:
;    OS: Win10
;    Win10: Power & Sleep settings: Never/Never
;    Edge: Page Layout: Custom: Disable both checkboxes. Background: Disabled; Content: Disabled.
;    Outlook: E-mail pre-configured. Otherwise read and send e-mails will fail.
;    Excel, Word, Notepad, Calc: Nothing specific. It must open without any warnings or prompts
;    and able to edit files (obviously don't use with read-only MS Office version)
;
; Usage:
; - Just run the app.
; - If needed, press SHIFT+ESC to terminate .
;


;;;;;;;;;;;;;;;;;;;;
;
; General Settings
;
;;;;;;;;;;;;;;;;;;;;

Global $hotkey = "+{ESC}"
HotKeySet($hotkey, "Quit") 						; SHIFT+ESC to quit
Global Const $typingSpeed = 75					; Typing delay between chars
AutoItSetOption("SendKeyDelay", $typingSpeed)
Global Const $enableGUI = True					; Enable/Disable debug screen
Global Const $workDir = @TempDir & '\NotARobot'	; Guarantee directory permission no matter the user running this.
Global $minutes = 60000
Global $seconds = 1000

Global Const $numApps = 6						; Self-explanatory -> Apps have fixed position - see $ExecApp definition.
Global $listWin[$numApps]						; Array of opened apps to avoid reopen and easy exiting
Global $listFiles = [0]							; Dynamic array - increases for each file created
Global $ExecApp[$numApps]   = [ fCalc, fNotepad, fEdge, fWord, fExcel, fOutlook ]
Global $labelApps[$numApps] = [ "Calc", "Notepad", "Edge", "Word", "Excel", "Outlook" ]
Global $fDebug, $lAction, $lInfo				; GUI
Global $randomWin								; Current App running


;;;;;;;;;;;;;;;;;;;;
;
; Apps Settings
;
;;;;;;;;;;;;;;;;;;;;

$pathEdge  = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$pathMail  = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.exe"
$pathExcel = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe"
$pathWord  = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.exe"
$pathNote  = "C:\Windows\System32\notepad.exe"
$pathCalc  = "C:\Windows\System32\calc.exe"

; Edge
Global Enum $twitter, $facebook, $instagram, $nytimes, $google, $youtube, $hobbits, $rick
Global Const $num_sites = 8
Global $sites[$num_sites]
$sites[$twitter]   = 'https://www.twitter.com'
$sites[$facebook]  = 'https://www.facebook.com'
$sites[$instagram] = 'https://www.instagram.com'
$sites[$nytimes]   = 'https://www.nytimes.com'
$sites[$google]    = 'https://www.google.com'
$sites[$youtube]   = 'https://www.youtube.com'
$sites[$hobbits]   = 'https://www.youtube.com/watch?v=uE-1RPDqJAY'
$sites[$rick]      = 'https://www.youtube.com/watch?v=dQw4w9WgXcQ'


;;;;;;;;;;;;;;;;;;;;
;
; Main code
;
;;;;;;;;;;;;;;;;;;;;


Main()


;;;;;;;;;;;;;;;;;;;;
;
; Functions
;
;;;;;;;;;;;;;;;;;;;;

; Main Code
Func Main()
	If $enableGUI Then
		$fDebug  = GUICreate("NotARobot", 143, 78, -1, -1, Default, BitOr($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
		$lAction = GUICtrlCreateLabel("Idle", 27, 2, 94, 37, $SS_CENTER)
		GUICtrlSetFont(-1, 20, 400, 0, "Calibri")

		$lInfo = GUICtrlCreateLabel("--:--", 11, 35, 127, 37, $SS_CENTER)
		GUICtrlSetFont(-1, 20, 400, 0, "Calibri")
		GUICtrlSetColor(-1, 0x800000)

		$pos = WinGetPos($fDebug)
		WinMove($fDebug, "", @DesktopWidth-$pos[2], @DesktopHeight-$pos[3]-40);40 = START task/tray menu
		GUISetState(@SW_SHOW)
	EndIf

	DirRemove($workDir, 1)
	DirCreate($workDir)
	While 1
		; Randomize running or closing app
		$randomWin = Random(0,($numApps - 1),1)
		If Not IsHWnd($listWin[$randomWin]) Then
			$ExecApp[$randomWin]()
		Else
			CloseApp($listWin[$randomWin])
		EndIf

		; 10% chance to exit the simulation
		If Random(0,99,1) >= 90 Then
			Quit(0)
		EndIf

		; Wait from 3 seconds to 15 minutes
;~ 		TrapSleep(Random(3*$seconds,15*$minutes,1), True)
		TrapSleep(5000, True)
	WEnd
EndFunc

; Exit closing all apps
Func Quit($ret = 0)
	For $w = 0 To ($numApps - 1)
		CloseApp($listWin[$w])
	Next
	If @HotKeyPressed == $hotkey Then
		Exit 2
	Else
		Exit $ret
    EndIf
EndFunc

; Close an app by handle
Func CloseApp($handle)
   If IsHWnd($handle) Then
	  WinKill($handle)
   EndIf
EndFunc

; Show Current App on GUI
Func ShowCurrentApp($action, $info)
	If $enableGUI Then
		GUICtrlSetData($lAction, $action)
		GUICtrlSetData($lInfo, $info)
		GUICtrlSetColor($lInfo, ($action <> "Idle") ? 0x0078D7 : 0x800000)
	EndIf
EndFunc

; Sleep timer on GUI
Func TrapSleep($delay, $show = False)
	Local $countdown = Int($delay/1000)

	If $enableGUI And $show Then
		ShowCurrentApp("Idle", "--:--")
		Local $timerMin = Int($delay/60000)
		Local $timerSec = Int(($delay - $timerMin * 60000) / 1000)
		Local $timerTxt
	EndIf
	Do
		Local $tBegin = TimerInit()
		Do
			If GUIGetMsg() = $GUI_EVENT_CLOSE Then
				Quit(1)
			ElseIf @HotKeyPressed == $hotkey Then
				Quit(2)
			EndIf
		Until TimerDiff($tBegin) > 1000
		$countdown -= 1

		If $enableGUI And $show Then
			$timerSec -= 1
			If $timerSec == 0 Then
				$timerSec = 59
				$timerMin -= 1
			EndIf
			$timerTxt  = ($timerMin < 10) ? "0" & $timerMin & ":" : $timerMin & ":"
			$timerTxt  = $timerTxt & (($timerSec < 10) ? "0" & $timerSec : $timerSec)

			ShowCurrentApp("Idle", $timerTxt)
			GUISetState(@SW_SHOW)
		EndIf
	Until $countdown <= 0

	If $enableGUI And $show Then
		ShowCurrentApp("Running", $labelApps[$randomWin])
	EndIf
EndFunc

; Send keys with close button handler
Func TrapSend($keys)
	Local $charArray = StringSplit($keys, "")
	Local $tBegin = TimerInit()

	For $i = 1 To $charArray[0]
		Do
			If GUIGetMsg() = $GUI_EVENT_CLOSE Then
				Quit(1)
			ElseIf @HotKeyPressed == $hotkey Then
				Quit(2)
			EndIf
		Until TimerDiff($tBegin) > $typingSpeed
		Send($charArray[$i])
	Next
	AutoItSetOption("SendKeyDelay", $typingSpeed)
EndFunc

; Open Explorer and navigate to desired directory
Func fOpenDir ($fullpath)
	; Open Explorer and goes to address bar
	ShowCurrentApp("Running", "Explorer")
	Send("{LWINDOWN}e{LWINUP}")
	TrapSleep(1000)
	Send("!d")

	; Goes to program directory
	TrapSend($fullpath)
	Send("{ENTER}")
	TrapSleep(1000)

	Return WinGetHandle(WinGetTitle("[ACTIVE]"))
EndFunc

; Open App via Explorer
Func fOpenApp($fullpath)
	Local $split[5]
	_PathSplit($fullpath, $split[1], $split[2], $split[3], $split[4])
	Local $path   = $split[1] & $split[2]
	Local $target = $split[3] & $split[4]
	Local $win = fOpenDir($path)

	; Run desired program and kill Explorer
	TrapSend($target)
	Send("{ENTER}")
	ShowCurrentApp("Running", $labelApps[$randomWin])
	While 1
		If WinWaitNotActive($win) Then
			WinKill($win)
			ExitLoop
		EndIf
    WEnd
	$listWin[$randomWin] = WinGetHandle(WinGetTitle("[ACTIVE]"))
EndFunc

; Delete File via Explorer with Shift+Delete and confirms prompt.
Func fDelete($fullpath = $workDir)
	If FileExists($fullpath) = 1 Then
		Local $split[5]
		_PathSplit($fullpath, $split[1], $split[2], $split[3], $split[4])
		Local $path   = $split[1] & $split[2]
		Local $target = $split[3] & $split[4]
		Local $win = fOpenDir($path)

		; Delete
		TrapSend($target)
		Send("+{DEL}")
		TrapSleep(2000)
		Send("{ENTER}")
		TrapSleep(2000)

		WinKill($win)
		TrapSleep(2000)
	EndIf
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Calc.exe
;
;;;;;;;;;;;;;;;;;;;;

Func fCalc()
    fOpenApp($pathCalc)
    Local $calc = ""
    Local $randomCalc = (Random(3,5,1)*2)-1

    For $i = 0 To $randomCalc
	    If $i = 0 Or Mod($i,2) <> 0 Then
		    $calc &= Random(1,9,1)
	    Else
			Switch Random(0,3,1)
				Case 0
					$calc &= "+"
				Case 1
					$calc &= "-"
				Case 2
					$calc &= "*"
				Case 3
					$calc &= "/"
		    EndSwitch
	    EndIf
    Next

    $calc &= "="
	TrapSend($calc)
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Notepad.exe
;
;;;;;;;;;;;;;;;;;;;;

Func fNotepad()
	Local $actionOptions = 2
	Local $randomChoice  = Random(1,$actionOptions,1)

	TrapSleep(2000)
	If $listFiles[0] = 0 Or $randomChoice = 1 Then
		; Open Notepad and create a random file
		fOpenApp($pathNote)
		fNotepadCreateFile()
	Else
		; Delete a random previously created file
		fDelete($listFiles[Random(0,UBound($listFiles)-1,1)])
	EndIf
EndFunc

Func fNotepadCreateFile()
	; Check if directory exists
	If FileExists($workDir) = False Then
		DirCreate($workDir)
	EndIf

	; Add to the files array
	If $listFiles[0] = 0 Then
		$listFiles[0] = $workDir & "\" & Random(999,9999999,1) & ".txt"
	Else
		ReDim $listFiles[UBound($listFiles)+1]
		$listFiles[UBound($listFiles)-1] = $workDir & "\" & Random(999,9999999,1) & ".txt"
	EndIf

	; Save via Notepad Save As menu
	TrapSend(Random(999,9999999,1) & Random(999,9999999,1) & Random(999,9999999,1))
	Send("{ENTER}" & "^S")
	TrapSleep(1000)
	TrapSend($listFiles[UBound($listFiles)-1])
	Send("{ENTER}")
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Edge
;
;;;;;;;;;;;;;;;;;;;;

Func fEdge()
	fOpenApp($pathEdge)
	Local $rSite, $rSleep
	Local $numRandomSites = Random(1,$num_sites,1)
	Local $tabs = 0
	Do
		$rSite = Random(0,($num_sites - 1),1)

		If $tabs > 0 Then
			Send("^t")
		TrapSleep(1000)
		EndIf

		TrapSend($sites[$rSite])
		Send("{ENTER}")

		If $rSite == $google Then
			TrapSleep(1000)
			TrapSend(Random(99,999999,1))
			Send("{ENTER}")
		EndIf

		TrapSleep(Random(3000, 6000, 1))

		$tabs += 1
		; 30% chance to switch tab and refresh
		If $tabs > 1 Then
			If Random(1, 10, 1) > 7 Then
				For $i = 1 To Random(1, $numRandomSites, 1)
					Send("^{TAB}")
					TrapSleep(1000)
				Next
			Send("{F5}")
			EndIf
		EndIf

		TrapSleep(Random(3000, 6000, 1))
	Until $tabs >= $numRandomSites
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Word
;
;;;;;;;;;;;;;;;;;;;;

Func fWord()
	fOpenApp($pathWord)
	TrapSleep(3000)
	Send("{ENTER}")

;	TODO:
;	- Create a random file
;	- Delete one of the previous files created
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Excel
;
;;;;;;;;;;;;;;;;;;;;

Func fExcel()
	fOpenApp($pathExcel)
	TrapSleep(3000)
	Send("{ENTER}")

;	TODO:
;	- Create a random file
;	- Delete one of the previous files created
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Outlook
;
;;;;;;;;;;;;;;;;;;;;

Func fOutlook()
	fOpenApp($pathMail)

;	TODO:
;	- Create a random e-mail and send.
;	- Open e-mail attachments.
EndFunc