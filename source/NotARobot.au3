#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_x64=NotARobot_v0.94-beta.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_ProductName=NotARobot
#AutoIt3Wrapper_Res_ProductVersion=0.94-beta
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Add_Constants=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /mo /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Array.au3>
#include <AutoItConstants.au3>
#include <Constants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#include <StringConstants.au3>
#include <Timers.au3>
#include <WindowsConstants.au3>

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
; - If needed, press SHIFT+ESC to terminate script.
;


;;;;;;;;;;;;;;;;;;;;
;
; General Settings
;
;;;;;;;;;;;;;;;;;;;;

Global $hotkey = "+{ESC}"
HotKeySet($hotkey, "Quit")                     ; SHIFT+ESC to quit
Global $typingDelay = 50                       ; Typing delay between chars
AutoItSetOption("SendKeyDelay", $typingDelay)
Global $enableGUI = True                       ; Enable/Disable debug screen
Global $workDir = @TempDir & '\NotARobot'      ; Guarantee directory permission no matter the user running this.
Global $randomWindow = 15                      ; Range between 3 seconds and X Minutes between executions

Global $runningApps[][] = [[0, 0]]             ; Array of app handles and names for easy exiting
Global $filesCreated[] = [0]                   ; Dynamic array - increases for each file created
Global $fDebug, $lAction, $lInfo               ; GUI
Global $currentApp = 0                         ; Dynamic label of current app running
Global $ExecApp[] = [0]                        ; Dynamic list of apps
Global $sites[] = ['https://www.google.com']   ; Sites to open via browser; Minimum is random Google search (built-in)

;;;;;;;;;;;;;;;;;;;;
;
; Apps Settings
;
;;;;;;;;;;;;;;;;;;;;

Global $calc     = "C:\Windows\System32\calc.exe"
Global $notepad  = "C:\Windows\System32\notepad.exe"
Global $edge     = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
Global $word     = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.exe"
Global $excel    = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe"
Global $outlook  = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.exe"
Global $snipTool = "C:\Windows\System32\SnippingTool.exe"

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
	Local $search

	fLoadConfig()
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

	Local $minutes = 60000
	Local $seconds = 1000
	While 1
		; Run or close an app
		$ExecApp[Random(0,(UBound($ExecApp)-1),1)]()

		; 10% chance to exit the simulation
		If Random(0,99,1) >= 90 Then
			Quit(0)
		EndIf

		; Wait from 3 seconds to 15 minutes
		TrapSleep(Random(3*$seconds,$randomWindow*$minutes,1), True)
	WEnd
EndFunc

; Dynamic read config.ini file.
Func fLoadConfig()
	Local $configFilePath = (@Compiled = 1) ? fSplitPath(@AutoItExe)[1] & fSplitPath(@AutoItExe)[2] : @ScriptDir & "\"
	Local $handleFileOpen = FileOpen($configFilePath & 'config.ini', $FO_READ)

	; Open File and Test
	If $handleFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "NotARobot", "Please create a config file.")
		Exit 1
	EndIf

	; Read the contents of the file using the handle returned by FileOpen.
	Local $configFile = FileReadToArray($handleFileOpen)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "NotARobot", "There was an error reading config file. @error: " & @error)
		Exit 1
	Else
		Local $line, $key, $value, $pos
		For $line = 0 To (UBound($configFile) - 1)
			$keyvalue = StringSplit($configFile[$line], ';', $STR_NOCOUNT)[0]
			If $keyvalue == "" Or StringInStr($keyvalue, '[') <> 0 Or StringInStr($keyvalue, ']') <> 0 Then
				ContinueLoop
			EndIf

			$key   = StringStripWS(StringSplit($keyvalue, "=", $STR_NOCOUNT)[0], BitOr($STR_STRIPLEADING, $STR_STRIPTRAILING))
			$value = StringStripWS(StringTrimLeft($keyvalue, StringLen($key)+1), BitOr($STR_STRIPLEADING, $STR_STRIPTRAILING))
			Switch StringLower($key)
				Case "typingdelay"
					If StringIsInt($value) And $value < 1000 Then
						$typingDelay = ($value <= 5) ? 5 : $value
					EndIf
					AutoItSetOption("SendKeyDelay", $typingDelay)
					ContinueLoop
				Case "enablegui"
					If $value = 1 Then
						$enableGUI = True
					ElseIf $value = 0 Then
						$enableGUI = False
					EndIf
					ContinueLoop
				Case "randomwindowminutes"
					If StringIsInt($value) Then
						$randomWindow = $value
					EndIf
					ContinueLoop
				Case "workdir"
					DirCreate($value)
					If FileExists($value) = 1 Then
						$workDir = $value
					EndIf
					ContinueLoop
				Case "enablecalc"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fCalc
					EndIf
					ContinueLoop
				Case "enablenotepad"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fNotepad
					EndIf
					ContinueLoop
				Case "enableedge"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fEdge
					EndIf
					ContinueLoop
				Case "enableword"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fWord
					EndIf
					ContinueLoop
				Case "enableexcel"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fExcel
					EndIf
					ContinueLoop
				Case "enableoutlook"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fOutlook
					EndIf
					ContinueLoop
				Case "enablesnippingtool"
					If $value = 1 Then
						$pos = ($ExecApp[0] = 0) ? 0 : UBound($ExecApp)
						If $pos > 0 Then
							ReDim $ExecApp[$pos+1]
						EndIf
						$ExecApp[$pos] = fSnipTool
					EndIf
					ContinueLoop
				Case "calc"
					$calc     = (FileExists($value) = 1) ? $value : $calc
					ContinueLoop
				Case "notepad"
					$notepad  = (FileExists($value) = 1) ? $value : $notepad
					ContinueLoop
				Case "edge"
					$edge     = (FileExists($value) = 1) ? $value : $edge
					ContinueLoop
				Case "word"
					$word     = (FileExists($value) = 1) ? $value : $word
					ContinueLoop
				Case "excel"
					$excel    = (FileExists($value) = 1) ? $value : $excel
					ContinueLoop
				Case "outlook"
					$outlook  = (FileExists($value) = 1) ? $value : $outlook
					ContinueLoop
				Case "snippingtool"
					$snipTool = (FileExists($value) = 1) ? $value : $snipTool
					ContinueLoop
				Case "site"
					$pos = UBound($sites)
					ReDim $sites[$pos+1]
					$sites[$pos] = $value
					ContinueLoop
			EndSwitch
		Next
	EndIf

	If FileExists($workDir) = 0 Then
		DirCreate($workDir)
	EndIf

	; Close the handle returned by FileOpen.
	FileClose($handleFileOpen)
EndFunc

; Exit closing all apps
Func Quit($ret = 0)
	; If no current running apps, just finish this script
	Local $apps = ($runningApps[0][0] = 0) ? -1 : (UBound($runningApps) - 1)
	If $apps = -1 Then
		Exit $ret
	Else
		For $w = 0 To $apps
			CloseApp($runningApps[$w][0])
		Next
	EndIf

	; Delete all files
	Local $files = ($filesCreated[0] = 0) ? -1 : (UBound($filesCreated) - 1)
	If $files <> -1 Then
		For $f = 0 To $files
			fDelete($filesCreated[$f])
		Next
	EndIf

	; Exit Codes
	If @HotKeyPressed == $hotkey Then
		Exit 2
	Else
		Exit $ret
    EndIf
EndFunc

; Close an app by handle
Func CloseApp($handle)
	Local $pos = _ArraySearch($runningApps, $handle)
	If IsHWnd($handle) Then
		WinKill($handle)
		If UBound($runningApps) > 1 Then
			_ArrayDelete($runningApps, $pos)
		Else
			$runningApps[0][0] = 0
			$runningApps[0][1] = 0
		EndIf
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
		ShowCurrentApp("Running", $currentApp)
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
		Until TimerDiff($tBegin) > $typingDelay
		Send($charArray[$i])
	Next
	AutoItSetOption("SendKeyDelay", $typingDelay)
EndFunc

; Split path string
Func fSplitPath($fullpath)
	Local $split[5]
	_PathSplit($fullpath, $split[1], $split[2], $split[3], $split[4])

	Return $split
EndFunc

; Open Explorer and navigate to desired directory
Func fOpenDir ($fullpath)
	; Open Explorer and goes to address bar
	ShowCurrentApp("Running", "explorer")
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
	Local $split = fSplitPath($fullpath)
	Local $path   = $split[1] & $split[2]
	Local $target = $split[3] & $split[4]
	$currentApp = StringLower($split[3])
	Local $pos = _ArraySearch($runningApps, $currentApp)
	If $runningApps[0][0] = 0 Or ($pos = -1 And @error = 6) Then
		; Run desired program and kill Explorer
		Local $win = fOpenDir($path)

		Send("!d{RIGHT}")
		TrapSend('\' & $target)
		Send("{ENTER}")
		ShowCurrentApp("Running", $currentApp)
		While 1
			If WinWaitNotActive($win) Then
				WinKill($win)
				ExitLoop
			EndIf
		WEnd

		; Add window to the array
		Local $row = ($runningApps[0][0] = 0) ? 0 : UBound($runningApps, $UBOUND_ROWS)
		Local $columns = UBound($runningApps, $UBOUND_COLUMNS)

		ReDim $runningApps[$row+1][$columns]
		$runningApps[$row][0] = WinGetHandle(WinGetTitle("[ACTIVE]"))
		$runningApps[$row][1] = $currentApp

		Return True
	Else
		CloseApp($runningApps[$pos][0])
		Return False
	EndIf

EndFunc

; Delete File via Explorer with Shift+Delete and confirms prompt.
Func fDelete($fullpath)
	If FileExists($fullpath) = 1 Then
		Local $split = fSplitPath($fullpath)
		Local $path   = $split[1] & $split[2]
		Local $target = $split[3] & $split[4]
		Local $win = fOpenDir($path)

		; Delete
		TrapSend($target)
		Send("+{DEL}")
		TrapSleep(2000)
		Send("{ENTER}")

		While 1
			If WinWaitNotActive($win) Then
				WinKill($win)
				ExitLoop
			EndIf
		WEnd
	EndIf
EndFunc

; Create random files
Func fCreateFile($app)
	; Check if directory exists
	If FileExists($workDir) = False Then
		DirCreate($workDir)
	EndIf

	; Add to the files array
	Local $pos = ($filesCreated[0] = 0) ? 0 : UBound($filesCreated)
	If $pos > 0 Then
		ReDim $filesCreated[$pos+1]
	EndIf
	$filesCreated[$pos] = $workDir & "\" & Random(999,9999999,1)

	; Save file
	TrapSend(Random(999,9999999,1) & Random(999,9999999,1) & Random(999,9999999,1))
	Send("{ENTER}")
	Switch $app
		Case "notepad"
			$filesCreated[$pos] &= ".txt"
			Send("^s")
		Case "word"
			$filesCreated[$pos] &= ".docx"
			Send("{F12}")
		Case "excel"
			$filesCreated[$pos] &= ".xlsx"
			Send("{F12}")
	EndSwitch
	TrapSleep(1000)
	TrapSend($filesCreated[$pos])
	Send("{ENTER}")
EndFunc

;;;;;;;;;;;;;;;;;;;;
;
; Calc
;
;;;;;;;;;;;;;;;;;;;;

Func fCalc()
	If Not fOpenApp($calc) Then
		Return
	EndIf

    Local $c = ""
    Local $randomCalc = (Random(3,5,1)*2)-1

    For $i = 0 To $randomCalc
	    If $i = 0 Or Mod($i,2) <> 0 Then
		    $c &= Random(1,9,1)
	    Else
			Switch Random(0,3,1)
				Case 0
					$c &= "+"
				Case 1
					$c &= "-"
				Case 2
					$c &= "*"
				Case 3
					$c &= "/"
		    EndSwitch
	    EndIf
    Next

    $c &= "="
	TrapSend($c)
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Notepad
;
;;;;;;;;;;;;;;;;;;;;

Func fNotepad()
	Local $actionOptions = 2
	Local $randomChoice  = Random(1,$actionOptions,1)

	TrapSleep(2000)
	If $filesCreated[0] = 0 Or $randomChoice = 1 Then
		; Close Notepad OR open it and create a random file
		If Not fOpenApp($notepad) Then
			Return
		Else
			fCreateFile("notepad")
		EndIf
	Else
		; Delete a random previously created file via Explorer
		fDelete($filesCreated[Random(0,UBound($filesCreated)-1,1)])
	EndIf
EndFunc

;;;;;;;;;;;;;;;;;;;;
;
; SnippingTool
;
;;;;;;;;;;;;;;;;;;;;

Func fSnipTool()
	If Not fOpenApp($snipTool) Then
		Return
	EndIf

	; Take screenshot
	TrapSend("!m")
	Sleep(200)
	TrapSend("s")
	TrapSleep(1000)

	; Add screenshot to the files array
	Local $pos = ($filesCreated[0] = 0) ? 0 : UBound($filesCreated)
	If $pos > 0 Then
		ReDim $filesCreated[$pos+1]
	EndIf
	$filesCreated[$pos] = $workDir & "\" & Random(999,9999999,1) & ".png"

	; Save screenshot in workDir
	Send("^s")
	TrapSleep(1000)
	TrapSend($filesCreated[$pos])
	Send("{ENTER}")
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Edge
;
;;;;;;;;;;;;;;;;;;;;

Func fEdge()
	If Not fOpenApp($edge) Then
		Return
	EndIf

	Local $rSite, $rSleep
	Local $numRandomSites = Random(1,UBound($sites),1)
	Local $tabs = 0
	Do
		$rSite = Random(0,(UBound($sites)-1),1)

		If $tabs > 0 Then
			Send("^t")
		TrapSleep(1000)
		EndIf

		TrapSend($sites[$rSite])
		Send("{ENTER}")

		If $rSite == 0 Then
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
	If Not fOpenApp($word) Then
		Return
	EndIf

	TrapSleep(2000)
	Send("{ENTER}")

	Local $actionOptions = 2
	Local $randomChoice  = Random(1,$actionOptions,1)

	; Create or Delete random file
	If $filesCreated[0] = 0 Or $randomChoice = 1 Then
		fCreateFile("word")
	Else
		; Delete a random previously created file via Explorer
		fDelete($filesCreated[Random(0,UBound($filesCreated)-1,1)])
	EndIf
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Excel
;
;;;;;;;;;;;;;;;;;;;;

Func fExcel()
	If Not fOpenApp($excel) Then
		Return
	EndIf

	TrapSleep(2000)
	Send("{ENTER}")

	Local $actionOptions = 2
	Local $randomChoice  = Random(1,$actionOptions,1)

	; Create or Delete random file
	If $filesCreated[0] = 0 Or $randomChoice = 1 Then
		fCreateFile("excel")
	Else
		; Delete a random previously created file via Explorer
		fDelete($filesCreated[Random(0,UBound($filesCreated)-1,1)])
	EndIf
EndFunc


;;;;;;;;;;;;;;;;;;;;
;
; Outlook
;
;;;;;;;;;;;;;;;;;;;;

Func fOutlook()
	If Not fOpenApp($outlook) Then
		Return
	EndIf

;	TODO:
;	- Create a random e-mail and send.
;	- Open e-mail attachments.
EndFunc

