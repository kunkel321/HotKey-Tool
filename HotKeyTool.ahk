﻿#SingleInstance
#Requires AutoHotkey v2.0

; ==============================================================================
; Title:	    HotKey Lister, Filter'er, and Launcher.
; Author:	    Stephen Kunkel321, with help from Claude.ai
; Version:	    8-4-2024
; GitHub:       https://github.com/kunkel321/HotKey-Tool
; AHK Forum:    https://www.autohotkey.com/boards/viewtopic.php?f=83&t=132224
; ========= INFORMATION ========================================================
; Mostly it's just a "Cheetsheet" list of the hotkeys in your running scripts. 
; Also can be used to launch them though.  Launches via "Sending" the hotkey.
; Double-click an item to launch it.  Enter key launches selected item too.
; Made for "portably-running" scripts.  Those are scripts where a copy
; of AutoHotkey.exe has been put in the same folder as the .ahk file and renamed  
; to match it. For example myScript.ahk and myScript.exe are both in the same folder. 
; If the exe is a compiled version of the ahk, then that should work too. 
; It is actually the ahk file that gets searched for hotkeys.
; Also the folder has to be a subfolder of the one specified via ahkFolder variable.
; Esc only hides form.  Must restart script to rescan hotkeys.
; There's no point scanning files with no hotkeys, so add those to ignore list array.
; Set your scripts up such that there is an in-line comment on each line of code
; that has a hotkey.  Use descriptive "search terms" in the comment.
; To hide individual hotkeys from scan, include the word "hide" (without quotes)
; in the comment.  Lines of code are ignored if they...
; - have the word 'hide'.
; - have single or double quotes.
; - don't contain "::".
; - do contain more than two colons.
; The Filter Box is a ComboBox and can have pre-defined hotkey search filters. 
; - Names of scripts are added automatically, to filter by containing script file.
; - Additional filters can be added to array in USER OPTIONS below.
; There are a few other options below as well.  See also, copious in-line comments.
; Tool will determine active window then wait for it before sending hotkey.
; ==============================================================================

;Below coloring (8) lines are specific to Steve's setup.  If you see them, he apparently forgot to remove them. 
SettingsFile := A_ScriptDir '\WayText\wtFiles\Settings.ini'
gColor := iniread(SettingsFile, "MainSettings", "GUIcolor", "Default")
lColor := iniread(SettingsFile, "MainSettings", "ListColor", "Default")
fColor := iniread(SettingsFile, "MainSettings", "FontColor", "Default")
;-----------------
formColor := strReplace(subStr(gColor, -6), "efault", "Default")
listColor := strReplace(subStr(lColor, -6), "efault", "Default")
fontColor := strReplace(subStr(fColor, -6), "efault", "Default")

; ======= USER OPTIONS =========================================================
; --- Below 3 color assignments should only be commented out for Steve.
; formColor := "00233A" ; Use hex code if desired. Use "Default" for default.
; listColor := "003E67"
; fontColor := "31FFE7"
mainHotkey := "!+q" ; main hotkey to show gui -- Alt+Shift+Q
guiWidth := 600 ; Width of form. (At least 600 recommended, depending on font size.)
maxRows := 16 ; Scroll if more row than this in listview
guiTitle := "Hotkey Tool" ; Change (here only) if desired. 
fontSize := 12 ; Don't include the 's'.
trans := 255 ; Set 0 (transparent) to 255 (opaque).
appIcon := "shell32.dll,278"
ahkFolder := "D:\AutoHotkey" ; We'll find the active hotkeys in scripts running from here.
ignoreList := ["AHK-ToolKit", "HotstringLib", "QuickSwitch", "mwClipboard"] ; We'll skip these files... 
preDefinedFilters := ["#","!+","!^"] ; Pre-defined filters for the combobox. Add more? 
debugMode := 0 ; 1=On.  Shows processes and target window.
; ==============================================================================

If !DirExist(ahkFolder) { ; Make sure folder is there.
	MsgBox "The folder `"" ahkFolder "`" doesn't seem to exist.  Please point the variable `"ahkFolder :=`" to the folder where your scripts are located. Now exiting"
	ExitApp
}

; Determine if icon is .ico or library item, parse as needed.
If InStr(appIcon, ",") 
    appIconA := StrSplit(appIcon, ",")[1], appIconB := StrSplit(appIcon, ",")[2]
Else 
    appIconA := appIcon, appIconB := ""

; Assign custom image to SysTrayIcon and add entry to menu.
TraySetIcon(appIconA, appIconB) ; icon of blue 'info' circle. 
htMenu := A_TrayMenu
htMenu.Add("Start with Windows", StartUpHkTool)
if FileExist(A_Startup "\HotKeyTool.lnk")
    htMenu.Check("Start with Windows")

; a WMI query returns list of processes running from given location.
processlist := ComObject("WbemScripting.SWbemLocator").ConnectServer().ExecQuery("Select Name, ExecutablePath from Win32_Process")

; Get a list of the running exe processes originating from 'ahkFolder' and 
; swap the exe for ahk. Push into array. 
scriptNames := [] 
for process in processlist {
    if (process.ExecutablePath && InStr(process.ExecutablePath, ahkFolder) == 1) {
        processInfo := process.ExecutablePath
        processInfo := StrReplace(processInfo, ".exe", ".ahk")
        scriptNames.Push(processInfo)    
    }
}

; Do preliminary scan of the ahk files and check for '#Included' ahk files. 
; Add those to the array.
for item in scriptNames {
    try loop read item { 
        if SubStr(A_LoopReadLine, 1, 9) = "#Include " {
            if (RegExMatch(A_LoopReadLine, '#Include\s+"([^"]+\.ahk)"', &match)) {
                SplitPath item,, &dir
                scriptNames.Push(dir "\" match[1]) ; Put folder path onto #Included files.
            }
        }
    }
}

If debugMode = 1 {
	for idx, itm1 in scriptNames ; <---------- Leave here for debugging. 
		list1 .= idx ") " itm1 "`n"
	MsgBox "Scripts with running EXEs or #Included in one of them`n`n" list1
}

; Remove "ignore list" items from array. 
for idx, sname in scriptNames {
	for ignItem in ignoreList {
		;sleep 100 ; Occasionally fails to remove an item sometimes. 
		if InStr(sname, ignItem) {
			scriptNames.RemoveAt(idx)
		}
	}
}
; Cull a second time because AHK Toolkit always appears twice in list processes :- /
for idx, sname in scriptNames {
	if InStr(sname, "AHK-ToolKit") {
		scriptNames.RemoveAt(idx)
	}
}

If debugMode = 1 {
	for idx, itm2 in scriptNames ; <---------- Leave here for debugging.
		list2 .= idx ") " itm2 "`n"
	MsgBox "Scripts that we will scan for hotkeys`n`n" list2
}

; Search all the ahk files and collect lines of code that meet criteria
; indicating that a hotkey is present. Push them into 'hotkeys' array.
global hotkeys := []
for item in scriptNames {
    try loop read item { 
        if (InStr(A_LoopReadLine, "::")) {
            colonCount := StrLen(A_LoopReadLine) - StrLen(StrReplace(A_LoopReadLine, ":"))
            if (colonCount == 2 ; Ignore lines with more than 2 colons (I.e. hotstrings)
            && !InStr(A_LoopReadLine, "hide") ; Ignore lines of code with "hide"
            && !InStr(A_LoopReadLine, "'") 
            && !InStr(A_LoopReadLine, '"')) {
                scriptName := SplitPath(item, &name)
                hotkeys.Push(Trim(A_LoopReadLine) " [" name "]") ; Add to hotkey array.
            }
        }
    }
}

; Make array for the combo box list (for quick filtering list).
; Add script names, then array of pre-defined filters. 
global justNames := []
for jname in scriptNames { 
    SplitPath(jname, &jname)
    justNames.Push(jname)
} 
for preFilt in preDefinedFilters {
    justNames.Push(preFilt)
}

; Create the gui.
myKeys := Gui(, guiTitle)
fontCol := fontColor=""? "" : "c" fontColor ; see if custom value for font color set.
myKeys.SetFont("s" fontSize " " fontCol)
myKeys.BackColor := formColor 
WinSetTransparent(trans, myKeys) 
; Text label at top of form. Change as desired.
myKeys.Add("Text", "+wrap w" guiWidth, "Filter then Enter, or Double-Click item.") ; Change label if desired.
rows := (hotkeys.Length <= (maxRows))? hotkeys.Length : maxRows 
; The below combobox will serve as a filter box with some pre-defined filters. 
global hkFilter := myKeys.Add("ComboBox", "w" guiWidth " Background" listColor, justNames) 
myKeys.SetFont("s" fontSize-1)
Global StatBar := myKeys.Add("StatusBar",,) ; Appears at bottom of gui.
myKeys.SetFont("s" fontSize)
global hkList := myKeys.Add("ListView", "w" guiWidth " h" rows*20 " Background" listColor, ["Hotkey", "Action", "Script"])
hkList.ModifyCol(1, (guiWidth // 4)-50)      ; First column width: 1/4 of guiWidth
hkList.ModifyCol(2, (guiWidth // 2)+50)      ; Second column width: 1/2 of guiWidth
hkList.ModifyCol(3, guiWidth // 4)      ; Third column width: 1/4 of guiWidth

populateHkList() ; Call function that adds rows to list.
hkFilter.OnEvent("Change", filterChange) ; if the edit box is changed.
hkList.OnEvent("DoubleClick", runTool) ; if the list item is d-clicked.

; The main hotkey calls this function.  
hotkey(mainHotkey, showMyKeys)
showMyKeys(*)
{   Global targetWindow := WinActive("A")  ; Get the handle of the currently active window
    Global ThisWinTitle := WinGetTitle("ahk_id " targetWindow) ; remember win title so we can wait for it later...
	If debugMode = 1 {
		MsgBox "Target window: " ThisWinTitle
	}
    If WinActive(guiTitle) {
        myKeys.Hide() ; Makes hotkey work like a toggle.
        Return
    }
    Else {
        myKeys.Show("w" guiWidth + 28)
        hkFilter.Focus() ; Focus filter box each time gui is shown. 
    }
}
myKeys.OnEvent("Escape", (*) => myKeys.Hide()) ; Pressing Esc hides form. 

; This function gets called during gui creation.
populateHkList(*) {
    hkList.Delete()  ; Clear the list before populating
    for name in hotkeys {
        parts := StrSplit(name, "::") ; Split line of code.
        hotkey := Trim(parts[1], " `t;") ; Trim ';' from hotkey.
        action := parts[2]
        script := RegExReplace(action, ".*\[(.*)\]$", "$1") ; Part inside [braces] I.e. script file.
        action := Trim(RegExReplace(action, "\s*\[.*\]$", ""), " `t;")
        hkList.Add(, hotkey, action, script) ; Add items to row of list.
    }
    if hotkeys.Length > 0
        hkList.Modify(1, "Select Focus")  ; Pre-select the first item after populating list.
    StatBar.SetText("Showing All of " hotkeys.Length " hotkeys from " ahkFolder "....") ; Update status bar
}

; This function gets called whenever the filter box is updated (from typing in it).
filterChange(*) {
    partialName := hkFilter.Text 
    hkList.Delete()  ; Clear the list before repopulating
    count := 0
    for item in hotkeys {
        if (partialName = "" or InStr(item, partialName, 0)) {
            parts := StrSplit(item, "::")
            hotkey := parts[1]
            action := parts[2]
            script := RegExReplace(action, ".*\[(.*)\]$", "$1")
            action := RegExReplace(action, "\s*\[.*\]$", "")
            hkList.Add(, hotkey, action, script)
            count++
        }
    }
    if (count > 0)
        hkList.Modify(1, "Select Focus")  ; Pre-select the first item if the list is not empty
    StatBar.SetText("Showing " count " of " hotkeys.Length " hotkeys from " ahkFolder "....") ; Update status bar each time.
}

#HotIf WinActive(guiTitle) ; context-sensitive hotkeys
    Enter::runTool()
#HotIf

; This gets called via dbl-clicking list item, or pressing Enter. 
runTool(*) {
    myKeys.Hide()
    If (ThisWinTitle = "") or (ThisWinTitle = "Program Manager") or WinWaitActive(ThisWinTitle,, 4) {
		WinActivate ThisWinTitle
        selectedRow := hkList.GetNext(0, "F")  ; Get the index of the selected row
        if (selectedRow > 0) {
            thisKey := hkList.GetText(selectedRow, 1)  ; Get the text from the first column (Hotkey)
            SendInput thisKey
        }
    }
    else
        MsgBox 'Target window never refocused.'
}

; This function is only accessed via the systray menu item.  It toggles adding/removing
; link to this script in Windows Start up folder.  Uses custom icon too.
StartUpHkTool(*)
{	if FileExist(A_Startup "\HotKeyTool.lnk")
	{	FileDelete(A_Startup "\HotKeyTool.lnk")
		MsgBox("HotKey Tool will NO LONGER auto start with Windows.",, 4096)
	}
	Else 
	{	FileCreateShortcut(A_WorkingDir "\HotKeyTool.exe", A_Startup "\HotKeyTool.lnk", A_WorkingDir, "", "", appIconA, "", appIconB)
		MsgBox("HotKey Tool will auto start with Windows.",, 4096)
	}
    Reload()
}