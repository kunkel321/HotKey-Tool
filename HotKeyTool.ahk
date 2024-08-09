﻿#SingleInstance
#Requires AutoHotkey v2.0

; ==============================================================================
; Title:	    HotKey Lister, Filter'er, and Launcher.
; Author:	    Stephen Kunkel321, with help from Claude.ai
; Version:	    8-9-2024
; GitHub:       https://github.com/kunkel321/HotKey-Tool
; AHK Forum:    https://www.autohotkey.com/boards/viewtopic.php?f=83&t=132224
; ========= INFORMATION ========================================================
; Mostly it's just a "Cheatsheet" list of the hotkeys in your running scripts. 
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
; Added later: Specify a folder which contain link (.lnk) files.  Those files will
; get added to the list of hotkeys.  This tool will either 'Run()' the file, or
; 'Send()' the hotkey. 
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
If !FileExist(SettingsFile) {
    MsgBox "So sorry... Steve forgot to remove the part that reads from his personal custom `"Settings.ini`" file.   Please go to the code and remove this portion, near the top.  (Hint: Probably lines of code " A_LineNumber-10 " through " A_LineNumber+2 ".)  Also be sure to `"uncomment-out`" the three color assignments in the USER OPTIONS.  Now exiting."
    ExitApp
}

; ======= USER OPTIONS =========================================================
; --- Below 3 color assignments should only be commented out for Steve.
; formColor := "00233A" ; Use hex code if desired. Use "Default" for default.
; listColor := "003E67"
; fontColor := "31FFE7"
mainHotkey := "!+q" ; main hotkey to show gui -- Alt+Shift+Q
StickyFilter := 0 ; 0 = Clear the filter box each time the gui is reshown.
guiWidth := 600 ; Width of form. (At least 600 recommended, depending on font size.)
maxRows := 24 ; Scroll if more row than this in listview
guiTitle := "Hotkey Tool (Alt+Shift+Q)" ; Change (here only) if desired. 
fontSize := 12 ; Don't include the 's'.
trans := 255 ; Set 0 (transparent) to 255 (opaque).
appIcon := "shell32.dll,278" ; icon of blue 'info' circle. 
ahkFolder := "D:\AutoHotkey" ; We'll find the active hotkeys in scripts running from here.
lnkFolder := "D:\AutoHotkey\AHK FaveLinks" ; Each .lnk file in this folder is added to list.
ignoreList := ["AHK-ToolKit", "HotstringLib", "QuickSwitch", "mwClipboard"] ; We'll skip these files... 
preDefinedFilters := ["#","!+","!^","Link"] ; Pre-defined filters for the combobox. Add more? 
debugMode := 0 ; 1=On.  Shows processes to be scanned, and target window.
; ==============================================================================

If !DirExist(ahkFolder) { ; Make sure folder is there.
	MsgBox "The folder `"" ahkFolder "`" doesn't seem to exist.  Please point the variable `"ahkFolder :=`" to the folder where your scripts are located. Now exiting"
	ExitApp
}

if !DirExist(lnkFolder) { ; Make sure folder is there.
    MsgBox "The folder `"" lnkFolder "`" doesn't exist. Please check the lnkFolder variable.  It should point to a folder full of .LNK files.  Those files should point to .EXE files. "
    ExitApp
}

; Determine if icon is .ico or library item, parse as needed.
If InStr(appIcon, ",") 
    appIconA := StrSplit(appIcon, ",")[1], appIconB := StrSplit(appIcon, ",")[2]
Else 
    appIconA := appIcon, appIconB := ""

; Assign custom image to SysTrayIcon and add entry to menu.
TraySetIcon(appIconA, appIconB)
htMenu := A_TrayMenu
htMenu.Add("Start with Windows", (*) => StartUpHkTool(A_WorkingDir, appIconA, appIconB))
if FileExist(A_Startup "\HotKeyTool.lnk")
    htMenu.Check("Start with Windows")

; Function to get script names
GetScriptNames(ahkFolder, ignoreList) {
    processlist := ComObject("WbemScripting.SWbemLocator").ConnectServer().ExecQuery("Select Name, ExecutablePath from Win32_Process")
    
    scriptNames := []
    for process in processlist {
        if (process.ExecutablePath && InStr(process.ExecutablePath, ahkFolder) == 1) {
            processInfo := StrReplace(process.ExecutablePath, ".exe", ".ahk")
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

    ; Remove "ignore list" items from array. 
    filteredScriptNames := []
    for sname in scriptNames {
        shouldInclude := true
        for ignItem in ignoreList {
            if InStr(sname, ignItem) {
                shouldInclude := false
                break
            }
        }
        if shouldInclude {
            filteredScriptNames.Push(sname)
        }
    }

    ; Cull a second time because AHK Toolkit always appears twice in list processes :- /
    finalScriptNames := []
    for sname in filteredScriptNames {
        if !InStr(sname, "AHK-ToolKit") {
            finalScriptNames.Push(sname)
        }
    }

    return finalScriptNames
}

scriptNames := GetScriptNames(ahkFolder, ignoreList)

If debugMode = 1 {
	for idx, itm2 in scriptNames ; <---------- Leave here for debugging.
		list2 .= idx ") " itm2 "`n"
	MsgBox "Scripts that we will scan for hotkeys`n`n" list2
}

; Function to get hotkeys from scripts
GetHotkeys(scriptNames) {
    hotkeys := []
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
    return hotkeys
}

hotkeys := GetHotkeys(scriptNames)

;Function to get the target of a shortcut file and add the shortcut names to the
; array of hotkeys so they show up in the list box. 
GetLinks() {
    GetShortcutTarget(FilePath) {
        objShell := ComObject("WScript.Shell")
        objLink := objShell.CreateShortcut(FilePath)
        return objLink.TargetPath
    }

    linkHotkeys := []
    Loop Files, lnkFolder "\*.lnk"
    {
        target := GetShortcutTarget(A_LoopFilePath)
        linkHotkeys.Push("Link::" A_LoopFileName " [" target "]")
    }

    return linkHotkeys
}

linkHotkeys := GetLinks()
hotkeys.Push(linkHotkeys*)

; Function to get combo box list
GetComboBoxList(scriptNames, preDefinedFilters) {
    justNames := []
    for jname in scriptNames { 
        SplitPath(jname, &jname)
        justNames.Push(jname)
    } 
    for preFilt in preDefinedFilters {
        justNames.Push(preFilt)
    }
    return justNames
}

justNames := GetComboBoxList(scriptNames, preDefinedFilters)

; Create the GUI
myKeys := CreateGui(guiTitle, guiWidth, formColor, listColor, fontColor, fontSize, trans, maxRows, justNames, hotkeys, ahkFolder)

; The main hotkey calls this function.  
Hotkey mainHotkey, ShowMyKeysWrapper

ShowMyKeysWrapper(*)
{
    showMyKeys(myKeys, guiTitle, guiWidth)
}

; Function to create GUI
CreateGui(guiTitle, guiWidth, formColor, listColor, fontColor, fontSize, trans, maxRows, justNames, hotkeys, ahkFolder) {
    myKeys := Gui(, guiTitle)
    fontCol := fontColor=""? "" : "c" fontColor ; see if custom value for font color set.
    myKeys.SetFont("s" fontSize " " fontCol)
    myKeys.BackColor := formColor 
    WinSetTransparent(trans, myKeys) 
    ; Text label at top of form. Change as desired.
    myKeys.Add("Text", "+wrap w" guiWidth, "Filter then Enter, or Double-Click item.") ; Change label if desired.
    rows := (hotkeys.Length <= (maxRows))? hotkeys.Length : maxRows 
    ; The below combobox will serve as a filter box with some pre-defined filters. 
    myKeys.hkFilter := myKeys.Add("ComboBox", "w" guiWidth " Background" listColor, justNames) 
    myKeys.SetFont("s" fontSize-1)
    myKeys.StatBar := myKeys.Add("StatusBar",,) ; Appears at bottom of gui.
    myKeys.SetFont("s" fontSize)
    myKeys.hkList := myKeys.Add("ListView", "w" guiWidth " h" rows*20 " Background" listColor, ["Hotkey", "Action", "Script"])
    myKeys.hkList.ModifyCol(1, (guiWidth // 4)-50)      ; First column width: 1/4 of guiWidth
    myKeys.hkList.ModifyCol(2, (guiWidth // 2)+50)      ; Second column width: 1/2 of guiWidth
    myKeys.hkList.ModifyCol(3, guiWidth // 4)           ; Third column width: 1/4 of guiWidth

    updateHkList(myKeys.hkList, hotkeys, myKeys.StatBar, ahkFolder, "") ; Call function that adds rows to list.
    myKeys.hkFilter.OnEvent("Change", (*) => filterChange(myKeys.hkFilter, myKeys.hkList, hotkeys, myKeys.StatBar, ahkFolder)) ; if the edit box is changed.
    myKeys.hkList.OnEvent("DoubleClick", (*) => runTool(myKeys.hkList, myKeys)) ; if the list item is d-clicked.

    myKeys.OnEvent("Escape", (*) => myKeys.Hide()) ; Pressing Esc hides form. 

    return myKeys
}

; The hotkeys and link files get added to the array when the script starts, but this
; function is each time the gui is re-shown.  
showMyKeys(myKeys, guiTitle, guiWidth) {
    global targetWindow := WinActive("A") ; Get the handle of the currently active window
    global ThisWinTitle := WinGetTitle("ahk_id " targetWindow) ; remember win title so we can wait for it later...
    If debugMode = 1 {
        MsgBox "Target window: " ThisWinTitle
    }
    If WinActive(guiTitle) { ; Causes fgui to hid, if already showing (I.e. toggle))
        myKeys.Hide()
        global targetWindow := ""
        Return
    }
    Else {
        If StickyFilter = 0 {
            myKeys.hkFilter.Text := ""
            filterChange(myKeys.hkFilter, myKeys.hkList, hotkeys, myKeys.StatBar, ahkFolder)
        }
        myKeys.Show("w" guiWidth + 28)
        global targetWindow := ""
        myKeys.hkFilter.Focus() ; Focus filter box each time gui is shown. 
    }
}

; Combined function for populating and filtering the list
updateHkList(hkList, hotkeys, StatBar, ahkFolder, filter := "") {
    hkList.Delete()  ; Clear the list before populating
    count := 0
    for item in hotkeys {
        if (filter = "" or InStr(item, filter, 0)) {
            parts := StrSplit(item, "::")
            hotkey := Trim(parts[1], " `t;") ; Trim ';' from hotkey.
            action := parts[2]
            script := RegExReplace(action, ".*\[(.*)\]$", "$1")
            action := Trim(RegExReplace(action, "\s*\[.*\]$", ""), " `t;")
            hkList.Add(, hotkey, action, script)
            count++
        }
    }
    if (count > 0)
        hkList.Modify(1, "Select Focus")  ; Pre-select the first item if the list is not empty
    StatBar.SetText("Showing " count " of " hotkeys.Length " hotkeys from " ahkFolder "....") ; Update status bar
}

; This function gets called whenever the filter box is updated (from typing in it).
filterChange(hkFilter, hkList, hotkeys, StatBar, ahkFolder) {
    updateHkList(hkList, hotkeys, StatBar, ahkFolder, hkFilter.Text)
}

; #HotIf WinActive(guiTitle) ; context-sensitive hotkeys
#HotIf (myKeys.FocusedCtrl == myKeys.hkFilter) or (myKeys.FocusedCtrl == myKeys.hkList) ; context-sensitive hotkeys
    Enter::pressedEnter()
#HotIf

pressedEnter(*) {
    runTool(myKeys.hkList, myKeys)
}

; This gets called via dbl-clicking list item, or pressing Enter. 
runTool(hkList, myKeys) {
    myKeys.Hide()
    If (ThisWinTitle = "") or (ThisWinTitle = "Program Manager") or WinWaitActive(ThisWinTitle,, 3) {
		WinActivate ThisWinTitle
        selectedRow := hkList.GetNext(0, "F")  ; Get the index of the selected row
        if (selectedRow > 0) {
            thisKey := hkList.GetText(selectedRow, 1)  ; Get the text from the first column (Hotkey)
            If SubStr(thisKey, 1, 4) = "Link"{
                thisLink := hkList.GetText(selectedRow, 3) ; get third column, which is the taget path
                Run(thisLink)
            }
            Else 
                SendInput thisKey
        }
    }
    else
        MsgBox "Target window (" ThisWinTitle ") never refocused."
}

; This function is only accessed via the systray menu item.  It toggles adding/removing
; link to this script in Windows Start up folder.  Uses custom icon too.
StartUpHkTool(workingDir, iconA, iconB)
{	if FileExist(A_Startup "\HotKeyTool.lnk")
	{	FileDelete(A_Startup "\HotKeyTool.lnk")
		MsgBox("HotKey Tool will NO LONGER auto start with Windows.",, 4096)
	}
	Else 
	{	FileCreateShortcut(workingDir "\HotKeyTool.exe", A_Startup "\HotKeyTool.lnk", workingDir, "", "", iconA, "", iconB)
		MsgBox("HotKey Tool will auto start with Windows.",, 4096)
	}
    Reload()
}
