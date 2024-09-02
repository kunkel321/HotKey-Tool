#SingleInstance
#Requires AutoHotkey v2.0

; ==============================================================================
; Title:	    HotKey-Tool:  A Lister, Filter'er, and Launcher.
; Author:	    Stephen Kunkel321
; Tools:        Claude.ai was used extensively.  AI-generated code is indicated below. 
; Version:	    9-2-2024
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

; ======= USER OPTIONS =========================================================
; --- Below 3 color assignments should only be commented out for Steve.
formColor     := "00233A" ; Use hex code if desired. Use "Default" for default.
listColor     := "003E67"
fontColor     := "31FFE7"
mainHotkey      := "!+q"    ; main hotkey to show gui -- Alt+Shift+Q
StickyFilter    := 0        ; 0 = Clear the filter box each time the gui is reshown.
guiWidth        := 600      ; Width of form. (At least 600 recommended, depending on font size.)
maxRows         := 24       ; Scroll if more row than this in listview
guiTitle        := A_UserName "'s Hotkey Tool" ; Change (here only) if desired. 
fontSize        := 12       ; Don't include the 's'.
trans           := 255      ; Set 0 (transparent) to 255 (opaque).
appIcon         := "comctl32.dll,3" ; icon of 'info' word bubble. 
ahkFolder       := "D:\AutoHotkey" ; We'll find (recursively) the active hotkeys in scripts running from here.
lnkFolders      := [ ; Each .lnk file in this folder is added to list.
                    "D:\AutoHotkey\AHK FaveLinks", ; <-------------------------- Specific to Steve's computer! 
                    "C:\Users\" A_UserName "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs",
                    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
                ]
lnkFoldersIsRecursive := 1  ; 1=Search the subfolders from lnkFolders too.
ignoreList      := ["AHK-ToolKit", "HotstringLib", "QuickSwitch", "mwClipboard"] ; We'll skip these script files... 
preDefinedFilters := ["#","!+","!^","Link"] ; Pre-defined filters for the combobox. Add more? 
debugMode       := 0 ;       1=On.  Shows processes to be scanned, and target window.
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
TraySetIcon(appIconA, appIconB)
htMenu := A_TrayMenu
htMenu.Add("Start with Windows", (*) => StartUpHkTool(A_WorkingDir, appIconA, appIconB))
if FileExist(A_Startup "\HotKeyTool.lnk")
    htMenu.Check("Start with Windows")

; Function to get script names
GetScriptNames(ahkFolder, ignoreList) {
    ; Claude.ai worked out the ComObject, and how to filter for processes originating from a gived folder.
    processlist := ComObject("WbemScripting.SWbemLocator").ConnectServer().ExecQuery("Select Name, ExecutablePath from Win32_Process")
    
    
    scriptNames := []
    for process in processlist {
        if (process.ExecutablePath && InStr(process.ExecutablePath, ahkFolder) == 1) {
            processInfo := StrReplace(process.ExecutablePath, ".exe", ".ahk")
            scriptNames.Push(processInfo)    
        }
    }

    If debugMode = 1 {
        for idx, itm1 in scriptNames ; <---------- Leave here for debugging.
            list1 .= idx ") " itm1 "`n"
        MsgBox "Processes found:`n`n" list1
    }
    ; Do preliminary scan of the ahk files and check for '#Included' ahk files. 
    ; Add those to the array.
    for item in scriptNames {
        try loop read item { 
            if SubStr(A_LoopReadLine, 1, 9) = "#Include " {
                ; Claude.ai wrote this regex.
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
                && !InStr(A_LoopReadLine, "hide")   ; Ignore lines of code with "hide"
                && !InStr(A_LoopReadLine, "'")      ; Ignore lines of code with single-quotes
                && !InStr(A_LoopReadLine, '"')) {   ; Ignore lines of code with double-quotes
                    scriptName := SplitPath(item, &name)
                    hotkeys.Push(Trim(A_LoopReadLine) " [" name "]") ; Add to hotkey array.
                }
            }
        }
    }
    return hotkeys
}

hotkeys := GetHotkeys(scriptNames)

; Function to get the target of a shortcut file and add the shortcut names to the
; array of hotkeys so they show up in the list box. Claude.ai worked-out how to write
; and use this ComObject to get the link targets. 
GetLinks() {
    GetShortcutTarget(FilePath) {
        objShell := ComObject("WScript.Shell")
        objLink := objShell.CreateShortcut(FilePath)
        return objLink.TargetPath
    }

    linkHotkeys := []
    recurse := (lnkFoldersIsRecursive=1)? "R": "" ; Search in subfolders too?
    for folder in lnkFolders {
        if !DirExist(folder) { ; Make sure folder is there.
            MsgBox "The folder `"" folder "`" doesn't exist. Please check the lnkFolders array elements.  It should point to folders with .LNK files.`n`n`nNow exiting."
            ExitApp
        }
        Loop Files, folder "\*.lnk", recurse {
            target := GetShortcutTarget(A_LoopFilePath)
            linkHotkeys.Push("Link::" A_LoopFileName " [" target "]")
        }
    }
    ;Dedupe links.
    uniqueArr := []
    for item in linkHotkeys {
            isDuplicate := false
            for uniqueItem in uniqueArr {
                if (item = uniqueItem) {
                    isDuplicate := true
                    break
                }
            }
            if (!isDuplicate)
                uniqueArr.Push(item)
        }
    linkHotkeys := uniqueArr
    return linkHotkeys
}

linkHotkeys := GetLinks()
hotkeys.Push(linkHotkeys*)

; Function to get combobox list
GetComboBoxList(scriptNames, preDefinedFilters) {
    justNames := []
    justNames.Push("") ; Add blank to top of list.
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

; The main hotkey calls this function.  It was Claude.ai that incorporated the "wrapper." 
; this was done in response to an error I was getting.  (I don't remember the error.)
Hotkey mainHotkey, ShowMyKeysWrapper

ShowMyKeysWrapper(*) {
    showMyKeys(myKeys, guiTitle, guiWidth)
}

; Function to create GUI.  I (the human) had made a normal simple gui that had a ListBox
; control for the list of hotkeys.  I asked Claude.ai to change the ListBox into a 
; ListView.  I wrote a long narrative explaining what to make.  The AI created this function
; with all the paramaters, and seems to have used "dot.notation" for creating the gui controls
; which I found interesting. I tweaked parts, but mostly this function is "all Claude." 
CreateGui(guiTitle, guiWidth, formColor, listColor, fontColor, fontSize, trans, maxRows, justNames, hotkeys, ahkFolder) {
    myKeys := Gui(, guiTitle)
    fontCol := fontColor=""? "" : "c" fontColor ; see if custom value for font color set.
    myKeys.SetFont("s" fontSize " " fontCol)
    myKeys.BackColor := formColor 
    WinSetTransparent(trans, myKeys) 
    ; Text label at top of form. Change as desired.
    global txtLbl := myKeys.Add("Text", "+wrap w" guiWidth, "Filter then Enter, or Double-Click item.  Target: ") ; Change label if desired.
    rows := (hotkeys.Length <= (maxRows))? hotkeys.Length : maxRows 
    ; The below combobox will serve as a filter box with some pre-defined filters. 
    myKeys.hkFilter := myKeys.Add("ComboBox", "w" guiWidth " Background" listColor, justNames) 
    myKeys.SetFont("s" fontSize-1)
    myKeys.StatBar := myKeys.Add("StatusBar",,) ; Appears at bottom of gui.
    myKeys.SetFont("s" fontSize)
    myKeys.hkList := myKeys.Add("ListView", "w" guiWidth " h" rows*20 " Background" listColor, ["Hotkey", "Action", "Script"])
    myKeys.hkList.ModifyCol(1, (guiWidth // 4)-50)      ; First column width: 1/4 of guiWidth
    myKeys.hkList.ModifyCol(2, (guiWidth // 2)+50)      ; Second column width: 1/2 of guiWidth
    myKeys.hkList.ModifyCol(3, guiWidth)           ; Third column width: 1/4 of guiWidth
    ;myKeys.hkList.ModifyCol(3, guiWidth // 4)           ; Third column width: 1/4 of guiWidth

    updateHkList(myKeys.hkList, hotkeys, myKeys.StatBar, ahkFolder, "") ; Call function that adds rows to list.
    myKeys.hkFilter.OnEvent("Change", (*) => filterChange(myKeys.hkFilter, myKeys.hkList, hotkeys, myKeys.StatBar, ahkFolder)) ; if the edit box is changed.
    myKeys.hkList.OnEvent("DoubleClick", (*) => runTool(myKeys.hkList, myKeys)) ; if the list item is d-clicked.

    myKeys.OnEvent("Escape", (*) => myKeys.Hide()) ; Pressing Esc hides form. 

    return myKeys
}

; The hotkeys and link files get added to the array when the script starts, but this
; function is each time the gui is re-shown.  
showMyKeys(myKeys, guiTitle, guiWidth) {
    global targetWindow := WinActive("A")   ; Get the handle of the currently active window
    global ThisWinTitle := ""
    try ThisWinTitle := WinGetTitle("ahk_id " targetWindow) ; remember win title so we can wait for it later...
    If debugMode = 1 {
        MsgBox "Target window: " ThisWinTitle
    }
    If WinActive(guiTitle) {                ; Causes gui to hide, if already showing (I.e. toggle))
        myKeys.Hide()
        global targetWindow := ""
        Return
    }
    Else {
        If StickyFilter = 0 {
            myKeys.hkFilter.Text := ""
            filterChange(myKeys.hkFilter, myKeys.hkList, hotkeys, myKeys.StatBar, ahkFolder)
        }
        txtLbl.Text := "Filter then Enter, or Double-Click item.  Target: " ThisWinTitle
        myKeys.Show("w" guiWidth + 28)      ; Increase by 28pix, so there's a margin around the controls.
        global targetWindow := ""
        myKeys.hkFilter.Focus()             ; Focus filter box each time gui is shown. 
    }
}

; Combined function for populating and filtering the list.
; I had written this function to push a single line of text onto the array, for use
; in a ListBox.  When I prompted the AI to replace the ListBox with a ListView, it
; also updated this function. 
updateHkList(hkList, hotkeys, StatBar, ahkFolder, filter := "") {

    static il := IL_Create()

    static Icons := Map(
        '.ahk', IL_Add(il,"C:\Program Files\Autohotkey\v2\Autohotkey64.exe")
    )

    if Icons.Count = 1
        hkList.SetImageList(il)
    
    hkList.Delete()  ; Clear the list before populating
    count := 0
    for item in hotkeys {
        if (filter = "" or InStr(item, filter, 0)) {
            parts := StrSplit(item, "::")
            hotkey := Trim(parts[1], " `t;") ; Trim ';' from hotkey.
            action := parts[2]
            script := RegExReplace(action, ".*\[(.*)\]$", "$1")
            action := Trim(RegExReplace(action, "\s*\[.*\]$", ""), " `t;")

            ; RaptorX (a human), with The Automator group added this icon part during an AHK Hero zoom call. 
            SplitPath script,,,&Ext

            if Ext = 'ahk'
                i_indx := Icons['.ahk']
            else if !Icons.Has(script)
                i_indx := Icons[script] := IL_Add(il, script)
            else
                i_indx := Icons[script]

            hkList.Add('Icon' i_indx, hotkey, action, script)
            count++
        }
    }
    if (count > 0)
        hkList.Modify(1, "Select Focus")  ; Pre-select the first item if the list is not empty
    StatBar.SetText("Showing " count " of " hotkeys.Length " items.") ; Update status bar
}

; This function gets called whenever the filter box is updated (from typing in it).
filterChange(hkFilter, hkList, hotkeys, StatBar, ahkFolder) {
    updateHkList(hkList, hotkeys, StatBar, ahkFolder, hkFilter.Text)
}

; #HotIf WinActive(guiTitle) ; context-sensitive hotkeys
#HotIf (myKeys.FocusedCtrl == myKeys.hkFilter) or (myKeys.FocusedCtrl == myKeys.hkList) ; context-sensitive hotkeys
    Enter::pressedEnter() ; hide
#HotIf

pressedEnter(*) {
    runTool(myKeys.hkList, myKeys)
}

; This gets called via dbl-clicking list item, or pressing Enter. 
runTool(hkList, myKeys) {
    myKeys.Hide()
    If (ThisWinTitle = "") or (ThisWinTitle = "Program Manager") or WinWaitActive(ThisWinTitle,, 3) {
        selectedRow := hkList.GetNext(0, "F")               ; Get the index of the selected row
        if (selectedRow > 0) {
            thisKey := hkList.GetText(selectedRow, 1)       ; Get the text from the first column (Hotkey)
            If SubStr(thisKey, 1, 4) = "Link"{
                thisLink := hkList.GetText(selectedRow, 3)  ; get third column, which is the taget path
                If FileExist(thisLink)
                    Run(thisLink)
                else
                    MsgBox "The file `"" thisLink "`" doesn't appear to exist."
            }
            Else {
                If RegExMatch(thisKey, "i).*?[a-z]{2,}") ; If hotkey has word like "space" then wrap it in braces like {space}
                    thisKey := RegExReplace(thisKey, "i)(.*?)([a-z]{2,})", "$1{$2}")
                SendInput thisKey
            }
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
