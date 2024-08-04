# HotKey Tool
 A tool to list, filter, and launch hotkeys from running scripts. 
 This tool assumes that the associated (same-named) ahk file is in the same folder as the exe file that is currently running. It won't work if you are running compiled scripts without the corresponding .ahk file. It is the ahk files that actually get scaned.

The tool will look running processes from the folder you specify (and subfolders thereof) and scan the same-named .ahk files. It will do a preliminary scan for #Included .ahk files, then scan all of those for hotkeys. See code for additional comments and user options.

FROM CODE COMMENTS
* ==============================================================================
* Title:	    HotKey Lister, Filter'er, and Launcher.
* Author:	    Stephen Kunkel321, with help from Claude.ai
* Version:	    8-4-2024
* GitHub:       https://github.com/kunkel321/HotKey-Tool
* AHK Forum:    https://www.autohotkey.com/boards/viewtopic.php?f=83&t=132224
* ========= INFORMATION ========================================================
* Mostly it's just a "Cheetsheet" list of the hotkeys in your running scripts. 
* Also can be used to launch them though.  Launches via "Sending" the hotkey.
* Double-click an item to launch it.  Enter key launches selected item too.
* Made for "portably-running" scripts.  Those are scripts where a copy of AutoHotkey.exe has been put in the same folder as the .ahk file and renamed to match it. For example myScript.ahk and myScript.exe are both in the same folder. 
* If the exe is a compiled version of the ahk, then that should work too. 
* It is actually the ahk file that gets searched for hotkeys.
* Also the folder has to be a subfolder of the one specified via ahkFolder variable.
* Esc only hides form.  Must restart script to rescan hotkeys.
* There's no point scanning files with no hotkeys, so add those to ignore list array.
* Set your scripts up such that there is an in-line comment on each line of code
* that has a hotkey.  Use descriptive "search terms" in the comment.
* To hide individual hotkeys from scan, include the word "hide" (without quotes)in the comment.  Lines of code are ignored if they...
* - have the word 'hide'.
* - have single or double quotes.
* - don't contain "::".
* - do contain more than two colons.
* The Filter Box is a ComboBox and can have pre-defined hotkey search filters. 
* - Names of scripts are added automatically, to filter by containing script file.
* - Additional filters can be added to array in USER OPTIONS below.
* There are a few other options below as well.  See also, copious in-line comments.
* Tool will determine active window then wait for it before sending hotkey.
* ==============================================================================
![Screenshot of Hotkey Tool main form](https://i.imgur.com/GgTuK1l.png)
