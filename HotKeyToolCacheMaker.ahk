; Extracts icons from executables targeted by .lnk files and saves them as PNGs
; Saves them in folder "iconCache" which is read by HotkeyTool app.  
; HotkeyTool must be in same folder as this file (and the iconCache folder).
; Made by kunkel321 using ClaudeAI.  Version:  12-23-2024

#SingleInstance
#Requires AutoHotkey v2.0

; ========== USER OPTIONS ======================================================
; Link folders to scan - copied from HotKeyTool code, near top.
lnkFolders := [
    "D:\PortableApps\FavePortableLinks",    ; <-------------------------- Specific to Steve's computer! 
    "D:\AutoHotkey\AHK FaveLinks",          ; <-------------------------- Specific to Steve's computer! 
    "C:\Users\" A_UserName "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
]
lnkFoldersIsRecursive := 1  ; Search subfolders
hideUninstallers := 1       ; Skip uninstaller links
; ==============================================================================

; Check for AutoHotkey first
ahkPath := "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
if !FileExist(ahkPath) {
    ahkPath := FileSelect(3,, "Select AutoHotkey Executable", "Executable (*.exe)")
    if !ahkPath {
        MsgBox "AutoHotkey executable is required."
        ExitApp
    }
}

; Initialize GDI+
if !(token := Gdip_Startup()) {
    MsgBox "Failed to start GDI+."
    ExitApp
}

; Extract AHK icon first
outputFolder := A_ScriptDir "\IconCache"
if !DirExist(outputFolder)
    DirCreate(outputFolder)

ahkIconPath := outputFolder "\Program Files-AutoHotkey-v2-AutoHotkey64.png"
if !FileExist(ahkIconPath) {
    if ExtractExeIcon(ahkPath, ahkIconPath)
        FileAppend("Extracted: " ahkPath "`n", "*")
}

; Embedded Gdip functions
Gdip_Startup() {
    if (!DllCall("LoadLibrary", "str", "gdiplus", "UPtr")) {
        throw Error("Could not load GDI+ library")
    }

    si := Buffer(A_PtrSize = 4 ? 20:32, 0)
    NumPut("uint", 0x2, si)
    NumPut("uint", 0x4, si, A_PtrSize = 4 ? 16:24)
    DllCall("gdiplus\GdiplusStartup", "UPtr*", &pToken:=0, "Ptr", si, "UPtr", 0)

    return pToken
}

Gdip_Shutdown(pToken) {
    DllCall("gdiplus\GdiplusShutdown", "UPtr", pToken)
    if hModule := DllCall("GetModuleHandle", "str", "gdiplus", "UPtr")
        DllCall("FreeLibrary", "UPtr", hModule)
    return 0
}

Gdip_CreateBitmapFromHICON(hIcon) {
    DllCall("gdiplus\GdipCreateBitmapFromHICON", "UPtr", hIcon, "UPtr*", &pBitmap:=0)
    return pBitmap
}

Gdip_SaveBitmapToFile(pBitmap, sOutput, Quality:=75) {
    _p := 0

    SplitPath sOutput,,, &extension:=""
    if (!RegExMatch(extension, "^(?i:BMP|DIB|RLE|JPG|JPEG|JPE|JFIF|GIF|TIF|TIFF|PNG)$")) {
        return -1
    }
    extension := "." extension

    DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", &nCount:=0, "uint*", &nSize:=0)
    ci := Buffer(nSize)
    DllCall("gdiplus\GdipGetImageEncoders", "UInt", nCount, "UInt", nSize, "UPtr", ci.Ptr)

    loop nCount {
        address := NumGet(ci, (idx := (48+7*A_PtrSize)*(A_Index-1))+32+3*A_PtrSize, "UPtr")
        sString := StrGet(address, "UTF-16")
        if !InStr(sString, "*" extension)
            continue

        pCodec := ci.Ptr+idx
        break
    }

    if (!pCodec)
        return -3

    if (Quality != 75) {
        Quality := (Quality < 0) ? 0 : (Quality > 100) ? 100 : Quality
        if RegExMatch(extension, "^\.(?i:JPG|JPEG|JPE|JFIF)$") {
            DllCall("gdiplus\GdipGetEncoderParameterListSize", "UPtr", pBitmap, "UPtr", pCodec, "uint*", &nSize)
            EncoderParameters := Buffer(nSize, 0)
            DllCall("gdiplus\GdipGetEncoderParameterList", "UPtr", pBitmap, "UPtr", pCodec, "UInt", nSize, "UPtr", EncoderParameters.Ptr)
            nCount := NumGet(EncoderParameters, "UInt")
            loop nCount
            {
                elem := (24+(A_PtrSize ? A_PtrSize : 4))*(A_Index-1) + 4 + (pad := A_PtrSize = 8 ? 4 : 0)
                if (NumGet(EncoderParameters, elem+16, "UInt") = 1) && (NumGet(EncoderParameters, elem+20, "UInt") = 6)
                {
                    _p := elem + EncoderParameters.Ptr - pad - 4
                    NumPut("UInt", Quality, NumGet(NumPut("UInt", 4, NumPut("UInt", 1, _p+0)+20), "UInt"))
                    break
                }
            }
        }
    }

    return DllCall("gdiplus\GdipSaveImageToFile"
                    , "UPtr", pBitmap
                    , "UPtr", StrPtr(sOutput)
                    , "UPtr", pCodec
                    , "UInt", _p ? _p : 0)
}

Gdip_DisposeImage(pBitmap) {
    return DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
}

; Function to get shortcut target (from HotKeyTool)
GetShortcutTarget(FilePath) {
    objShell := ComObject("WScript.Shell")
    objLink := objShell.CreateShortcut(FilePath)
    return objLink.TargetPath
}

; Extract icon from exe and save as PNG
ExtractExeIcon(exePath, savePath) {
    hIcon := 0
    iconCount := DllCall("Shell32\ExtractIconExW", "Str", exePath, "Int", 0, "UPtr*", &hIcon, "UPtr*", 0, "UInt", 1)
    
    if (!hIcon)
        return false

    if !(pBitmap := Gdip_CreateBitmapFromHICON(hIcon)) {
        DllCall("DestroyIcon", "UPtr", hIcon)
        return false
    }

    result := Gdip_SaveBitmapToFile(pBitmap, savePath)

    DllCall("DestroyIcon", "UPtr", hIcon)
    Gdip_DisposeImage(pBitmap)
    
    return result = 0
}

; Process links in a folder
ProcessLinkFolder(folder, recurse, &processedCount, &skipCount) {
    outputFolder := A_ScriptDir "\IconCache"
    if !DirExist(outputFolder)
        DirCreate(outputFolder)

    loop files folder "\*.lnk", (recurse ? "R" : "") {
        ; Skip uninstallers if configured
        if hideUninstallers && (InStr(A_LoopFileName, "Uninstall ") 
            || InStr(A_LoopFileName, "Uninstall.")
            || InStr(A_LoopFileName, "Uninstall)")) {
            skipCount++
            continue
        }

        try {
            targetPath := GetShortcutTarget(A_LoopFileFullPath)
            if (!targetPath || !FileExist(targetPath)) {
                skipCount++
                continue
            }

            ; Create PNG filename from target path
            pngName := SubStr(targetPath, 3)  ; Remove drive letter
            pngName := StrReplace(pngName, "\", "-")
            pngName := RegExReplace(pngName, "\.exe$|\.EXE$", "")  ; Remove .exe or .EXE
            pngName := RegExReplace(pngName, "^-+", "")  ; Remove leading hyphens
            pngPath := outputFolder "\" pngName ".png"

            if (!FileExist(pngPath) && ExtractExeIcon(targetPath, pngPath)) {
                processedCount++
                FileAppend "Extracted: " targetPath "`n", "*"
            } else {
                skipCount++
            }
        } catch {
            skipCount++
        }
    }
}

; Main GUI setup
MainGui := Gui()
MainGui.OnEvent("Close", (*) => ExitApp())
MainGui.Add("Text",, "Extracting icons from shortcut target executables...")
progressText := MainGui.Add("Text", "w400", "Starting...")
MainGui.Show()

; Initialize GDI+
if !(token := Gdip_Startup()) {
    MsgBox "Failed to start GDI+."
    ExitApp
}

processedCount := 0
skipCount := 0

; Process each link folder
for folder in lnkFolders {
    if !DirExist(folder) {
        progressText.Value := "Skipping non-existent folder: " folder
        continue
    }
    
    progressText.Value := "Processing: " folder
    ProcessLinkFolder(folder, lnkFoldersIsRecursive, &processedCount, &skipCount)
}

Gdip_Shutdown(token)

MainGui.Destroy()
MsgBox "Icon extraction complete!`n`n" 
    . "Icons processed: " processedCount "`n"
    . "Items skipped: " skipCount "`n`n"
    . "Icons saved to: " A_ScriptDir "\IconCache"

ExitApp