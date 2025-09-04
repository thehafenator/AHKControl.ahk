#Requires AutoHotkey v2.0
#SingleInstance Force

/*
Readme:
Last updated 01.20.2025
AHKControl for AutoHotkey Version 2 by thehafenator. This is a 'port' from Lexicos' (The AutoHotkey Developer) AHKControl written in V1, and contains much of the functionality. Though not a true port, I was able to bring most of the functionality. 
See the V1 script on github here: https://github.com/Lexikos/AHKControl.ahk

Function of this V2 code:
1. View all AHK scripts, V1 or V2, in a single tray icon.
2. View each AHK script's original TrayMenu through the menu creation.
3. Send messages to reload, edit, suspend, pause, or edit individual scripts.
4. The 'Edit' command will show a list of all of the #include files in an AHK script, allowing each to be accessed.
5. An editor path (the variable is 'editorpath') can be specified to allow the scripts to be opened with 'edit' in an editor of choice. If the editor path does not work, it will default to opening with notepad. 
6. Pause/Suspend checkmarks. The script will show checkmarks under each submenu if the respective script is paused or suspended. It does this in three ways:
a. On startup or reload, it runs a wmi query to check the suspended/paused status of each script. 
b. By default, the script updates every 15 seconds
c. If a script is paused through an AHKControl submenu, it keeps an internal map called scriptstates(). This keeps the checkmarks more accurate until the timer refreshes sets the script state again.
Though not perfect, the only way that the pause/suspend checkmarks will be innaccurate will be if the user suspends a script through that script's own icon. However, it will only remain out of sync for 15 seconds or if the user reloads AHKControl, whichever comes first.
7. Apply a different icon for light and dark theme in Windows 10 and 11. The script reads a registry value to determine if it is in dark theme, and will try to set the theme depending on which value it reads, and have a back in case it doesn't work. If all do not work, the script will proceed with the default AHK icon.
8. Dark Mode Win 32 menus. With contextcolor(), which was provided on a forum response post by Lexicos to another user.


Funcationality that I would like to add, but am not sure how to implement:
1. Adding .exe/compiled AHK scripts into the GetRunningAHKScripts(). I haven't been able to figure this one out. If anyone wants to help, that would be great. 

Differences that that others may be able to tackle:
1. Grab icons from the scripts themselves. The V1 script was able to find and identify the icons within each script itself, even when #Notrayicon was used. I personally prefer the no icon look and think it looks cleaner.
2. The Scripts were numbered on Lexicos' version within A_TrayMenu, though I haven't been able how to achieve this and keep the paths clean throughout the script. 
I also prefer this look without the numbering.

*/

#Requires AutoHotkey v2.0
#SingleInstance Force

; try {
;     isDarkMode := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme") = 0
;     if (isDarkMode)
;         TraySetIcon A_ScriptDir "\Macropad\Icons\m3dark.ico" ; if in dark theme, try to set a dark theme icon
;     else
;         TraySetIcon A_ScriptDir "\Macropad\Icons\mdark.ico" ; otherwise, set to set a light theme icon
; } catch {
;     try {
;         TraySetIcon A_ScriptDir "\Macropad\Icons\3mdark.ico" ; if both light/dark theme icons don't work, fall back on this.
;     }
; }


try {
    isDarkMode := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme") = 0
    if (isDarkMode)
        TraySetIcon A_ScriptDir "\Icons\macropaddark.ico"
    else
        TraySetIcon A_ScriptDir "\Icons\macropad.ico"
} catch {
    try{
        TraySetIcon A_ScriptDir "\Icons\macropad.ico"
    }
    }
 
contextcolor()
contextcolor(Dark:=2) { ; leave this at 1 to have system aware menus in light/dark. You may need to reload if you change your theme
    static uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
    static SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
    static FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")
    DllCall(SetPreferredAppMode, "int", Dark)
    DllCall(FlushMenuThemes)
}

#q::ShowMenu() ; add a hotkey to show the menu without the icon. I personally don't use this, but some may find use for it. 

ShowMenu() {
    TrayManager.RefreshMenu()
    ; SetTimer(ShowDelayedMenu, -100)  ; Wait 100ms before showing the menu. I personally don't use this, but you can add this if you want to ensure pause/suspend update states are read before showing the menu.
}

ShowDelayedMenu() {
    A_TrayMenu.Show()
}

class TrayManager {
    static runningScripts := Map()
    static scriptStates := Map()  ; New map to track script states between refreshes
    static trayMenu := A_TrayMenu
    static WM_COMMAND := 0x111
    static WM_ENTERMENULOOP := 0x211
    static WM_EXITMENULOOP := 0x212
    static WM_RBUTTONDOWN := 0x204
    static WM_RBUTTONUP := 0x205
    static editAction := ""
        static excludedScripts := [
        "launcher.ahk",
        "AHK Scripts.ahk", 
        "AutoHotkeyUX.exe",
        "Run AHK Scripts.ahk"
        ]
    static commands := Map(
        "Open", 65300,
        "Help", 65301,
        "Spy", 65302,
        "Reload", 65303,
        "Edit", 65304,
        "Suspend", 65305,
        "Pause", 65306,
        "Exit", 65307
    )

    static __New() {
        this.RefreshMenu()
        SetTimer(() => this.RefreshMenu(), 15000) ; change the 15000 value (15 seconds) to update more/less frequently, if desired.
    }

    static GetScriptState(hwnd) {
        static MF_CHECKED := 0x0008
        
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return { paused: false, suspended: false }
            
        mainMenu := DllCall("GetMenu", "Ptr", hwnd, "Ptr")
        if (!mainMenu)
            return { paused: false, suspended: false }
            
        fileMenu := DllCall("GetSubMenu", "Ptr", mainMenu, "Int", 0, "Ptr")
        if (!fileMenu) {
            DllCall("CloseHandle", "Ptr", mainMenu)
            return { paused: false, suspended: false }
        }
        
        pauseState := DllCall("GetMenuState", "Ptr", fileMenu, "UInt", 4, "UInt", 0x400)
        suspendState := DllCall("GetMenuState", "Ptr", fileMenu, "UInt", 5, "UInt", 0x400)
        
        DllCall("CloseHandle", "Ptr", fileMenu)
        DllCall("CloseHandle", "Ptr", mainMenu)
        
        return { 
            paused: (pauseState & MF_CHECKED) == MF_CHECKED,
            suspended: (suspendState & MF_CHECKED) == MF_CHECKED
        }
    }

    static UpdateScriptState(pid, command) {
        if (!this.scriptStates.Has(pid))
            this.scriptStates[pid] := { paused: false, suspended: false }
            
        if (command = "Pause")
            this.scriptStates[pid].paused := !this.scriptStates[pid].paused
        else if (command = "Suspend")
            this.scriptStates[pid].suspended := !this.scriptStates[pid].suspended
    }

   static GetRunningAHKScripts() {
    scripts := []
    DetectHiddenWindows(true)
    
    ; Get all windows with the AutoHotkey class (both .ahk and .exe)
    winList := WinGetList("ahk_class AutoHotkey")
    
    for hwnd in winList {
        try {
            ; Get the process path for the window
            pid := WinGetPID(hwnd)
            processPath := ProcessGetPath(pid)
            
            ; Skip if the process path is empty or doesn't exist
            if (!processPath || !FileExist(processPath))
                continue
            
            ; Get the full window title which contains the script path
            fullTitle := WinGetTitle(hwnd)
            foundName := false
            scriptPath := ""  ; Initialize scriptPath variable
            
            ; For standard AutoHotkey interpreters
            if (InStr(processPath, "AutoHotkey") && (InStr(processPath, "AutoHotkeyU64.exe") || 
                InStr(processPath, "AutoHotkeyU32.exe") || InStr(processPath, "AutoHotkey64.exe") || 
                InStr(processPath, "AutoHotkey32.exe") || InStr(processPath, "AutoHotkey64_UIA.exe"))) {
                
                ; Extract the script path from the window title
                if (RegExMatch(fullTitle, "^(.*) - AutoHotkey(?: v[0-9\.]+)?$", &match)) {
                    scriptPath := match[1]
                    
                    ; Skip if the script path doesn't exist
                    if (FileExist(scriptPath)) {
                        ; Get the script name for display
                        SplitPath(scriptPath, &scriptName)
                        foundName := true
                    }
                }
            }
            
            ; For compiled .exe scripts or fallback
            if (!foundName) {
                ; For elevated scripts, still try to keep track of the actual .ahk file
                if (InStr(processPath, "AutoHotkey64_UIA.exe")) {
                    ; Try to extract from window title again as a fallback
                    if (RegExMatch(fullTitle, "^(.*\.ahk)", &match)) {
                        possiblePath := match[1]
                        if (FileExist(possiblePath)) {
                            scriptPath := possiblePath
                            SplitPath(scriptPath, &scriptName)
                            foundName := true
                        }
                    }
                    
                    if (!foundName) {
                        scriptName := "UIA_Script_" pid  ; Generic fallback
                    }
                } else {
                    ; Get script name from process path for compiled scripts
                    SplitPath(processPath, &scriptName)
                    scriptPath := processPath
                    foundName := true
                }
            }
            
            if (!foundName) {
                ; Last fallback - just use the PID
                scriptName := "AHK_Script_" pid
                scriptPath := processPath
            }
            
            if (this.ShouldExcludeScript(scriptName))
                continue
            
            ; Get the script state
            state := this.GetScriptState(hwnd)
            
            if (this.scriptStates.Has(pid)) {
                state := this.scriptStates[pid]
            } else {
                this.scriptStates[pid] := state
            }

            scripts.Push({
                pid: pid,
                name: scriptName,
                path: scriptPath,  ; Store the actual script path when available
                processPath: processPath,  ; Also store the process path separately
                hwnd: hwnd,
                paused: state.paused,
                suspended: state.suspended
            })
        }
    }
    return scripts
}

    static ShouldExcludeScript(scriptName) {
        for excludedScript in this.excludedScripts {
            if (scriptName = excludedScript)
                return true
        }
        return false
    }



    static RefreshMenu() {
        scripts := this.GetRunningAHKScripts()
    
        this.runningScripts.Clear()
        for script in scripts
            this.runningScripts[script.pid] := script
    
        this.trayMenu.Delete()
    
        ; Collect scripts into an array
        sortedScripts := []
        for pid, script in this.runningScripts
            sortedScripts.Push({pid: pid, script: script})
    
                ; Sort scripts alphabetically by name
        for i, item1 in sortedScripts {
            for j, item2 in sortedScripts {
                if (j <= i)
                    continue

                if (StrCompare(sortedScripts[i].script.name, sortedScripts[j].script.name, true) > 0) {
                    temp := sortedScripts[i]
                    sortedScripts[i] := sortedScripts[j]
                    sortedScripts[j] := temp
                }
            }
        }

    
        ; Add scripts to the tray menu
        for entry in sortedScripts {
            pid := entry.pid
            script := entry.script
    
            scriptSubMenu := Menu()
            scriptSubMenu.Add("Tray Menu", this.CreateShowTrayMenuCallback(pid))
            scriptSubMenu.Add()
            scriptSubMenu.Add("Open File Location", this.CreateOpenFileLocationCallback(script.path))
            scriptSubMenu.Add("Reload", this.CreateCommandCallback(pid, "Reload"))
    
            includes := this.GetIncludeFiles(script.path)
            if (includes.Length > 0) {
                editSubMenu := Menu()
                editSubMenu.Add(script.name, this.CreateEditCallback(script.path))
                editSubMenu.Add()
    
                for includePath in includes {
                    SplitPath(includePath,, &dir, &ext, &name)
                    menuName := name "." ext
                    editSubMenu.Add(menuName, this.CreateEditCallback(includePath))
                }
                scriptSubMenu.Add("Edit", editSubMenu)
            } else {
                scriptSubMenu.Add("Edit", this.CreateEditCallback(script.path))
            }
    
            scriptSubMenu.Add("Pause", this.CreateCommandCallback(pid, "Pause"))
            scriptSubMenu.Add("Suspend", this.CreateCommandCallback(pid, "Suspend"))
    
            if (script.paused)
                scriptSubMenu.Check("Pause")
            if (script.suspended)
                scriptSubMenu.Check("Suspend")
    
            scriptSubMenu.Add()
            scriptSubMenu.Add("Exit", this.CreateCommandCallback(pid, "Exit"))
    
            this.trayMenu.Add(script.name, scriptSubMenu)
        }
    
        this.trayMenu.Add()
        this.trayMenu.Add("Reload", (*) => (this.RefreshMenu(), Reload())) ; Sleep(500), Reload()))
        this.trayMenu.Add("Window Spy", (*) => OpenWindowSpy())
    
        this.trayMenu.Add()
        this.trayMenu.Add("Exit All", (*) => this.ExitAllScripts())
    }
    
    static CreateOpenFileLocationCallback(scriptPath) {
        return (*) => this.OpenFileLocation(scriptPath)
    }
    

    static OpenFileLocation(scriptPath) {
        if (scriptPath = "") {
            MsgBox("Script path is empty.")
            return
        }
    
        SplitPath(scriptPath,, &dir)
        if (!DirExist(dir)) {
            MsgBox("Directory does not exist: " dir)
            return
        }
    
        try {
            Run('explorer.exe "' dir '"')
        } catch as err {
            MsgBox("Failed to open directory: " err.Message)
        }
    }
    
    

    static CreateEditCallback(scriptPath) {
        return (*) => this.EditScriptFile(scriptPath)
    }

    static ShowEditMenu(script) {
        includes := this.GetIncludeFiles(script.path)
        
        if (includes.Length = 0) {
            this.EditScriptFile(script.path)
            return
        }
        
        editMenu := Menu()
        editMenu.Add(script.name, (*) => this.EditScriptFile(script.path))
        editMenu.Add()
        
        for includePath in includes {
            SplitPath(includePath,, &dir, &ext, &name)
            menuName := name "." ext
            editMenu.Add(menuName, this.CreateIncludeCallback(includePath))
        }
        
        MouseGetPos(&mouseX, &mouseY)
        editMenu.Show(mouseX, mouseY)
    }

    static GetIncludeFiles(scriptPath) {
        includes := []
        
        if !FileExist(scriptPath)
            return includes
            
        SplitPath(scriptPath,, &scriptDir)
        
        try {
            fileContent := FileRead(scriptPath)
            
            loop parse, fileContent, "`n", "`r" {
                if RegExMatch(A_LoopField, "i)^\s*#Include\s+(.+)$", &match) {
                    includePath := match[1]
                    includePath := Trim(includePath, " `t`"'")
                    
                    if !RegExMatch(includePath, "^[A-Za-z]:\\") {
                        includePath := scriptDir "\" includePath
                    }
                    
                    if FileExist(includePath) {
                        includes.Push(includePath)
                    }
                    else if FileExist(A_MyDocuments "\AutoHotkey\Lib\" includePath) {
                        includes.Push(A_MyDocuments "\AutoHotkey\Lib\" includePath)
                    }
                }
            }
        }
        return includes
    }

    static CreateIncludeCallback(path) {
        return (*) => this.EditScriptFile(path)
    }

    static CreateShowTrayMenuCallback(pid) {
        return (*) => this.ShowOriginalTrayMenu(pid)
    }

    static CreateCommandCallback(pid, command) {
        if (command = "Reload") {
            return (*) => (
                this.SendCommand(pid, command),
                this.scriptStates.Delete(pid),
                SetTimer(() => this.RefreshMenu(), -500)
            )
        }
        return (*) => (
            this.SendCommand(pid, command),
            this.UpdateScriptState(pid, command),
            this.RefreshMenu()
        )
    }

    static GetScriptNameFromPath(path) {
        SplitPath(path,, &dir, &ext, &name)
        return name "." ext
    }

    static GetScriptPathFromCmd(cmdLine) {
        if (RegExMatch(cmdLine, '\"([^\"]*\.ahk)\"', &match))
            return match[1]
        
        if (RegExMatch(cmdLine, "([^\s]*\.ahk)", &match))
            return match[1]
            
        return ""
    }

 static EditScriptFile(scriptPath) {
    if (scriptPath = "")
        return
        
    ; Check if this is an .exe file and look for a corresponding .ahk file
    SplitPath(scriptPath, &fileName, &scriptDir, &fileExt, &fileNameNoExt)
    
    ; Special handling for AutoHotkey64_UIA.exe
    if (InStr(scriptPath, "AutoHotkey64_UIA.exe")) {
        ; Try to find the actual script being run by looking at command line arguments
        for hwnd in WinGetList("ahk_exe AutoHotkey64_UIA.exe") {
            pid := WinGetPID(hwnd)
            ; Try to get window title which often contains script path
            fullTitle := WinGetTitle(hwnd)
            if (RegExMatch(fullTitle, "^(.*) - AutoHotkey(?: v[0-9\.]+)?$", &match)) {
                scriptPath := match[1]
                if (FileExist(scriptPath)) {
                    break  ; Found the script, exit the loop
                }
            }
        }
    }
    ; If it's an .exe file, check if there's a matching .ahk file in the same directory
    else if (fileExt = "exe") {
        ahkPath := scriptDir "\" fileNameNoExt ".ahk"
        if (FileExist(ahkPath))
            scriptPath := ahkPath  ; Use the .ahk file instead
    }
        
    try {
        editorpath := A_AppData "\..\Local\Programs\Microsoft VS Code\Code.exe"
        if FileExist(editorpath) {
            Run('"' editorpath '" "' scriptPath '"')
            return
        }
    }
        
    if (this.editAction != "") {
        action := StrReplace(this.editAction, "$SCRIPT_PATH", scriptPath)
        try {
            Run(action)
            return
        }
    }
    
    try {
        Run("edit " scriptPath)
        return
    }
    
    try {
        Run('notepad.exe "' scriptPath '"')
    }
}

    static ShowOriginalTrayMenu(targetPID) {
        DetectHiddenWindows(true)
        if (hwnd := WinExist("ahk_pid " targetPID " ahk_class AutoHotkey")) {
            PostMessage(0x404, 0, this.WM_RBUTTONDOWN,, "ahk_id " hwnd)
            PostMessage(0x404, 0, this.WM_RBUTTONUP,, "ahk_id " hwnd)
        }
    }

    static SendCommand(targetPID, command) {
        DetectHiddenWindows(true)
        if (hwnd := WinExist("ahk_pid " targetPID " ahk_class AutoHotkey")) {
            PostMessage(this.WM_COMMAND, this.commands[command], 0,, "ahk_id " hwnd)
            if (command = "Reload") {
                Sleep(200)  ; Give some time for the reload to complete
            }
        }
    }

    static ExitAllScripts() {
        Result := MsgBox("Are you sure you want to exit all AutoHotkey scripts?",, "YesNo")
        if Result = "No"
            return
        
        for pid, script in this.runningScripts {
            if (pid != ProcessExist()) {
                this.scriptStates.Delete(pid)  ; Clear state when exiting
                this.SendCommand(pid, "Exit")
            }
        }
        ExitApp
    }
}

OpenWindowSpy() {
    Run('"C:\Users\' A_UserName '\OneDrive\Documents\AutoHotkey\lib\WindowSpy.ahk"')
}

TrayManager.__New()
