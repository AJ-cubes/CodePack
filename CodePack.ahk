#Requires AutoHotkey v2.0
#SingleInstance Force
#Include <ClipboardHistory>
; Big thanks to: AutoHotkey Forum – teadrinker, "ClipboardHistory class (partial implementation of Windows UWP Clipboard API for Windows 10/11)"
; https://www.autohotkey.com/boards/viewtopic.php?p=606025#p515653

; Icon taken from https://www.flaticon.com/free-icon/folder_1091647 - Thanks!

Initialize()

!c:: {
    try {
        tmpDir := A_Temp "\CodePack\Copy\"
        DirCreate tmpDir

        Send "^c"
        Sleep 500
        if DllCall("IsClipboardFormatAvailable", "UInt", 0xF) {
            files := DetectClipboardFiles()
            if files.Length {
                relPaths := GetRelativePaths(files)
                tmpFile := tmpDir "manifest.txt"
                f := FileOpen(tmpFile, "w", "UTF-8")
                for i, path in files {
                    if DirExist(path) {
                        continue
                    }
                    rel := relPaths[A_Index]
                    try {
                        contents := FileRead(path, "UTF-8")
                    } catch {
                        contents := "(could not read contents)"
                    }
                    f.Write("===" rel "===" "`n" contents "`n")
                }
                f.Close()
                ClipboardHistory.DeleteHistoryItem(1)
                FilesToClipboard([tmpFile])
            }
        } else if DllCall("IsClipboardFormatAvailable", "UInt", 13) {
            text := ClipboardHistory.GetHistoryItemText(1)
            tmpFile := tmpDir "code.txt"
            f := FileOpen(tmpFile, "w", "UTF-8")
            f.Write(text "`n")
            f.Close()
            ClipboardHistory.DeleteHistoryItem(1)
            FilesToClipboard([tmpFile])
        }
    } catch as err {
        msg := "Error occurred:`n"
        for prop, val in err.OwnProps()
            msg .= prop ": " val "`n"
        MsgBox msg, "Alt+C Error"
    }
}

!v:: {
    try {
        clip := 0
        tmpDir := A_Temp "\CodePack\Paste\"
        DirCreate tmpDir

        if DllCall("IsClipboardFormatAvailable", "UInt", 0xF) {
            files := DetectClipboardFiles()

            if files.Length != 1 {
                Send "^v"
                return
            }

            clip := FileRead(files[1], "UTF-8")
        } else if DllCall("IsClipboardFormatAvailable", "UInt", 13) {
            clip := ClipboardHistory.GetHistoryItemText(1)
        }

        if !clip {
            Send "^v"
            return
        }

        matches := []
        filePaths := []
        pos := 1

        while RegExMatch(clip, "s)===([^\r\n]+)===(.*?)(?=(?:\R===|$))", &m, pos) {
            path := tmpDir Trim(StrReplace(m[1], "/", "\"))
            content := Trim(RegExReplace(m[2], "^\R"))
            matches.Push({path: path, content: content})
            filePaths.Push(path)
            pos := m.Pos(0) + m.Len(0)
        }

        if matches.Length = 0 {
            Send "^v"
            return
        }

        for each, block in matches {
            SplitPath block.path,, &dir
            if dir != ""
                DirCreate dir
            f := FileOpen(block.path, "w", "UTF-8")
            f.Write(block.content)
            f.Close()
        }

        FilesToClipboard(filePaths, tmpDir)
        Send "^v"

        ClipboardHistory.PutHistoryItemIntoClipboard(1)
    } catch as err {
        msg := "Error occurred:`n"
        for prop, val in err.OwnProps()
            msg .= prop ": " val "`n"
        MsgBox msg, "Alt+V Error"
    }
}

DetectClipboardFiles() {
    if !DllCall("IsClipboardFormatAvailable", "UInt", 0xF)
        return []
    DllCall("OpenClipboard", "Ptr", 0)
    hDrop := DllCall("GetClipboardData", "UInt", 0xF, "Ptr")
    DllCall("CloseClipboard")
    if !hDrop
        return []
    count := DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", 0xFFFFFFFF, "Ptr", 0, "UInt", 0)
    files := []
    buf := Buffer(32768)
    Loop count {
        len := DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", A_Index-1, "Ptr", buf, "UInt", buf.Size//2)
        files.Push(StrGet(buf, len, "UTF-16"))
    }
    return files
}

FilesToClipboard(paths, basePath:=GetCommonPrefix(paths), method:="copy") {
    adjusted := []
    for p in paths {
        full := InStr(p, ":") ? p : basePath "\" p
        rel := SubStr(full, StrLen(basePath) + (InStr(basePath, "\",, -1) ? 1 : 0))
        target := InStr(rel, "\") ? basePath "\" StrSplit(rel, "\")[1] : full
        if !HasValue(adjusted, target) {
            adjusted.Push(target)
        }
    }
    PathLength := 0
    for f in adjusted
        PathLength += StrLen(f)
    pid := DllCall("GetCurrentProcessId","uint")
    hwnd := WinExist("ahk_pid " . pid)
    hPath := DllCall("GlobalAlloc","uint",0x42,"uint",20 + (PathLength + adjusted.Length + 1) * 2,"UPtr")
    pPath := DllCall("GlobalLock","UPtr",hPath)
    NumPut("UInt",20,pPath,0)
    NumPut("UInt",1,pPath,16)
    offset := 0
    for f in adjusted
        offset += StrPut(f, pPath + 20 + offset, StrLen(f) + 1, "UTF-16")
    DllCall("GlobalUnlock","UPtr",hPath)
    DllCall("OpenClipboard","UPtr",hwnd)
    DllCall("EmptyClipboard")
    DllCall("SetClipboardData","uint",0xF,"UPtr",hPath)
    mem := DllCall("GlobalAlloc","uint",0x42,"uint",4,"UPtr")
    str := DllCall("GlobalLock","UPtr",mem)
    DllCall("RtlFillMemory","UPtr",str,"uint",1,"UChar", (method="cut")?0x02:0x05)
    DllCall("GlobalUnlock","UPtr",mem)
    cfFormat := DllCall("RegisterClipboardFormat","Str","Preferred DropEffect")
    DllCall("SetClipboardData","uint",cfFormat,"UPtr",mem)
    DllCall("CloseClipboard")
    return adjusted
}

GetCommonPrefix(files) {
    if files.Length = 0
        return ""
    prefix := files[1]
    for str in files {
        while !InStr(str, prefix) = 1 {
            prefix := SubStr(prefix, 1, StrLen(prefix) - 1)
            if prefix = ""
                return ""
        }
    }
    pos := InStr(prefix, "\",, -1)
    if pos
        prefix := SubStr(prefix, 1, pos)
    return prefix
}

GetRelativePaths(files) {
    prefix := GetCommonPrefix(files)
    result := []
    for str in files {
        rel := SubStr(str, StrLen(prefix) + 1)
        result.Push(rel)
    }
    return result
}

HasValue(arr, val) {
    for item in arr {
        if item = val
            return 1
    }
    return 0
}

Initialize() {
    tmpDir := A_Temp "\CodePack\Icons\"
    DirCreate tmpDir

    FileInstall "Icons\favicon.ico", tmpDir "favicon.ico", true
    TraySetIcon tmpDir "favicon.ico"
}
