#NoEnv
#Warn
SetWorkingDir %A_ScriptDir% 
#SingleInstance force
#Persistent

; Set a default Status
CurrentStatus = Unknown

logPath = %A_AppData%\Microsoft\Teams\logs.txt
lt := new CLogTailer(logPath, Func("NewLine"))
return

NewLine(text)
{
    global CurrentStatus
    ReadStatus := RegExMatch(text, "StatusIndicatorStateService: Added (?!NewActivity)(\w+)", StatusText)
    if (ReadStatus != 0)
    {
        CurrentStatus := RegExReplace(StatusText1, "[^A-Z\s]\K([A-Z])", " $1")
        ShowNotification()
    }
}

class CLogTailer {
    __New(logfile, callback){
        this.file := FileOpen(logfile, "r-d")
        this.callback := callback
        ; Move seek to end of file
        this.file.Seek(0, 2)
        fn := this.WatchLog.Bind(this)
        SetTimer, % fn, 100
    }
    
    WatchLog(){
        Loop {
            p := this.file.Tell()
            l := this.file.Length
            line := this.file.ReadLine(), "`r`n"
            len := StrLen(line)
            if (len){
                RegExMatch(line, "[\r\n]+", matches)
                if (line == matches)
                    continue
                this.callback.Call(Trim(line, "`r`n"))
            }
        } until (p == l)
    }
}

; Function to display a message box with the current status
ShowNotification()
{
    global CurrentStatus
    MsgBox, % "Your Microsoft Teams Status: " CurrentStatus
}
