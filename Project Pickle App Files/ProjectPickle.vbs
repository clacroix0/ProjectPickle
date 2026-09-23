Option Explicit

Dim fso
Dim shell
Dim appFolder
Dim appScript
Dim splashScript
Dim powershellPath
Dim powershell32Path
Dim powershell64Path
Dim tempFolder
Dim readyFile
Dim readyStream
Dim splashCommand
Dim appCommand
Dim quote

quote = Chr(34)

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appFolder = fso.GetParentFolderName(WScript.ScriptFullName)

appScript = fso.BuildPath( _
    appFolder, _
    "ProjectPickle.ps1" _
)

splashScript = fso.BuildPath( _
    appFolder, _
    "ProjectPickleSplash.ps1" _
)

If Not fso.FileExists(appScript) Then
    MsgBox _
        "Project Pickle could not find:" & _
        vbCrLf & vbCrLf & _
        appScript & _
        vbCrLf & vbCrLf & _
        "Keep ProjectPickle.vbs and ProjectPickle.ps1 together " & _
        "inside the Project Pickle App Files folder.", _
        vbCritical, _
        "Project Pickle"

    WScript.Quit 1
End If

powershell32Path = _
    shell.ExpandEnvironmentStrings("%WINDIR%") & _
    "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

powershell64Path = _
    shell.ExpandEnvironmentStrings("%WINDIR%") & _
    "\System32\WindowsPowerShell\v1.0\powershell.exe"

If fso.FileExists(powershell32Path) Then
    powershellPath = powershell32Path
ElseIf fso.FileExists(powershell64Path) Then
    powershellPath = powershell64Path
Else
    MsgBox _
        "Project Pickle could not find Windows PowerShell.", _
        vbCritical, _
        "Project Pickle"

    WScript.Quit 1
End If

shell.CurrentDirectory = appFolder

If fso.FileExists(splashScript) Then
    tempFolder = shell.ExpandEnvironmentStrings("%TEMP%")

    If Not fso.FolderExists(tempFolder) Then
        tempFolder = appFolder
    End If

    readyFile = fso.BuildPath( _
        tempFolder, _
        "ProjectPickleSplash_" & _
        Replace(fso.GetTempName, ".", "_") & _
        ".ready" _
    )

    Set readyStream = fso.CreateTextFile( _
        readyFile, _
        True _
    )

    readyStream.WriteLine "Project Pickle is starting."
    readyStream.Close

    Set readyStream = Nothing

    splashCommand = _
        quote & powershellPath & quote & _
        " -NoProfile" & _
        " -STA" & _
        " -ExecutionPolicy Bypass" & _
        " -WindowStyle Hidden" & _
        " -File " & quote & splashScript & quote & _
        " -ReadyFile " & quote & readyFile & quote

    appCommand = _
        quote & powershellPath & quote & _
        " -NoProfile" & _
        " -STA" & _
        " -ExecutionPolicy Bypass" & _
        " -WindowStyle Hidden" & _
        " -File " & quote & appScript & quote & _
        " -SplashReadyFile " & quote & readyFile & quote

    shell.Run splashCommand, 0, False

    WScript.Sleep 150
Else
    ' The splash is optional. Project Pickle should still open if the
    ' splash script is accidentally missing from a shared copy.
    appCommand = _
        quote & powershellPath & quote & _
        " -NoProfile" & _
        " -STA" & _
        " -ExecutionPolicy Bypass" & _
        " -WindowStyle Hidden" & _
        " -File " & quote & appScript & quote
End If

shell.Run appCommand, 0, False

Set shell = Nothing
Set fso = Nothing