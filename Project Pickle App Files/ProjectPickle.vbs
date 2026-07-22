Option Explicit

Dim shell, fso, scriptDir, scriptPath, splashPath
Dim systemRoot, powershellPath, tempFolder, splashReadyFile
Dim markerFile, markerCreated, splashStarted
Dim splashCommand, powershellCommand, launchError, i

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "ProjectPickle.ps1")
splashPath = fso.BuildPath(scriptDir, "ProjectPickleSplash.hta")

If Not fso.FileExists(scriptPath) Then
    MsgBox "ProjectPickle.ps1 was not found next to this launcher:" & vbCrLf & vbCrLf & scriptPath, vbCritical, "Project Pickle"
    WScript.Quit 1
End If

systemRoot = shell.ExpandEnvironmentStrings("%SystemRoot%")

' Prefer 32-bit Windows PowerShell, which is commonly required for
' 32-bit Microsoft Access database components.
powershellPath = systemRoot & "\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

If Not fso.FileExists(powershellPath) Then
    powershellPath = systemRoot & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

If Not fso.FileExists(powershellPath) Then
    MsgBox "Windows PowerShell could not be found.", vbCritical, "Project Pickle"
    WScript.Quit 1
End If

' Create a temporary file. The splash remains open while this file exists.
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%")
splashReadyFile = fso.BuildPath(tempFolder, "ProjectPickleSplash_" & fso.GetTempName)

markerCreated = False
Set markerFile = Nothing

On Error Resume Next

Set markerFile = fso.CreateTextFile(splashReadyFile, True)

If Err.Number = 0 Then
    markerFile.WriteLine "loading"
    markerFile.Close
    markerCreated = True
End If

Err.Clear
On Error GoTo 0

' Open the splash before starting PowerShell.
splashStarted = False

If markerCreated And fso.FileExists(splashPath) Then
    splashCommand = QuoteArgument(systemRoot & "\System32\mshta.exe") & _
        " " & QuoteArgument(splashPath) & _
        " " & QuoteArgument(splashReadyFile)

    On Error Resume Next

    shell.Run splashCommand, 1, False

    If Err.Number = 0 Then
        splashStarted = True
    End If

    Err.Clear
    On Error GoTo 0
End If

' Give the lightweight splash a small head start.
If splashStarted Then
    WScript.Sleep 80
End If

powershellCommand = QuoteArgument(powershellPath) & _
    " -NoLogo -NoProfile -STA -WindowStyle Hidden" & _
    " -ExecutionPolicy Bypass -File " & QuoteArgument(scriptPath)

If markerCreated Then
    powershellCommand = powershellCommand & _
        " -SplashReadyFile " & QuoteArgument(splashReadyFile)
End If

' Preserve any arguments passed to the VBS launcher.
For i = 0 To WScript.Arguments.Count - 1
    powershellCommand = powershellCommand & _
        " " & QuoteArgument(WScript.Arguments(i))
Next

launchError = ""

On Error Resume Next

shell.Run powershellCommand, 0, False

If Err.Number <> 0 Then
    launchError = Err.Description
End If

Err.Clear
On Error GoTo 0

If Len(launchError) > 0 Then
    If markerCreated Then
        On Error Resume Next
        fso.DeleteFile splashReadyFile, True
        On Error GoTo 0
    End If

    MsgBox "Project Pickle could not be started:" & _
        vbCrLf & vbCrLf & launchError, _
        vbCritical, "Project Pickle"

    WScript.Quit 1
End If

WScript.Quit 0

Function QuoteArgument(ByVal value)
    QuoteArgument = Chr(34) & _
        Replace(CStr(value), Chr(34), Chr(34) & Chr(34)) & _
        Chr(34)
End Function
