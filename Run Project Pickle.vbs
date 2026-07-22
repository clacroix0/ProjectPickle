Option Explicit

Dim shell, fso, folder, launcher, exitCode, extraArgs

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
launcher = fso.BuildPath(folder, "Project Pickle App Files\ProjectPickle.vbs")

If Not fso.FileExists(launcher) Then
    MsgBox "Could not find the Project Pickle app launcher in the Project Pickle App Files folder.", vbCritical, "Project Pickle"
    WScript.Quit 1
End If

extraArgs = ""
If WScript.Arguments.Named.Exists("SelfTest") Then extraArgs = extraArgs & " /SelfTest"
If WScript.Arguments.Named.Exists("UiSmokeTest") Then extraArgs = extraArgs & " /UiSmokeTest"
If WScript.Arguments.Named.Exists("NoUi") Then extraArgs = extraArgs & " /NoUi"

exitCode = shell.Run(Chr(34) & launcher & Chr(34) & extraArgs, 0, True)
WScript.Quit exitCode
