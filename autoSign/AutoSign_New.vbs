
while i<WScript.Arguments.count
    sArgument=WScript.Arguments.item(i)
    select case LCase(sArgument)
    case "-s":
        sPath=WScript.Arguments.item(i+1)
        i=i+1
    case "-e":
        sExtensions=WScript.Arguments.item(i+1)
        i=i+1
    case "-m":
        sMailList=WScript.Arguments.item(i+1)
        i=i+1
    end select
    i=i+1
wend

sCommand="python.exe SignCertificates.py -s "&spath&" -e "&sExtensions&" -m "&sMailList
WScript.echo sCommand
ShellRun sCommand

Sub ShellRun(sCommandStringToExecute)
    Dim oShellObject
    Dim iErr
    Set oShellObject = CreateObject("Wscript.Shell")
    oShellObject.Run "%comspec% /c " & sCommandStringToExecute, 0, True
    If iErr <> 0 Then 
        WScript.echo "Error:"&Err.Number
        Exit Sub 
    End If 
End Sub