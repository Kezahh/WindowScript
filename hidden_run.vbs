output_file_path = "C:\users\kerryn\desktop\output.txt"
set objFSO = CreateObject("Scripting.FileSystemObject")
set out_file = objFSO.CreateTextFile(output_file_path, True)

Function output_message(message)
    out_file.write message & vbCrLf
End Function


'msgbox(Wscript.Arguments(0))

'CreateObject("WScript.Shell").Run "" & WScript.Arguments(0) & "", 0, False


output_message Wscript.Arguments(0)
output_message "hello"


set objShell = CreateObject("WScript.shell")
objShell.run "powershell -nologo -file move_window.ps1 1", 0
output_message "done"