- 👋 Hi, I’m @morteza43288
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...

<!---
morteza43288/morteza43288 is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcesses
    If InStr(objProcess.CommandLine, "explorer.exe /e,E:") > 0 Then
        WScript.Echo "Process Name: " & objProcess.Name
        WScript.Echo "Process ID: " & objProcess.ProcessID
        WScript.Echo "Command Line: " & objProcess.CommandLine
        WScript.Echo "-----------------------------------"
    End If
Next



______<<<
Set colProcesses = objWMIService.ExecQuery("Select Name, ProcessID, Caption, CommandLine, ThreadCount, WorkingSetSize from Win32_Process")

For Each objProcess in colProcesses
    WScript.Echo "Process Name: " & objProcess.Name
    WScript.Echo "Process ID: " & objProcess.ProcessID
    WScript.Echo "Caption: " & objProcess.Caption
    WScript.Echo "Command Line: " & objProcess.CommandLine
    WScript.Echo "Thread Count: " & objProcess.ThreadCount
    WScript.Echo "Working Set Size: " & objProcess.WorkingSetSize
    WScript.Echo "-----------------------------------"
Next

------
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcesses
    WScript.Echo "Process Name: " & objProcess.Name
    WScript.Echo "Process ID: " & objProcess.ProcessID
    WScript.Echo "Executable Path: " & objProcess.ExecutablePath
    WScript.Echo "-----------------------------------"
Next
