Attribute VB_Name = "GetPath"
'https://zhidao.baidu.com/question/319190532.html
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Function GetProcessFilePath(PID As Long) As String   'PID 是你要获取文件路径的进程的PID，这个可以在任务管理器中查到
    Dim PH As Long
    Dim FileName As String * 1024
    PH = OpenProcess(&H1E00FF Or &H10F00, False, PID)
    Call GetModuleFileNameExA(PH, 0, FileName, 1024)
    GetProcessFilePath = Trim(FileName)
    CloseHandle PH
End Function



Public Function GetPrPath(processName As String)
 GetPrPath = GetProcessFilePath(GetPsPid(processName))
 PrPid = GetPsPid(processName)
 'MsgBox GetPrPath
End Function
