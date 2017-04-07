Attribute VB_Name = "GetPsPids"
Option Explicit
Private Declare Function CreateToolhelp32Snapshot _
                Lib "kernel32" (ByVal dwFlags As Long, _
                                ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First _
                Lib "kernel32" (ByVal hSnapShot As Long, _
                                lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next _
                Lib "kernel32" (ByVal hSnapShot As Long, _
                                lppe As PROCESSENTRY32) As Long
Public Declare Function TerminateProcess _
               Lib "kernel32" (ByVal hProcess As Long, _
                               ByVal uExitCode As Long) As Long
'VB 通过进程名称获取进程PID函数
'From http://www.newxing.com/Tech/Program/VisualBasic/PID_406.html

Private Declare Function OpenProcess _
                Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                ByVal bInheritHandle As Long, _
                                ByVal dwProcessId As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS = &H2&
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Const PROCESS_TERMINATE = 1
Function GetPsPid(sProcess As String) As Long
    Dim lSnapShot    As Long
    Dim lNextProcess As Long
    Dim tPE          As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

    If lSnapShot <> -1 Then
        tPE.dwSize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)

        Do While lNextProcess

            If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                Dim lProcess  As Long
                Dim lExitCode As Long
                GetPsPid = tPE.th32ProcessID

                CloseHandle lProcess
            End If

            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop

        CloseHandle (lSnapShot)
    End If

End Function

