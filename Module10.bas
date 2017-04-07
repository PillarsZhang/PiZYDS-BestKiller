Attribute VB_Name = "FirstPrepation"
Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public MyVision As String
Public MyCHName As String
Public MyENName As String

Public NtsdFile As String
Public AppEXE() As Byte
Public FileNum As Long

Public Sub FirstPrepare()
 MyVision = "V" & App.Major & "." & App.Minor & "." & App.Revision
 MyCHName = "PiZYDS-极杀-BestKiller " & MyVision
 MyENName = "PiZYDS-BestKiller " & MyVision
 Call PrepareNTSD
End Sub

'http://bbs.csdn.net/topics/16139

Public Function GetWindowTempPath() As String
       Dim Dummy As Long, StrLen As Long, TempPath As String
       StrLen = 255
       TempPath = String$(StrLen, 0)
       Dummy = GetTempPath(StrLen, TempPath)
       If Dummy Then
          GetWindowTempPath = Left$(TempPath, Dummy)
       Else
          GetWindowTempPath = ""
       End If
End Function

'http://www.educity.cn/wenda/335217.html
'VB调用资源文件例子(文件打包与释放)
'章鱼DS 改进后可以释放到系统临时目录


Public Sub PrepareNTSD()
NtsdFile = GetWindowTempPath() & "PiZYDS-BestKiller\"

'On Error Resume Next
If Dir(NtsdFile, vbDirectory) = "" Then   '判断文件夹是否存在
    MkDir (NtsdFile)   '创建文件夹
End If
    
'MsgBox NtsdFile
AppEXE = LoadResData(101, "CUSTOM")
FileNum = FreeFile                   '以二进制方式写（生成）temp1.exe到当前目录
Open NtsdFile & "ntsd.exe" For Binary As #FileNum
Put #1, , AppEXE
Close #FileNum
End Sub
