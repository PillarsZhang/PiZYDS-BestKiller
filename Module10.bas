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
 MyCHName = "PiZYDS-��ɱ-BestKiller " & MyVision
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
'VB������Դ�ļ�����(�ļ�������ͷ�)
'����DS �Ľ�������ͷŵ�ϵͳ��ʱĿ¼


Public Sub PrepareNTSD()
NtsdFile = GetWindowTempPath() & "PiZYDS-BestKiller\"

'On Error Resume Next
If Dir(NtsdFile, vbDirectory) = "" Then   '�ж��ļ����Ƿ����
    MkDir (NtsdFile)   '�����ļ���
End If
    
'MsgBox NtsdFile
AppEXE = LoadResData(101, "CUSTOM")
FileNum = FreeFile                   '�Զ����Ʒ�ʽд�����ɣ�temp1.exe����ǰĿ¼
Open NtsdFile & "ntsd.exe" For Binary As #FileNum
Put #1, , AppEXE
Close #FileNum
End Sub
