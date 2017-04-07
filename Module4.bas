Attribute VB_Name = "PublicAs"
Option Explicit
'强制声明变量

Public KeyCode1 As Long
Public KeyCode2 As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public PrPath As String
Public PrName As String
Public PrPid As Long
Public HookFirst As Boolean
Public HookHide As Boolean
Public FirstHide As Boolean

Public Fss As Boolean
