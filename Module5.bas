Attribute VB_Name = "BestKilling"

Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Public Sub bestkill()
 Dim PathKill As String
 Dim PathOpen As String
 PrPid = GetPsPid(PrName)
 'MsgBox PrPid
 PathOpen = Chr(34) & PrPath & Chr(34)
 'pathkill = App.Path & "ntsd -c q -p " & PrPid
 
 If GetPsPid(PrName) = 0 Then
  If Frm_Main.Check1.Value = 1 Then
   If PrPath = "" Or PrPath = "?" Then
    MsgBox ("打开地址未知" & PrPath)
    Else
   Shell PathOpen
  End If
  End If
 Else
 Do While GetPsPid(PrName) <> 0
  'MsgBox pathkill
  PrPid = GetPsPid(PrName)
  PathKill = NtsdFile & "\ntsd -c q -p " & PrPid
  Shell PathKill, vbHide
  Sleep 500
  Loop
 End If
End Sub

