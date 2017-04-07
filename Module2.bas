Attribute VB_Name = "WindowShowHiding"
Public Sub WindowShowHide() ' 用于隐藏显示窗口
    Select Case Frm_Main.Visible
        Case True
            If frmAbout.Visible Then Unload frmAbout
            Frm_Main.Hide
            If Frm_Senior.Visible Then
             Call Frm_Main.ShowSenior
             FirstHide = True
             Else
             FirstHide = False
             End If
        Case False
            Frm_Main.Show
            If FirstHide Then Call Frm_Main.ShowSenior
    End Select
End Sub
