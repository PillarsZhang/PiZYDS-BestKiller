Attribute VB_Name = "KeyCodeToString"
'该部分功能为将输入的键位转化为键位的符号以及键值(顺便使这个值合法)
'部分代码来自(由章鱼DS做部分修改) http://www.newxing.com/Tech/Program/VisualBasic/KeyCode_363.html

Public Function KeyCodeToStr(vbKeycode As Integer)

    If vbKeycode > 47 And vbKeycode < 91 Then
        KeyCodeToStr = Chr(vbKeycode)
        Exit Function
    ElseIf vbKeycode > 111 And vbKeycode < 124 Then
        KeyCodeToStr = "F" & vbKeycode - 111
        Exit Function
    ElseIf vbKeycode > 95 And vbKeycode < 106 Then
        KeyCodeToStr = vbKeycode - 96
        Exit Function
    End If
   
    Select Case vbKeycode
       
        Case 8
            KeyCodeToStr = "Back"

        Case 9
            KeyCodeToStr = "Tab"

        Case 12
            KeyCodeToStr = "Clear"

        Case 13
            KeyCodeToStr = "Enter"

        Case 16
            KeyCodeToStr = "Shift"

        Case 17
            KeyCodeToStr = "Ctrl"

        Case 18
            KeyCodeToStr = "Alt"

        Case 19
            KeyCodeToStr = "Pause"

        Case 20
            KeyCodeToStr = "Caps Lock"

        Case 27
            KeyCodeToStr = "Esc"

        Case 32
            KeyCodeToStr = "Space"

        Case 33
            KeyCodeToStr = "Page Up"

        Case 34
            KeyCodeToStr = "Page Down"

        Case 35
            KeyCodeToStr = "End"

        Case 36
            KeyCodeToStr = "Home"
           
        Case 41
            KeyCodeToStr = "Select"

        Case 42
            KeyCodeToStr = "Print Screen"
           
        Case 43
            KeyCodeToStr = "Execute"

        Case 44
            KeyCodeToStr = "SnapShot"
           
        Case 45
            KeyCodeToStr = "Insert"

        Case 46
            KeyCodeToStr = "Delete"
           
        Case 47
            KeyCodeToStr = "Help"
           
        Case 144
            KeyCodeToStr = "Num Lock"
           
        Case 189
            KeyCodeToStr = "-_"
           
        Case 187
            KeyCodeToStr = "=+"
           
        Case 255
            KeyCodeToStr = "Unknown"
           
        Case 192
            KeyCodeToStr = "`~"
           
        Case 37
            KeyCodeToStr = "Left Arrow"
                       
        Case 38
            KeyCodeToStr = "Up Arrow"
                       
        Case 39
            KeyCodeToStr = "Right Arrow"
                       
        Case 40
            KeyCodeToStr = "Dowm Arrow"
                       
        Case 219
            KeyCodeToStr = "[{"
                       
        Case 221
            KeyCodeToStr = "]}"
                       
        Case 186
            KeyCodeToStr = ";:"
                       
        Case 222
            KeyCodeToStr = "'"""
                       
        Case 220
            KeyCodeToStr = "\|"
                                   
        Case 188
            KeyCodeToStr = ",<"
                       
        Case 190
            KeyCodeToStr = ".>"
                       
        Case 191
            KeyCodeToStr = "/?"
                                   
        Case 193
            KeyCodeToStr = "\"
        Case Else
            KeyCodeToStr = "Unknown"
    End Select

End Function

