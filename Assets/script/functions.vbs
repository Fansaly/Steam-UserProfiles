On Error Resume Next

Function vbPopup(info, time, title, icon)
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    vbPopup = WshShell.Popup(info, time, title, icon)
End Function

Function vbMsgBox(info, icon, title)
    vbMsgBox = MsgBox(info, icon, title)
End Function

Function RegExpTest(patrn, Str, ReStr)
    Dim regEx, Match, Matches   ' 创建变量
    Set regEx = New RegExp      ' 创建正则表达式
    regEx.Pattern = patrn       ' 设置模式
    regEx.IgnoreCase = True     ' 设置是否区分大小写
    regEx.Global = True         ' 设置全程匹配

    If (regEx.Test(Str) = True) Then
        If ReStr = "" Then
            Set Matches = regEx.Execute(Str)    ' 执行搜索
            For Each Match in Matches           ' 循环遍历 Matches 集合
                'RetStr = RetStr & "偏移量 "
                'RetStr = RetStr & Match.FirstIndex & "。" &vbCrLf& "字符：'"
                'RetStr = RetStr & Match.Value & "'。" & vbCrLf
                RetStr = RetStr & Match.Value
            Next
        Else
            RetStr = regEx.Replace(Str, ReStr)
        End If
    Else
        RetStr = False
    End If

    RegExpTest = RetStr
End Function
