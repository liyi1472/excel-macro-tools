Attribute VB_Name = "utils"
'将题库导出为Markdown格式
Sub export2md()
Attribute export2md.VB_ProcData.VB_Invoke_Func = " \n14"

    '激活第一张工作表
    Worksheets(1).Activate
    '记录当前行号
    Dim curRow As Integer
    '设置起始行号
    curRow = 2
    '记录当前问题、选项和正确答案
    Dim question As String
    Dim optionA As String
    Dim optionB As String
    Dim optionC As String
    Dim optionD As String
    Dim answer As String
    '收集Markdown文本内容
    Dim mdContents As String
    '逐行处理生成Markdown文本
    question = Range("B" & curRow).Value
    While question <> ""
        '问题中的重点内容标注
        question = Replace(question, "错误", "<span style=""color:red;"">**错误**</span>")
        question = Replace(question, "不正确", "<span style=""color:red;"">**不正确**</span>")
        '问题结尾问号统一标准
        question = question & "？"
        question = Replace(question, "。？", "？")
        question = Replace(question, "？？", "？")
        question = Replace(question, "】？", "】")
        '追加问题
        mdContents = mdContents & (curRow - 1) & ". " & question & vbCrLf
        '追加选项
        optionA = Trim(Range("G" & curRow).Value)
        optionB = Trim(Range("H" & curRow).Value)
        optionC = Trim(Range("I" & curRow).Value)
        optionD = Trim(Range("J" & curRow).Value)
        '正确答案
        answer = Range("F" & curRow).Value
        If InStr(answer, "A") <> 0 Then
            optionA = "**" & optionA & "**"
        End If
        If InStr(answer, "B") <> 0 Then
            optionB = "**" & optionB & "**"
        End If
        If InStr(answer, "C") <> 0 Then
            optionC = "**" & optionC & "**"
        End If
        If InStr(answer, "D") <> 0 Then
            optionD = "**" & optionD & "**"
        End If
        '拼接文本
        mdContents = mdContents & "     - " & optionA & vbCrLf
        mdContents = mdContents & "     - " & optionB & vbCrLf
        mdContents = mdContents & "     - " & optionC & vbCrLf
        mdContents = mdContents & "     - " & optionD & vbCrLf
        '读取下一行题目和选项
        curRow = curRow + 1
        question = Range("B" & curRow).Value
    Wend
    
    '写入文件
    writeFile mdContents

End Sub

'写入文件
Sub writeFile(contents As String)
    '文件名
    Dim fileName As String
    fileName = ThisWorkbook.Name
    fileName = Replace(fileName, ".xlsm", "")
    fileName = Replace(fileName, ".xlsx", "")
    fileName = Replace(fileName, ".xls", "")
    fileName = fileName & ".md"
    '文件路径
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & fileName
    Open filePath For Output As #1
    Print #1, contents
    Close #1
End Sub
