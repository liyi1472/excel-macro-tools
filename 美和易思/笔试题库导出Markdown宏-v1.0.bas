Attribute VB_Name = "utils"
'����⵼��ΪMarkdown��ʽ
Sub export2md()
Attribute export2md.VB_ProcData.VB_Invoke_Func = " \n14"

    '�����һ�Ź�����
    Worksheets(1).Activate
    '��¼��ǰ�к�
    Dim curRow As Integer
    '������ʼ�к�
    curRow = 2
    '��¼��ǰ���⡢ѡ�����ȷ��
    Dim question As String
    Dim optionA As String
    Dim optionB As String
    Dim optionC As String
    Dim optionD As String
    Dim answer As String
    '�ռ�Markdown�ı�����
    Dim mdContents As String
    '���д�������Markdown�ı�
    question = Range("B" & curRow).Value
    While question <> ""
        '�����е��ص����ݱ�ע
        question = Replace(question, "����", "<span style=""color:red;"">**����**</span>")
        question = Replace(question, "����ȷ", "<span style=""color:red;"">**����ȷ**</span>")
        '�����β�ʺ�ͳһ��׼
        question = question & "��"
        question = Replace(question, "����", "��")
        question = Replace(question, "����", "��")
        question = Replace(question, "����", "��")
        '׷������
        mdContents = mdContents & (curRow - 1) & ". " & question & vbCrLf
        '׷��ѡ��
        optionA = Trim(Range("G" & curRow).Value)
        optionB = Trim(Range("H" & curRow).Value)
        optionC = Trim(Range("I" & curRow).Value)
        optionD = Trim(Range("J" & curRow).Value)
        '��ȷ��
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
        'ƴ���ı�
        mdContents = mdContents & "     - " & optionA & vbCrLf
        mdContents = mdContents & "     - " & optionB & vbCrLf
        mdContents = mdContents & "     - " & optionC & vbCrLf
        mdContents = mdContents & "     - " & optionD & vbCrLf
        '��ȡ��һ����Ŀ��ѡ��
        curRow = curRow + 1
        question = Range("B" & curRow).Value
    Wend
    
    'д���ļ�
    writeFile mdContents

End Sub

'д���ļ�
Sub writeFile(contents As String)
    '�ļ���
    Dim fileName As String
    fileName = ThisWorkbook.Name
    fileName = Replace(fileName, ".xlsm", "")
    fileName = Replace(fileName, ".xlsx", "")
    fileName = Replace(fileName, ".xls", "")
    fileName = fileName & ".md"
    '�ļ�·��
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & fileName
    Open filePath For Output As #1
    Print #1, contents
    Close #1
End Sub
