Attribute VB_Name = "No1�����򿪹ر�ģ��"
Sub �����򿪹ر�ģ��()
    Dim path As String
    Dim File As String
    Dim WB As Workbook
    Dim i As String
    Dim sResult As String
    i = InputBox("Your xls/xlsx Path") & "\"
    sResult = Dir(i, vbDirectory)
    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    If Len(sResult) = 0 Then
        MsgBox sPath & "·�������ڣ�"
    Else
        Application.DisplayAlerts = False
        Application.AskToUpdateLinks = False
        Application.ScreenUpdating = False
        path = i
        File = Dir(path & "*.xlsx")
        Do While File <> ""
            Set WB = Workbooks.Open(path & File)
            'ģ�����
            
            '����
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            File = Dir
        Loop
    End If
End Sub
