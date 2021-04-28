Attribute VB_Name = "No1批量打开关闭模块"
Sub 批量打开关闭模块()
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
        MsgBox sPath & "路径不存在！"
    Else
        Application.DisplayAlerts = False
        Application.AskToUpdateLinks = False
        Application.ScreenUpdating = False
        path = i
        File = Dir(path & "*.xlsx")
        Do While File <> ""
            Set WB = Workbooks.Open(path & File)
            '模块插入
            
            '结束
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            File = Dir
        Loop
    End If
End Sub
