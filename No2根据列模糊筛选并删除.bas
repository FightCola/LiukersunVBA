Attribute VB_Name = "No3������ģ��ɸѡ��ɾ��"
Sub ������ģ��ɸѡ��ɾ����()
    Dim i As Long, arr, K As Long
    Dim path As String
    Dim File As String
    Dim WB As Workbook
    Dim m As String
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    path = InputBox("Your xls/xlsx Path") & "\"
    File = Dir(path & "*.xlsx")
    Do While File <> ""
        Application.DisplayAlerts = False
        Application.AskToUpdateLinks = False
        Application.ScreenUpdating = False
        Set WB = Workbooks.Open(path & File)
        
        For i = 50 To 2 Step -1
            If Cells(1, i) Like "*�ֱ��ɾ��" Then
                Cells(1, i).EntireColumn.Delete
            End If
        Next i
        Application.ScreenUpdating = True
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        File = Dir
    Loop
End Sub
