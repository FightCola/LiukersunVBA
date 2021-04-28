Attribute VB_Name = "toPDF"
Sub topdf()
    Dim sName As String
    Dim path As String
    Dim File As String
    Dim WB As Workbook
    Dim i As String
    Dim sResult As String
    i = InputBox("Your xls/xlsx Path") & "\"
    sResult = Dir(i, vbDirectory)
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    If Len(sResult) = 0 Then
        MsgBox sPath & "·�������ڣ�"
    Else
        path = i
        File = Dir(path & "*.xlsx")
        Do While File <> ""
            Set WB = Workbooks.Open(path & File)
            'ģ�����
            sName = ActiveWorkbook.Name
            sName = Left(sName, InStrRev(sName, ".") - 1)
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=i & sName & ".pdf", OpenAfterPublish:=False
            '����
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            File = Dir
        Loop
    End If
End Sub
