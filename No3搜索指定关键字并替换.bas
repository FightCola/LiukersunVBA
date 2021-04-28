Attribute VB_Name = "No7搜索指定关键字并替换"
Sub 搜索指定关键字并替换()
    Dim path As String
    Dim File As String
    Dim WB As Workbook
    Dim i As String
    Dim sResult As String
    i = InputBox("Your xls/xlsx Path") & "\"
    sResult = Dir(i, vbDirectory)
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    If Len(sResult) = 0 Then
        MsgBox sPath & "路径不存在！"
    Else
        path = i
        File = Dir(path & "*.xlsx")
        Do While File <> ""
            Set WB = Workbooks.Open(path & File)
            Rows("1:1").Select
            Cells.Replace What:="-分表后保留", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            File = Dir
        Loop
    End If
End Sub




