'merge one or more xlsx' sheets to one xlsx's sheets
'and create index table of content and link
'by lwj1396@163.com
'2015-4-25


'定义一个find函数
Private Function find(s1 As String, s2 As String) As String
    Dim i As Integer
    For i = 1 To Len(s1)
        If i = InStr(1, s1, s2, vbTextCompare) Then
        find = Left(s1, i - 1) & Right(s1, Len(s1) - Len(s2) - i + 1)
        Exit Function
        End If
    Next i
    find = s1
End Function


'生成目录及超链接
Private Function getAllWorkSheets()
' 得到所有的sheet页名称，并加上超连接
    totalNum = Worksheets.Count
    Worksheets(1).Activate
    Range("B:B").Select
    Selection.NumberFormatLocal = "@"
    ' 从2开始就是不带“目录”Sheet页，如果要带，则从1开始
    For index_i = 2 To totalNum
        sheetName = Worksheets(index_i).Name
        strarr = Split(sheetName, "-")
        Cells(index_i, 1) = strarr(0)
        Cells(index_i, 2) = sheetName
        tar_sheet = "'" & sheetName & "'"
        Cells(index_i, 3).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        tar_sheet & "!A1", TextToDisplay:=">>>>>>>>>>>>>>>>>>>>>>"
    Next index_i
    mergeA(totalNum)
End Function


'适用于单列区域
Private Function mergeA(c)                                      
    With [A:A]
        .Offset(0, 1).EntireColumn.Insert
        For i = 1 To c - 1
            If .Cells(i) = .Cells(i + 1) Then .Cells(i).Offset(0, 1).Resize(2, 1).Merge
        Next
        .Offset(0, 1).Copy
        .PasteSpecial xlPasteFormats
        .Offset(0, 1).EntireColumn.Delete
    End With
End Function


Sub 文件合并()
    On Error Resume Next
    Dim wb As Workbook, sh As Worksheet
    Dim fn As String, pt As String, t
    t = Timer
    '------------------------检测是否打开了多个excel文件-------------------
    If Workbooks.Count > 1 Then
      MsgBox "请关闭其余的工作簿！"
      Exit Sub
    End If
    '------------------------选择要合并的工作簿所在文件夹，获取路径---------
    With Application.FileDialog(msoFileDialogFolderPicker)
      .Show
      .AllowMultiSelect = False
      If .SelectedItems.Count = 0 Then
        MsgBox "没有选择任何文件夹！"
        Exit Sub
      Else
        pt = .SelectedItems(1)
      End If
    End With
    '-------------------------遍历文件夹中的所有Excel文件，并进行处理--------
    fn = Dir(pt & "\*.xlsx")      '若你的文档是2007或更新版本，则将*.xls改成*.xlsx
    Do While fn <> ""
      If fn <> ThisWorkbook.Name Then
        k = k + 1
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(pt & "\" & fn, , True)
        For i = 1 To wb.Worksheets.Count
            Set sh = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            wb.Worksheets(i).Rows.Copy sh.Rows
            '--------sh.Name = "" & Left(fn, Len(fn) - IIf(Right(fn, 1) = "x", 5, 4))
            sh.Name = wb.Name & "-" & wb.Worksheets(i).Name
            s = find(sh.Name, "工作簿")
            sh.Name = s
            s1 = find(sh.Name, ".xlsx")
            sh.Name = s1
        Next
        wb.Close
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
      End If
      fn = Dir
    Loop
      getAllWorkSheets
      MsgBox "处理结束。共处理" & k & "个文件，耗时" & Timer - t & "秒"
      ThisWorkbook.Save
End Sub
