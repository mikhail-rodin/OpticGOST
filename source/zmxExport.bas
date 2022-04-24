Attribute VB_Name = "zmxExport"
Sub exportSelectionXLSX()
Dim SrcSheet, DestSheet As Worksheet
Dim DestBook As Workbook
Dim SrcRange As Range
Dim sel As String
Dim strFilename As String
        
Set dialog = Application.FileDialog(msoFileDialogSaveAs)

    With dialog
'        .Filters.Clear
'        .Filters.Add "Все файлы", "*.*"
'        .Filters.Add "Excel", "*.xlsx", 1
        .Title = "Сохранение диапазона в отдельный документ XLSX"
        .InitialView = msoFileDialogViewList
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
        If .Show Then
            strFilename = .SelectedItems(1)
        Else
            MsgBox "No filename specified!", vbExclamation
            Exit Sub
        End If
    End With
    
    Set SrcSheet = ActiveSheet
        sel = Selection.Address
        If sel = "" Then
            MsgBox "Ничего не выбрано"
            Exit Sub
        End If
        Set SrcRange = SrcSheet.Range(sel)
            Set DestBook = Application.Workbooks.Add
                Set DestSheet = DestBook.Worksheets(1)
                    SrcRange.Copy DestSheet.Range("A1")
                Set DestSheet = Nothing
                DestBook.SaveAs filename:=strFilename, _
                    FileFormat:=XlFileFormat.xlOpenXMLWorkbook, _
                        CreateBackup:=False
                DestBook.Close 0
            Set DestBook = Nothing
        Set SrcRange = Nothing
    Set SrcSheet = Nothing

End Sub

Sub exportSheetXLSX()
Dim SrcSheet, DestSheet, sh As Worksheet
Dim DestBook As Workbook
Dim strFilename As String
        
Set dialog = Application.FileDialog(msoFileDialogSaveAs)

    With dialog
'        .Filters.Clear
'        .Filters.Add "Все файлы", "*.*"
'        .Filters.Add "Excel", "*.xlsx", 1
        .Title = "Сохранение листа в отдельный документ XLSX"
        .InitialView = msoFileDialogViewList
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
        If .Show Then
            strFilename = .SelectedItems(1)
        Else
            MsgBox "No filename specified!", vbExclamation
            Exit Sub
        End If
    End With
    
    Application.DisplayAlerts = False
    
    Set SrcSheet = ActiveSheet
        sel = Selection.Address
        If sel = "" Then
            MsgBox "Ничего не выбрано"
            Exit Sub
        End If
        Set SrcRange = SrcSheet.Range(sel)
            Set DestBook = Application.Workbooks.Add
                SrcSheet.Copy DestBook.Worksheets(1)
                'удалим пустые листы
                If DestBook.Worksheets.Count > 1 Then
                    For Each sh In Sheets
                        If IsEmpty(sh.UsedRange) Then
                            sh.Delete
                        End If
                    Next
                End If
                DestBook.SaveAs filename:=strFilename, _
                    FileFormat:=XlFileFormat.xlOpenXMLWorkbook, _
                        CreateBackup:=False
                DestBook.Close 0
            Set DestBook = Nothing
        Set SrcRange = Nothing
    Set SrcSheet = Nothing
    
    Application.DisplayAlerts = True
    
End Sub
