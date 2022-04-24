VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rayForm 
   Caption         =   "Импорт файлов ZEMAX Raytrace"
   ClientHeight    =   3410
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7820
   OleObjectBlob   =   "rayForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "rayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public workingFilePath As String 'для удобства сохраняем последнюю директорию


Private Sub createSheetChk_Change()
    With rayForm
        If .createSheetChk.Value = True Then
            .sheetName.Enabled = True
        Else
            .sheetName.Enabled = False
        End If
    End With
End Sub


Public Sub deleteBtn_Click()
    Static i As Integer
    Static listLen As Integer
    Static tmpStr As String
    With rayForm.fileList
        i = 0
        listLen = .ListCount 'зафиксируем, т.к. он скоро поменяется и цикл может сломаеться
        While i <= listLen - 1
            If .Selected(i) = True Then
                'так как печататься будут только столбцы с непустым именем файла
                'сотрём имя файла
                tmpStr = .List(i, 0)
                If tmpStr = "главный" Then 'если файл импортирован
                    zmxImport.rays(0).chiefRayH = ""
                ElseIf tmpStr = "верхний" Then 'если файл импортирован
                    zmxImport.rays(0).upperRayH = ""
                ElseIf tmpStr = "нижний" Then 'если файл импортирован
                    zmxImport.rays(0).lowerRayH = ""
                ElseIf tmpStr = "апертурный" Then 'если файл импортирован
                    zmxImport.rays(0).axialRayH = ""
                End If
                'If tmpStr = "главный" Then 'если файл импортирован
                    'zmxImport.rays(0).chiefRayH = ""
                'End If
                .RemoveItem (i)
                listLen = listLen - 1
            End If
            i = i + 1
        Wend
    End With
End Sub



Private Sub fileList_AfterUpdate()
'    Static listLen As Integer
'
'    With rayForm
'        listLen = .fileList.ListCount
'        Select Case listLen
'            Case 0
'                .rayFillTableBtn.Enabled = False
'                .status.Caption = _
'                    "Сохраните в ZEMAX и загрузите сюда 4 отчёта Raytrace в текстовом варианте:" _
'                        & vbCrLf & "для апертурного (Hy=0, Py=1), главного (1,0), верхнего (1,1), нижнего (1,-1) лучей."
'            Case 1, 2
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "Загрузите ещё " & 4 - listLen & "файла. Должны быть загружены файлы" _
'                        & vbCrLf & "для апертурного (Hy=0, Py=1), главного (1,0), верхнего (1,1), нижнего (1,-1) лучей."
'            Case 3
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "Загрузите ещё один файл. Должны быть загружены файлы" _
'                        & vbCrLf & "для апертурного (Hy=0, Py=1), главного (1,0), верхнего (1,1), нижнего (1,-1) лучей."
'            Case 4
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "Загружено 4 файла. Можно заполнить таблицу."
'            Case Else
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "Для заполнения таблицы нужно только 4 файла. Лишние можно удалить."
'        End Select
'
'    End With
End Sub


Private Sub fileList_Click()
    Static i As Integer
    With rayForm.fileList
        rayForm.deleteBtn.Enabled = False
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                rayForm.deleteBtn.Enabled = True
            End If
        Next
    End With
End Sub

Private Sub openBtn_Click()
    Dim dialog As Office.FileDialog
    Dim strFile As String
    Static i, importResult, fileNamePosition As Integer
    Static listLen As Integer
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
     
    With dialog
        .Filters.Clear
        .Filters.Add "Все файлы", "*.*"
        .Filters.Add "ASCII Plain Text", "*.txt", 1
        .Title = "Выберите файл Zemax Prescription Data"
        .AllowMultiSelect = True
        .InitialView = msoFileDialogViewList
        '.InitialFileName = Environ("USERPROFILE") & "\Documents\"
        .InitialFileName = workingFilePath
     
        If .Show = True Then
            If .SelectedItems.Count > 4 Then
                MsgBox "Нельзя загрузить более 4 файлов. Да и зачем?", _
                    vbExclamation, "Ошибка"
            Else
                For i = 1 To .SelectedItems.Count
                    strFile = .SelectedItems(i)
                    importResult = zmxImport.zmxRaytraceImport(strFile)
                    If importResult = 0 Then
                        rayForm.status.Caption = rayForm.status.Caption & _
                            " Ошибка при импорте файла " & i
                    Else 'если всё благополучно загрузилось
                        With rayForm
                            .rayFillTableBtn.Enabled = True
                            fileNamePosition = InStrRev(dialog.SelectedItems(i), "\")
                            .filePath.text = _
                                .filePath.text & _
                                Right(dialog.SelectedItems(i), _
                                    Len(dialog.SelectedItems(i)) - fileNamePosition) & "; "
                            workingFilePath = Left(dialog.SelectedItems(i), fileNamePosition)
                        End With
                    End If
                Next i
            End If
        End If
    End With
    
    With rayForm
        '.status.Caption = "Загрузите 4 файла ZEMAX Raytrace" & vbCrLf & _
        "для апертурного (H=0, P=1), главного (1,0), верхнего (1,1) и нижнего (1,-1) лучей."
        .status.ForeColor = RGB(0, 0, 0)
        .createSheetChk.Enabled = True
        .deleteBtn.Enabled = True
        .rayFillTableBtn.Enabled = True
        .startCell.Enabled = True
        .startLabel.Enabled = True
        .headerChk.Enabled = True
        .fieldLabel.Caption = ChrW(969) & " = " _
            & Round(ArcSin(zmxImport.fieldCos) * 180 / 3.1416, 2) & ChrW(176)
    End With

    With rayForm.fileList
        For i = 0 To .ListCount - 1
            If .List(i, 3) = zmxImport.rays(0).axialRayH Then
            'если файл занесён в структуру
                .List(i, 0) = "апертурный"
            ElseIf .List(i, 3) = zmxImport.rays(0).chiefRayH Then
            'если файл занесён в структуру
                .List(i, 0) = "главный"
            ElseIf .List(i, 3) = zmxImport.rays(0).lowerRayH Then
            'если файл занесён в структуру
                .List(i, 0) = "нижний"
            ElseIf .List(i, 3) = zmxImport.rays(0).upperRayH Then
            'если файл занесён в структуру
                .List(i, 0) = "верхний"
            Else
                .List(i, 0) = ""
            End If
        Next i
    End With
    
        
    
    With rayForm
        listLen = .fileList.ListCount
        Select Case listLen
            Case 0
                .rayFillTableBtn.Enabled = False
                .status.Caption = _
                    "Сохраните в ZEMAX и загрузите сюда 4 отчёта Raytrace в текстовом варианте:" _
                        & vbCrLf & "для апертурного (Hy=0, Py=1), главного (1,0), верхнего (1,1), нижнего (1,-1) лучей."
            Case 1, 2
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "Загрузите ещё " & 4 - listLen & " файла. Должны быть загружены файлы" _
                        & vbCrLf & "для апертурного (Hy=0, Py=1), главного (1,0), верхнего (1,1), нижнего (1,-1) лучей."
            Case 3
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "Загрузите ещё один файл. Должны быть загружены файлы" _
                        & vbCrLf & "для апертурного (Hy=0, Py=1), главного (1,0), верхнего (1,1), нижнего (1,-1) лучей."
            Case 4
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "Загружено 4 файла. Можно заполнить таблицу."
            Case Else
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "Для заполнения таблицы нужно только 4 файла. Лишние можно удалить."
        End Select

    End With
    rayForm.rayFillTableBtn.SetFocus
End Sub



Private Sub rayFillTableBtn_Click()
    Call zmxImport.rayFillTable
End Sub

Private Sub sheetName_Enter()
    With rayForm
        .sheetName.ForeColor = RGB(0, 0, 0)
        .sheetName.text = ""
    End With
End Sub

Private Sub UserForm_Initialize()

    workingFilePath = Environ("USERPROFILE") & "\Documents\"
    
    With rayForm
        .status.Caption = "Загрузите 4 файла ZEMAX Raytrace" & vbCrLf & _
        "для апертурного (H=0, P=1), главного (1,0), верхнего (1,1) и нижнего (1,-1) лучей."
        .status.ForeColor = RGB(0, 0, 0)
        .createSheetChk.Enabled = False
        .deleteBtn.Enabled = False
        .rayFillTableBtn.Enabled = False
        .startCell.Enabled = False
        .sheetName.Enabled = False
        .startLabel.Enabled = False
        .headerChk.Enabled = False
        
        .fieldLabel.Caption = ""
        .startCell = "A1"
        .sheetName.text = "имя листа"
        .sheetName.ForeColor = RGB(100, 100, 100)
        
        .filePath.text = ""
       
        .openBtn.SetFocus
        
        With .fileList
            .Clear
            .ColumnHeads = False
            .ColumnCount = 4
            'тип, Hy, Py, файл
            .ColumnWidths = "50;30;30"
            .MultiSelect = fmMultiSelectMulti
        End With
    End With
End Sub
