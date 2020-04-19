VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rndForm 
   Caption         =   "Импорт файла ZEMAX Prescription Data"
   ClientHeight    =   7710
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9390
   OleObjectBlob   =   "rndForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "rndForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const sheetNameStart As String = "введите имя листа"
Const helpText1 As String = _
    vbCrLf & vbCrLf & vbCrLf & _
    "По данным из файла Prescription Data можно построить:" & vbCrLf & _
    "- таблицу конструктивных параметров (r-n-d);" & vbCrLf & _
    "- таблицу параметров оптических деталей (диаметры/стрелки)." & vbCrLf & _
    ""
    
Option Explicit


Private Sub createSheetChk_Change()
Call UpdateStartCells
If createSheetChk.Value = True Then
    rndForm.sheetName.Enabled = True
Else
    rndForm.sheetName.Enabled = False
End If
End Sub


Private Sub fileOpenBtn_Click()
    Dim dialog As Office.FileDialog
    Dim strFile As String
     
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
     
    With dialog
        .Filters.Clear
        .Filters.Add "Все файлы", "*.*"
        .Filters.Add "ASCII Plain Text", "*.txt", 1
        .Title = "Выберите файл Zemax Prescription Data"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
     
        If .Show = True Then
            strFile = .SelectedItems(1)
        End If
    End With
    rndForm.filePath.text = strFile
    rndForm.importBtn.Enabled = True
    rndForm.importBtn.SetFocus
End Sub

Private Sub generateESKDchk_Change()
Call UpdateStartCells
End Sub


Private Sub generateZemaxTableChk_Change()
Call UpdateStartCells
End Sub

Private Sub importBtn_Click()
    rndForm.textBox.text = ""
    Call CleanUp
    rndForm.ESKDstart.text = "A1"
    Call zmxPrescriptionImport
    rndForm.textBox.SetFocus
End Sub
Private Sub lensSheetNameBox_AfterUpdate()
    With rndForm.statusLabel
            .ForeColor = RGB(0, 0, 0)
            .Caption = "Выбрана длина волны " _
            & rndForm.wavelengthList.Value & _
            vbCrLf & "Будем создан новый лист с именем" & _
            rndForm.lensSheetNameBox.text
    End With
    lensSheetNameBox.BackColor = &H80000005
End Sub

Private Sub lensSheetNameBox_Enter()
    If lensSheetNameBox.text = sheetNameStart Then
        lensSheetNameBox.text = ""
    End If
    lensSheetNameBox.BackColor = &H80000005
End Sub

Private Sub lensTableChk_Change()
With rndForm
    If .lensTableChk.Value = True Then
        .lensStart.Enabled = True
        .newLensSheetchk.Enabled = True
    Else
        .lensStart.Enabled = False
        .newLensSheetchk.Enabled = False
    End If
End With
Call UpdateStartCells
End Sub


Private Sub newLensSheetchk_Change()
    Call UpdateStartCells
    If rndForm.newLensSheetchk.Value = True Then
        rndForm.lensSheetNameBox.Enabled = True
    Else
        rndForm.lensSheetNameBox.Enabled = False
    End If
End Sub



Private Sub rndFillTableBtn_Click()
    Static i As Integer
    Static NoSelection As Boolean 'не выбрана длина волны
    
    If (rndForm.generateESKDchk.Value = True Or rndForm.generateZemaxTableChk = True) And _
    rndForm.createSheetChk.Value = True Then 'если мы создаем новый лист
    'проверим, есть ли у него имя
        If rndForm.sheetName = "" Or rndForm.sheetName = sheetNameStart Then
            With rndForm.statusLabel
                .Caption = statusLabel.Caption & vbCrLf & _
                "Введите имя листа для таблицы конструктивных параметров"
                .ForeColor = RGB(150, 0, 0)
            End With
            rndForm.sheetName.BackColor = RGB(200, 50, 50)
            Exit Sub
        End If
    'если имя есть, в основном модуле дёрнем его из sheetName
    End If
    
    If rndForm.lensTableChk.Value = True And rndForm.newLensSheetchk.Value = True Then
        If rndForm.lensSheetNameBox = "" Or rndForm.lensSheetNameBox = sheetNameStart Then
            With rndForm.statusLabel
                .Caption = statusLabel.Caption & vbCrLf & _
                "Введите имя листа для таблицы параметров оптических деталей"
                .ForeColor = RGB(150, 0, 0)
            End With
            rndForm.lensSheetNameBox.BackColor = RGB(200, 50, 50)
            Exit Sub
        End If
    End If
    
    NoSelection = True
    
    For i = 0 To rndForm.wavelengthList.ListCount - 1
        If rndForm.wavelengthList.Selected(i) = True Then
            NoSelection = False 'если хоть одна выбрана
        End If
    Next i
    
    If NoSelection = True And _
    (rndForm.generateESKDchk.Value = True Or rndForm.generateZemaxTableChk = True) _
    Then 'если не выбрана длина волны
        With rndForm.statusLabel 'ругаемся
            .Caption = "Выберите длину волны!"
            .ForeColor = RGB(255, 0, 0)
        End With
        Exit Sub
    End If
    
    'готовим данные
    Call CalculateSag
    Call glassIndexImport 'загружаем 3 длины волны: основную, длинную, короткую
    Call CalculateAbbe
    
    With rndForm
        If .lensTableChk.Value = True Then
            If .newLensSheetchk = False Then
                'если для таблицы линз не надо создавать новый лист,
                'разберёмся с ней не отходя от кассы
                Call lensFillTable
                'а теперь таблица конструктивных
                If .generateESKDchk.Value = True Or .generateZemaxTableChk = True Then
                    Call rndFillTable
                End If
            Else
                'если надо, попробуем сначала конструктивные
                'вдруг там не надо
                If .generateESKDchk.Value = True Or .generateZemaxTableChk = True Then
                    Call rndFillTable
                End If
                Call lensFillTable
            End If
        ElseIf .generateESKDchk.Value = True Or .generateZemaxTableChk = True Then
             Call rndFillTable
        Else
            Exit Sub
        End If
        
    End With
End Sub



Private Sub sheetName_AfterUpdate()
    With rndForm.statusLabel
            .ForeColor = RGB(0, 0, 0)
            .Caption = "Выбрана длина волны " _
            & rndForm.wavelengthList.Value & _
            vbCrLf & "Будет создан новый лист с именем" & _
            rndForm.sheetName.text
    End With
    sheetName.BackColor = &H80000005
End Sub

Private Sub sheetName_Enter()
    If sheetName.text = sheetNameStart Then
        sheetName.text = ""
    End If
    sheetName.BackColor = &H80000005
End Sub



Private Sub UserForm_Initialize()
    
    rndForm.createSheetChk.Value = False
    With rndForm.sheetName
        .text = sheetNameStart
        .ForeColor = RGB(0, 0, 0)
        .Enabled = False
    End With
    
    With rndForm.lensSheetNameBox
        .text = sheetNameStart
        .ForeColor = RGB(0, 0, 0)
        .Enabled = False
    End With
    
    
    With rndForm.statusLabel
            .Caption = "Выберите файл Zemax Prescription Data (.txt)"
            .ForeColor = RGB(0, 0, 0)
    End With
    
    With rndForm
        .textBox.text = helpText1
        .generateESKDchk.Value = True
        .generateZemaxTableChk.Value = False
        .lensTableChk.Value = True
        .rndFillTableBtn.Enabled = False
        .newLensSheetchk.Enabled = False
        .generateESKDchk.Enabled = False
        .generateZemaxTableChk.Enabled = False
        .lensTableChk.Enabled = False
        .lensSheetNameBox.Enabled = False
        .createSheetChk.Enabled = False
        .ZemaxStart.Enabled = False
        .ESKDstart.Enabled = False
        .lensStart.Enabled = False
        .fileOpenBtn.SetFocus
        .importBtn.Enabled = False
    End With
End Sub

Private Sub wavelengthList_AfterUpdate()
    With rndForm.statusLabel
        .ForeColor = RGB(0, 0, 0)
        .Caption = "В качестве основной выбрана длина волны " & rndForm.wavelengthList.Value
    End With
    waveSelection = CInt(Left(rndForm.wavelengthList.Value, 1))
End Sub

Private Sub UpdateStartCells()
Static surf As Integer
surf = zmxImport.surfCounter

If rndForm.generateESKDchk.Value = True Then 'если генерируем по ЕСКД
    rndForm.ESKDstart.text = "A1"
    rndForm.ZemaxStart.text = Range(rndForm.ESKDstart.text).Offset(surf * 2 + 3, 0).Address
    If rndForm.createSheetChk = True Or rndForm.newLensSheetchk = True Then
        'если констр. и линзы на разных страницах
        rndForm.lensStart.text = "A1"
    Else
        If rndForm.generateZemaxTableChk.Value = True Then
            rndForm.lensStart.text = Range(rndForm.ESKDstart.text).Offset(surf * 3 + 3, 0).Address
        Else
            rndForm.lensStart.text = Range(rndForm.ESKDstart.text).Offset(surf * 2 + 6, 0).Address
        End If
    End If
Else 'если НЕ генерируем по ЕСКД
    rndForm.ZemaxStart.text = "A1"
    If rndForm.createSheetChk = True Or rndForm.newLensSheetchk = True Then
        'если констр. и линзы на разных страницах
        rndForm.lensStart.text = "A1"
    Else
        rndForm.lensStart.text = Range("A1").Offset(surf + 3, 0).Address
    End If
End If

End Sub
