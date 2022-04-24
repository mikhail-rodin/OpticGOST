VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rayForm 
   Caption         =   "������ ������ ZEMAX Raytrace"
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
Public workingFilePath As String '��� �������� ��������� ��������� ����������


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
        listLen = .ListCount '�����������, �.�. �� ����� ���������� � ���� ����� ����������
        While i <= listLen - 1
            If .Selected(i) = True Then
                '��� ��� ���������� ����� ������ ������� � �������� ������ �����
                '����� ��� �����
                tmpStr = .List(i, 0)
                If tmpStr = "�������" Then '���� ���� ������������
                    zmxImport.rays(0).chiefRayH = ""
                ElseIf tmpStr = "�������" Then '���� ���� ������������
                    zmxImport.rays(0).upperRayH = ""
                ElseIf tmpStr = "������" Then '���� ���� ������������
                    zmxImport.rays(0).lowerRayH = ""
                ElseIf tmpStr = "����������" Then '���� ���� ������������
                    zmxImport.rays(0).axialRayH = ""
                End If
                'If tmpStr = "�������" Then '���� ���� ������������
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
'                    "��������� � ZEMAX � ��������� ���� 4 ������ Raytrace � ��������� ��������:" _
'                        & vbCrLf & "��� ����������� (Hy=0, Py=1), �������� (1,0), �������� (1,1), ������� (1,-1) �����."
'            Case 1, 2
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "��������� ��� " & 4 - listLen & "�����. ������ ���� ��������� �����" _
'                        & vbCrLf & "��� ����������� (Hy=0, Py=1), �������� (1,0), �������� (1,1), ������� (1,-1) �����."
'            Case 3
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "��������� ��� ���� ����. ������ ���� ��������� �����" _
'                        & vbCrLf & "��� ����������� (Hy=0, Py=1), �������� (1,0), �������� (1,1), ������� (1,-1) �����."
'            Case 4
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "��������� 4 �����. ����� ��������� �������."
'            Case Else
'                .rayFillTableBtn.Enabled = True
'                .createSheetChk.Enabled = True
'                .startCell.Enabled = True
'                .sheetName.Enabled = True
'                .deleteBtn.Enabled = True
'                .startLabel.Enabled = True
'                .headerChk.Enabled = True
'                .status.Caption = _
'                    "��� ���������� ������� ����� ������ 4 �����. ������ ����� �������."
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
        .Filters.Add "��� �����", "*.*"
        .Filters.Add "ASCII Plain Text", "*.txt", 1
        .Title = "�������� ���� Zemax Prescription Data"
        .AllowMultiSelect = True
        .InitialView = msoFileDialogViewList
        '.InitialFileName = Environ("USERPROFILE") & "\Documents\"
        .InitialFileName = workingFilePath
     
        If .Show = True Then
            If .SelectedItems.Count > 4 Then
                MsgBox "������ ��������� ����� 4 ������. �� � �����?", _
                    vbExclamation, "������"
            Else
                For i = 1 To .SelectedItems.Count
                    strFile = .SelectedItems(i)
                    importResult = zmxImport.zmxRaytraceImport(strFile)
                    If importResult = 0 Then
                        rayForm.status.Caption = rayForm.status.Caption & _
                            " ������ ��� ������� ����� " & i
                    Else '���� �� ������������ �����������
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
        '.status.Caption = "��������� 4 ����� ZEMAX Raytrace" & vbCrLf & _
        "��� ����������� (H=0, P=1), �������� (1,0), �������� (1,1) � ������� (1,-1) �����."
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
            '���� ���� ������ � ���������
                .List(i, 0) = "����������"
            ElseIf .List(i, 3) = zmxImport.rays(0).chiefRayH Then
            '���� ���� ������ � ���������
                .List(i, 0) = "�������"
            ElseIf .List(i, 3) = zmxImport.rays(0).lowerRayH Then
            '���� ���� ������ � ���������
                .List(i, 0) = "������"
            ElseIf .List(i, 3) = zmxImport.rays(0).upperRayH Then
            '���� ���� ������ � ���������
                .List(i, 0) = "�������"
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
                    "��������� � ZEMAX � ��������� ���� 4 ������ Raytrace � ��������� ��������:" _
                        & vbCrLf & "��� ����������� (Hy=0, Py=1), �������� (1,0), �������� (1,1), ������� (1,-1) �����."
            Case 1, 2
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "��������� ��� " & 4 - listLen & " �����. ������ ���� ��������� �����" _
                        & vbCrLf & "��� ����������� (Hy=0, Py=1), �������� (1,0), �������� (1,1), ������� (1,-1) �����."
            Case 3
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "��������� ��� ���� ����. ������ ���� ��������� �����" _
                        & vbCrLf & "��� ����������� (Hy=0, Py=1), �������� (1,0), �������� (1,1), ������� (1,-1) �����."
            Case 4
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "��������� 4 �����. ����� ��������� �������."
            Case Else
                .rayFillTableBtn.Enabled = True
                .createSheetChk.Enabled = True
                .startCell.Enabled = True
                .sheetName.Enabled = True
                .deleteBtn.Enabled = True
                .startLabel.Enabled = True
                .headerChk.Enabled = True
                .status.Caption = _
                    "��� ���������� ������� ����� ������ 4 �����. ������ ����� �������."
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
        .status.Caption = "��������� 4 ����� ZEMAX Raytrace" & vbCrLf & _
        "��� ����������� (H=0, P=1), �������� (1,0), �������� (1,1) � ������� (1,-1) �����."
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
        .sheetName.text = "��� �����"
        .sheetName.ForeColor = RGB(100, 100, 100)
        
        .filePath.text = ""
       
        .openBtn.SetFocus
        
        With .fileList
            .Clear
            .ColumnHeads = False
            .ColumnCount = 4
            '���, Hy, Py, ����
            .ColumnWidths = "50;30;30"
            .MultiSelect = fmMultiSelectMulti
        End With
    End With
End Sub
