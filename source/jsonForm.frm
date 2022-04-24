VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} jsonForm 
   Caption         =   "Импорт lensdata.json"
   ClientHeight    =   5844
   ClientLeft      =   84
   ClientTop       =   300
   ClientWidth     =   9780
   OleObjectBlob   =   "jsonForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "jsonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("OpticGOST")

Private Sub fieldAdd_Click()
    Call jsonDisplay.addFields(CLens)
    Call jsonDisplay.refreshFields(CLens)
End Sub

Private Sub fieldDel_Click()
    Call jsonDisplay.delFields(CLens)
    Call jsonDisplay.refreshFields(CLens)
End Sub
Private Sub generateTablesBtn_Click()
    Dim aberSheet As Excel.Worksheet
    Dim rndSheet As Excel.Worksheet
    Dim options As jsonDisplay.TOptions
    options.OPD = Me.OPDchk
    options.anamorphic = Me.anamorphicChk
    options.mRelative = Me.mRelativeChk
    options.tgSigma = Me.tgSigmaExitChk
    With Me
        If .aberTableChk.value = True Then
            Set aberSheet = Application.Worksheets.Add
            Call jsonDisplay.fillAberTable(CLens, aberSheet.Range("A1"), options)
        End If
        If .rndTableChk.value = True Then
            Set rndSheet = Application.Worksheets.Add
        End If
    End With
End Sub

Private Sub importBtn_Click()
    Dim filePath As String
    Dim json As String
    filePath = jsonForm.pathBox.text
    'filePath = "C:\Users\Rodin\Documents\optics\TV-wide70deg\retrofocus_v6\retrofocus_v6mod1_lensdata.json"
    json = hjsonParse.readTextToString(filePath)
    Call CLens.parse(json)
    With jsonForm
        .fieldFrm.Enabled = True
        .fieldAdd.Enabled = True
        .fieldDel.Enabled = True
        .waveFrm.Enabled = True
        .waveAdd.Enabled = True
        .waveDel.Enabled = True
        .generateTablesBtn.Enabled = True
        .tablesFrm.Enabled = True
        .OPDchk.Enabled = True
        .anamorphicChk.Enabled = True
        .aberTableChk.Enabled = True
        .rndTableChk.Enabled = True
    End With
    Call jsonDisplay.refreshWaves(CLens)
    Call jsonDisplay.refreshFields(CLens)
End Sub


Private Sub openFileBtn_Click()
    Dim dialog As Office.FileDialog
    Dim strFile As String
     
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
     
    With dialog
        .Filters.Clear
        .Filters.Add "Все файлы", "*.*"
        .Filters.Add "HJSON Lens Data File", "*.json", 1
        .Title = "Выберите файл JSON"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
     
        If .Show = True Then
            strFile = .SelectedItems(1)
        End If
    End With
    With jsonForm
        .pathBox.text = strFile
        .importBtn.Enabled = True
    End With
End Sub



Private Sub UserForm_Initialize()

    With jsonForm
        .fieldFrm.Enabled = False
        .fieldAdd.Enabled = False
        .fieldDel.Enabled = False
        .waveFrm.Enabled = False
        .waveAdd.Enabled = False
        .waveDel.Enabled = False
        .generateTablesBtn.Enabled = False
        .tablesFrm.Enabled = False
        
        .importBtn.Enabled = False
        
        .OPDchk.Enabled = False
        
        .status.Caption = "Откройте файл JSON, сохранённый макросом JSONexport.zpl "
    End With
End Sub

Private Sub waveAdd_Click()
    Call jsonDisplay.addWaves(CLens)
    Call jsonDisplay.refreshWaves(CLens)
End Sub

Private Sub waveDel_Click()
    Call jsonDisplay.delWaves(CLens)
    Call jsonDisplay.refreshWaves(CLens)
End Sub

Private Sub waveSel_AfterUpdate()
    Dim isAllowedWaveCount As Boolean
    isAllowedWaveCount = jsonDisplay.checkWaveCount(CLens)
    If isAllowedWaveCount Then
        Me.generateTablesBtn.Enabled = True
    Else
        If Me.aberTableChk.value = True Then
            Me.generateTablesBtn.Enabled = False
        End If
    End If
End Sub
