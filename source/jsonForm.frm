VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} jsonForm 
   Caption         =   "Импорт lensdata.json"
   ClientHeight    =   5840
   ClientLeft      =   80
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
    With jsonForm
        Call jsonDisplay.copyListboxItem(.fieldList, .fieldSel)
    End With
End Sub

Private Sub fieldDel_Click()
    Call jsonDisplay.delListboxItem(Me.fieldSel)
End Sub

Private Sub fieldList_Click()

End Sub

Private Sub importBtn_Click()
    Dim filePath, json As String
    Dim lens As Scripting.Dictionary
    Set lens = New Scripting.Dictionary
    'filePath = jsonForm.pathBox.text
    filePath = "C:\Users\Rodin\Documents\optics\TV-wide70deg\retrofocus_v6\retrofocus_v6mod1_lensdata.json"
    json = hjsonParse.readTextToString(filePath)
    Set lens = hjsonParse.jsonToDict(json)
    With jsonForm
        .fieldFrm.Enabled = True
        .fieldAdd.Enabled = True
        .fieldDel.Enabled = True
        .waveFrm.Enabled = True
        .waveAdd.Enabled = True
        .waveDel.Enabled = True
        .generateTablesBtn.Enabled = True
    End With
    Call jsonDisplay.dispWaves(lens)
    Call jsonDisplay.dispFields(lens)
End Sub

Private Sub Label1_Click()

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
        
        .OPDchk.Enabled = False
        
        .status.Caption = "Откройте файл JSON, сохранённый макросом JSONexport.zpl "
    End With
End Sub

Private Sub waveAdd_Click()
    With jsonForm
        Call jsonDisplay.copyListboxItem(.waveList, .waveSel)
    End With
End Sub

Private Sub waveDel_Click()
    Call jsonDisplay.delListboxItem(Me.waveSel)
End Sub
