VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} jsonForm 
   Caption         =   "Импорт lensdata.json"
   ClientHeight    =   5900
   ClientLeft      =   80
   ClientTop       =   300
   ClientWidth     =   7050
   OleObjectBlob   =   "jsonForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "jsonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("OpticGOST")

Private Sub importBtn_Click()
    Dim filePath, json As String
    Dim lens As Scripting.Dictionary
    Set lens = New Scripting.Dictionary
    'filePath = jsonForm.pathBox.text
    filePath = "C:\Users\Rodin\Documents\optics\lab1_triplet\lab1\triplet_f1_v6asph2mod1_lensdata.json"
    json = hjsonParse.readTextToString(filePath)
    Set lens = hjsonParse.jsonToDict(json)
    With jsonForm
        With .outputTB
            .Visible = True
            .text = ""
            .MultiLine = True
            .ScrollBars = fmScrollBarsVertical
        End With
        .newSheetChk.Visible = True
        .newSheetName.Visible = True
        .startCell.Visible = True
        .generateTablesBtn.Visible = True
    End With
    Call hjsonParse.displayDict(lens)
End Sub

Private Sub openFileBtn_Click()
    Dim dialog As Office.FileDialog
    Dim strFile As String
     
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
     
    With dialog
        .Filters.Clear
        .Filters.Add "Все файлы", "*.*"
        .Filters.Add "HJSON Lens Data File", "*.json", 1
        .Title = "Выберите файл Zemax Prescription Data"
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
        .newSheetChk.Visible = False
        .newSheetName.Visible = False
        .startCell.Visible = False
        .outputTB.Visible = False
    
        .generateTablesBtn.Visible = False
        
        .openFileBtn.Visible = True
        
        .typeOutput_aberration.Caption = ""
        .typeOutput_objSize.Caption = ""
        .status.Caption = "Откройте файл JSON, сохранённый макросом EXPORT_JSON.zpl """
    End With
End Sub
