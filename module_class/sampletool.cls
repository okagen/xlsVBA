VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sampletool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton_createPartsMaster_Click()
    Call createPartsMasterSheet(Sheets("$tool").TextBox_ConfigMasterFolder.Text, _
                                                Sheets("$tool").TextBox_PartsMasterSheet.Text)

End Sub

Private Sub CommandButton_getConfigMasterFolder_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "パーツ構成データの保存フォルダを選択"
        .InitialView = msoFileDialogViewLargeIcons
        .InitialFileName = "C:\Users\10007434\Desktop\my prj\excel_vba\"
        
        If .Show = True Then
            Sheets("$tool").TextBox_ConfigMasterFolder.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub CommandButton_getInputSheetList_Click()
    Call collectOriginalFileName(Sheets("$tool").TextBox_PartsMasterSheet.Text)
        
End Sub

Private Sub Worksheet_Activate()

    '配布用シートに含める列を選択するためのリストボックスの初期設定
    Dim lstBx As MSForms.ListBox
    Set lstBx = Sheets("$tool").ListBox_addColumn
    With lstBx
        If .ListCount < 1 Then
            .AddItem ("UGL備考")
            .AddItem ("UGL変更履歴")
            .AddItem ("UGL販売価格")
            .AddItem ("UGL管理No")
        End If
    End With



End Sub

