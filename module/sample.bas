Attribute VB_Name = "sample"
Option Explicit

Sub createPartsMaster()
    
    Dim bRet As Boolean
    Dim fls As New clFiles
    
    '指定フォルダ内のパーツ構成マスターファイルを取得
    Dim rootPath As String
    Dim files As New Collection
    rootPath = ThisWorkbook.path & "\sample_config_master"
    bRet = fls.getAllXlsFilePathCol(rootPath, files)
    If Not bRet Then
        Exit Sub
    End If
    
    '一つ一つ開いてデータを取得
    Dim shs As New clSheets
    Dim file As Variant
    Dim wb As Workbook
    Dim ignoreNames As New Collection
    Dim targetNames As New Collection
    Dim shName As Variant
    '=======================
    'Sheet names to ignore
    ignoreNames.Add ("tool")
    ignoreNames.Add ("$")
    ignoreNames.Add ("ugl-")
    '=======================
    
    Application.ScreenUpdating = False
    
    For Each file In files
        'workbookオブジェクトを取得
        bRet = fls.getWorkbookObj(file, wb)
        
        '検索対象のシート名を取得
        Set targetNames = New Collection
        bRet = shs.getTargetSheets(wb, ignoreNames, targetNames)
        
        
        If bRet Then
        
        
        
            'オブジェクトを閉じる
            wb.Close savechanges:=False
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub
