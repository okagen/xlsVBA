VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum shCond

    'データの領域はA6からスタート
    data_start_col = 1
    data_start_row = 6
    
    'データの領域はG（7列）まで
    data_last_col = 7

End Enum


'==================================================
'指定列の最後の行までのデータを取得しArrayに格納
'--------------------------------------------------
Public Function getColDataAsArray(ByVal shName As String, _
                                    ByVal col As Long, _
                                    ByVal allowDuplicate As Boolean, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    'シートの有無をチェック
    Dim xWsheet As Worksheet
    On Error Resume Next
    Set xWsheet = Worksheets(shName)
    On Error GoTo 0
    If xWsheet Is Nothing Then
        getColDataAsArray = False
        Exit Function
    End If
    
    '指定列の全データ取得
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.Count).row
        'データ取得
        ret = .Range(.Cells(shCond.data_start_row, col), _
                        .Cells(lastRow, col)).Value
    End With
    recNum = lastRow - data_start_row + 1
    
    
    '重複削除処理
    Dim retTmp(1 To 6000, 1 To 1) As Variant
    If allowDuplicate = False Then
        Dim dic As Variant
        Dim i As Long
        Dim j As Long

        j = 1
        Set dic = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(ret) Step 1
            'dic配列の中に未登録の場合、追加する。
            If Not dic.Exists(ret(i, 1)) Then
                dic.Add ret(i, 1), ret(i, 1)
                retTmp(j, 1) = ret(i, 1)
                j = j + 1
            End If
        Next i
        
        '返り値配列を作成
        recNum = j - 1
        ReDim ret(1 To recNum, 1 To 1)
        For i = 1 To recNum Step 1
            ret(i, 1) = retTmp(i, 1)
        Next i
    End If
    
    '返り値設定
    If lastRow - data_start_row > 0 Then
        dat = ret
        row = recNum
        getColDataAsArray = True
    Else
        getColDataAsArray = False
    End If
End Function

'==================================================
'Sheet内のデータ領域をArrayに格納する
'--------------------------------------------------
Public Function getAllDataAsArray(ByVal shName As String, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long, _
                                    ByRef col As Long) As Boolean
    'シートの有無をチェック
    Dim xWsheet As Worksheet
    On Error Resume Next
    Set xWsheet = Worksheets(shName)
    On Error GoTo 0
    If xWsheet Is Nothing Then
        getAllDataAsArray = False
        Exit Function
    End If
                            
    Dim lastRow As Long
    Dim ret As Variant
    With Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.Count).row
        'データ取得
        ret = .Range(.Cells(shCond.data_start_row, shCond.data_start_col), _
                        .Cells(lastRow, shCond.data_last_col)).Value
    End With
    
    '返り値設定
    If lastRow - data_start_row > 0 Then
        dat = ret
        col = shCond.data_last_col - shCond.data_start_col + 1
        row = lastRow - data_start_row + 1
        getAllDataAsArray = True
    Else
        getAllDataAsArray = False
    End If
End Function

'==================================================
'名前を指定してSheetを作成。
'同名のシートが存在した場合その中身を削除する。
'--------------------------------------------------
Public Function initSheet(ByVal shName As String)
    Dim xWsheet As Worksheet
    
    On Error Resume Next
    Set xWsheet = Worksheets(shName)
    On Error GoTo 0
        
    If xWsheet Is Nothing Then
        Worksheets.Add after:=Worksheets(Worksheets.Count)
        ActiveSheet.name = shName
    Else
        With Sheets(shName)
            Dim verLastRow As Long
            Dim delRow As String
            verLastRow = .UsedRange.Cells(.UsedRange.Count).row + 1
            delRow = "1:" & verLastRow
            .Range(delRow).Delete
        End With
    End If

End Function




