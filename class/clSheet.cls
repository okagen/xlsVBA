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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'==================================================
'指定Sheetの指定列にAutoFilterをかける
'  [i]wb        対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [i]tgtFields Dictionaryオブジェクトで作った{field, Criteria}ハッシュテーブル
'  ***フィルターは、datRowS-1の行に設定されるためdatRowS>2でないといけない
'--------------------------------------------------
Public Function setFilter(ByVal wb As Workbook, _
                            ByVal shName As String, _
                            ByVal datRowS As Long, _
                            ByVal datColS As Long, _
                            ByVal datColE As Long, _
                            ByVal tgtFields As Variant _
                            ) As Boolean
    Dim lastRow As Long
    Dim datRng As Range
    Dim fieldKeyArr As Variant
    Dim i As Long
    
    With wb.Worksheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '最終行までのデータRangeを取得
        Set datRng = .Range(.Cells(datRowS - 1, datColS), .Cells(lastRow, datColE))
        'フィルター設定を解除
        datRng.AutoFilter
        'ハッシュテーブルのキーの配列を取得
        fieldKeyArr = tgtFields.keys
        'フィルター設定
        For i = 0 To tgtFields.count - 1
            datRng.AutoFilter field:=fieldKeyArr(i), Criteria1:=tgtFields(fieldKeyArr(i))
        Next i
    End With
    
    setFilter = True
End Function

'==================================================
'指定した文字が、指定列に存在した場合、その行を取得しArrayに格納
'シートの有無チェックは呼び出し側でやっておく
'  [i]wb        対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [i]col       検索対象列
'  [i]val       検索文字列
'  [o]dat       取得したデータ
'  [o]row       最終行数
'--------------------------------------------------
Public Function getRowDataVLookUp(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal datColS As Long, _
                                    ByVal datColE As Long, _
                                    ByVal col As Long, _
                                    ByVal val As String, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    Dim lastRow As Long
    Dim rngCol As Range
    Dim rngTmp As Range
    Dim firstAddress As String
    Dim retTmpBucket As Variant
    ReDim retTmpBucket(1 To MAX_ROW, 1 To datColE - datColS + 1)
    Dim retTmp As Variant
    Dim tgtRng As Range
    Dim datArr As New clDatArr
    Dim lastIndex As Long
    Dim bRet As Boolean
    
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        '最終行までの1列分のデータを取得
        Set rngCol = .Range(.Cells(datRowS, col), .Cells(lastRow, col))
        '検索する
        '***.Findは最初のRangeを基点にして次のRangeから検索してしまう。
        '***.Findの検索基点を最後のRangeにする。
        Set rngTmp = rngCol.Find(What:=val, LookIn:=xlValues, LookAt:=xlPart, _
                                after:=.Range(.Cells(lastRow, col), .Cells(lastRow, col)))
        '検索できた場合
        If Not rngTmp Is Nothing Then
            '最初に見つかったセルのアドレスを取得
            firstAddress = rngTmp.Address
            Dim i As Long
            i = 1
            Do
                '検索で一致した行のデータを取得
                retTmp = .Range(.Cells(rngTmp.row, datColS), _
                                .Cells(rngTmp.row, datColE)).Value
                'retTmpBucketに格納
                bRet = datArr.addArray(retTmp, i, retTmpBucket, lastIndex)

                '次を検索
                Set rngTmp = rngCol.FindNext(rngTmp)
                i = i + 1
            Loop Until rngTmp.Address = firstAddress '最初に戻ったら終了
        Else
            getRowDataVLookUp = False
            Exit Function
        End If
    End With
        
    '返り値配列を作成
    Dim recNum As Long
    Dim ret As Variant
    Dim newCol As Long
    recNum = i - 1
    newCol = datColE - datColS + 1
    bRet = datArr.formatArray(retTmpBucket, recNum, newCol, ret)
    dat = ret
    row = recNum
    getRowDataVLookUp = True
End Function

'==================================================
'指定列の最後の行までのデータを削除
'  [i]wb                対象のWorkbook
'  [i]shName            シート名
'  [i]datRowS           データ領域のスタート行
'  [i]col               ターゲットの列
'--------------------------------------------------
Public Function deleteColData(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long _
                                    ) As Boolean
    '指定列の全データ取得
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        'データ削除
        .Range(.Cells(datRowS, col), .Cells(lastRow, col)).Clear
    End With
    
    deleteColData = True
End Function

'==================================================
'指定列の最後の行までのデータを取得しArrayに格納
'シートの有無チェックは呼び出し側でやっておく
'  [i]wb                対象のWorkbook
'  [i]shName            シート名
'  [i]datRowS           データ領域のスタート行
'  [i]col               ターゲットの列
'  [i]allowDuplicate    重複を許すかどうか(true=許す)
'  [o]dat               取得データ
'  [o]row               取得データの行数
'--------------------------------------------------
Public Function getColDataAsArray(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal col As Long, _
                                    ByVal allowDuplicate As Boolean, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    '指定列の全データ取得
    Dim lastRow As Long
    Dim ret As Variant
    Dim recNum As Long
    With wb.Sheets(shName)
        '最終行取得
        lastRow = .UsedRange.Cells(.UsedRange.count).row
        'データ取得
        ret = .Range(.Cells(datRowS, col), .Cells(lastRow, col)).Value
    End With
    recNum = lastRow - datRowS + 1
    
    
    '重複削除処理
    Dim retTmp(1 To MAX_ROW, 1 To 1) As Variant
    Dim datArr As New clDatArr
    Dim bRet As Boolean
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
        bRet = datArr.formatArray(retTmp, recNum, 1, ret)
    End If
    
    '返り値設定
    If lastRow - datRowS > 0 Then
        dat = ret
        row = recNum
        getColDataAsArray = True
    Else
        getColDataAsArray = False
    End If
End Function

'==================================================
'Sheet内のデータ領域をArrayに格納する
'シートの有無チェックは呼び出し側でやっておく
'  [i]wb        対象のWorkbook
'  [i]shName    シート名
'  [i]datRowS   データ領域のスタート行
'  [i]datRowE   データ領域の最終行「0」の場合シートの最後の行まで検索する。
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [o]dat       取得したデータ
'  [o]row       最終行数
'  [o]col       列数
'--------------------------------------------------
Public Function getDataAsArray(ByVal wb As Workbook, _
                                    ByVal shName As String, _
                                    ByVal datRowS As Long, _
                                    ByVal datRowE As Long, _
                                    ByVal datColS As Long, _
                                    ByVal datColE As Long, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long, _
                                    ByRef col As Long) As Boolean
    Dim lastRow As Long
    Dim ret As Variant
    With wb.Sheets(shName)
        '最終行取得
        If datRowE = 0 Then
            lastRow = .UsedRange.Cells(.UsedRange.count).row
        Else
            lastRow = datRowE
        End If
        'データ取得
        ret = .Range(.Cells(datRowS, datColS), .Cells(lastRow, datColE)).Value
    End With
    
    '返り値設定
    If lastRow - datRowS >= 0 Then
        dat = ret
        col = datColE - datColS + 1
        row = lastRow - datRowS + 1
        getDataAsArray = True
    Else
        getDataAsArray = False
    End If
End Function

'==================================================
'名前を指定してSheetを作成。
'同名のシートが存在した場合その中身を削除する。
'  [i]wb        対象のWorkbook
'  [i]shName    作成するシート名
'--------------------------------------------------
Public Function initSheet(ByVal wb As Workbook, _
                            ByVal shName As String) As Boolean
                            
    If Not existSheet(wb, shName) Then
        wb.Worksheets.Add after:=wb.Worksheets(wb.Worksheets.count)
        ActiveSheet.name = shName
    Else
        With wb.Sheets(shName)
            Dim verLastRow As Long
            Dim delRow As String
            verLastRow = .UsedRange.Cells(.UsedRange.count).row + 1
            delRow = "1:" & verLastRow
            '行削除
            .Range(delRow).Delete
            'すべてのコントロールを削除
            .DrawingObjects.Delete
        End With
    End If

End Function

'==================================================
'名前を指定してSheetを作成。
'同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。
'  [i]wb        対象のWorkbook
'  [i]shName    作成するシート名
'--------------------------------------------------
Public Function newSheet(ByVal wb As Workbook, _
                            ByVal shName As String) As String

    Dim xlsheet As Worksheet
    Dim num As Integer
    Dim flg As Boolean
    Dim newName As String
    
    newName = shName
    num = 1
    flg = True
    Do While flg
        If Not existSheet(wb, newName) Then
            flg = False
        Else
            newName = shName & " " & Format(num, "(#)")
            num = num + 1
        End If
    Loop
    
    wb.Worksheets.Add after:=wb.Worksheets(Worksheets.count)
    ActiveSheet.name = newName
    
    newSheet = newName
    
End Function

'==================================================
'名前を指定してSheetの有無をチェック
'  [i]wb        対象のWorkbook
'  [i]shName    確認するシート名
'--------------------------------------------------
Public Function existSheet(ByVal wb As Workbook, _
                            ByVal shName As String) As Boolean
    Dim xlsheet As Worksheet
    On Error Resume Next
    Set xlsheet = wb.Worksheets(shName)
    On Error GoTo 0
    If xlsheet Is Nothing Then
        existSheet = False
    Else
        existSheet = True
    End If
    Set xlsheet = Nothing
End Function

'==================================================
'指定したRange範囲内にある図形,ActiveXコントロールを削除する
'  [i]wb            対象のWorkbook
'  [i]shName    対象のシート名
'  [i]rowUL       削除対象範囲 左上行番号
'  [i]colUL        削除対象範囲 左上列番号
'  [i]rowLR       削除対象範囲 右下行番号
'  [i]colLR        削除対象範囲 右下列番号
'--------------------------------------------------
Public Function deleteObjectInRange(ByVal wb As Workbook, _
                                                        ByVal shName As String, _
                                                        ByVal rowUL As Long, _
                                                        ByVal colUL As Long, _
                                                        ByVal rowLR As Long, _
                                                        ByVal colLR As Long _
                                                        ) As Boolean
    Dim shp As Shape
    Dim rng_shp As Range
    Dim rng As Range
 
    With wb.Worksheets(shName)
        .Select
        'Rangeをセレクト
        Set rng = .Range(.Cells(rowUL, colUL), .Cells(rowLR, colLR))
    End With
 
    'すべての図形にループ処理
    For Each shp In wb.Worksheets(shName).Shapes
        '図形の配置されているセル範囲をオブジェクト変数にセット
        Set rng_shp = Range(shp.TopLeftCell, shp.BottomRightCell)
 
        '図形の配置されているセル範囲と選択されているセル範囲が重なっていれば図形を削除
        If Not Intersect(rng_shp, rng) Is Nothing Then
            shp.Delete
        End If
    Next
    Set rng = Nothing
    deleteObjectInRange = True
End Function
