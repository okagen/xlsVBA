VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1




'==================================================
' 指定したシート以外のシートを削除する
'  [i]wb                対象のWorkbook
'  [i]remainSheets     確認するモジュール名
'--------------------------------------------------
Public Function deleteUnSpecifiedSheets(ByVal wb As Workbook, _
                                                                ByVal remainSheets As Variant) As Boolean

    Dim res As Variant
    Dim sh As Worksheet
    For Each sh In wb.Sheets
        '要素を検索
        res = Filter(remainSheets, sh.name)
        '戻り値配列のサイズで要素の有無を判定
        If UBound(res) = -1 Then
            sh.Delete
        End If
    Next sh
End Function

'---------------------------------------------
' ↓見直し中

'==================================================
'VLOOKUP関数を用いて、複数シートを参照してデータを取得する
'  [i]wb                   検索対象Workbook
'  [i]targetSheets    検索対象シート名のCollection
'  [i]key                   検索キー
'  [i]keyCol               検索シート内 キーが書いてある列
'  [i]dataCol         検索シート内 値の列
'  [o]ret                   検索結果
'--------------------------------------------------
Public Function getDataFromSheetsByVlookup(ByVal wb As Workbook, _
                                                                        ByRef targetSheets As Collection, _
                                                                        ByVal key As String, _
                                                                        ByVal keyCol As Long, _
                                                                        ByVal dataCol As Long, _
                                                                        ByRef ret As String _
                                                                        ) As Boolean
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim rngTxt As String
    Dim i As Long
    Dim result As Variant
    Dim rng As Range
    
    '2つの列番号をRangeを表すアルファベットに変換する
    bRet = sh.colNo2Rng(keyCol, dataCol, rngTxt)
    
    'シートを数分ループ
    For i = 1 To targetSheets.count Step 1
        With wb.Worksheets(targetSheets(i))
            On Error Resume Next
            Set rng = .Range(rngTxt)
            result = WorksheetFunction.VLookup( _
                                   key, _
                                   rng, _
                                   dataCol - keyCol + 1, _
                                   False)
            If ERR <> 0 Then result = ""
        End With
        
        '値を取得できたら関数を抜ける
        If result <> "" Then
            ret = result
            getDataFromSheetsByVlookup = True
            Exit Function
        End If
    Next i
    
    getDataFromSheetsByVlookup = False
End Function

'==================================================
'無視するシート名を引数で受けて、Book内の検索対象Sheetの名前Collectionを作成
'  [i]wb                対象のWorkbook
'  [i]ignoreSheetNames  無視するシート名が入ったCollection
'  [o]targetSheets      検索対象のシート名が入ったCollection
'--------------------------------------------------
Public Function getTargetSheets(ByVal wb As Workbook, _
                                ByVal ignoreSheetNames As Collection, _
                                ByRef targetSheets As Collection) As Boolean
                                
    Dim xlsheet As Worksheet
    Dim ignoreSheetName As Variant
    Dim bTgt As Boolean
    
    For Each xlsheet In wb.Worksheets
        
        bTgt = True
        
        For Each ignoreSheetName In ignoreSheetNames
            '対象ならば、collectionに追加
            If Left(xlsheet.name, Len(ignoreSheetName)) = ignoreSheetName Then
                bTgt = False
            End If
        Next
        
        If bTgt = True Then
            targetSheets.Add (xlsheet.name)
        End If
    Next
   
    If targetSheets.count < 1 Then
        getTargetSheets = False
    Else
        getTargetSheets = True
    End If
    
End Function

'==================================================
'複数シートの中のデータを結合して、Arrayに格納
'  [i]wb        対象のWorkbook
'  [i]shNames   検索対象のシート名が入ったCollection
'  [i]datRowS   データ領域のスタート行
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [o]dat       結合したデータ
'  [o]row       最終行数
'--------------------------------------------------
Public Function combineSheets(ByVal wb As Workbook, _
                                ByVal shNames As Collection, _
                                ByVal datRowS As Long, _
                                ByVal datColS As Long, _
                                ByVal datColE As Long, _
                                ByRef dat As Variant, _
                                ByRef row As Long) As Boolean
    Dim sh As New clSheet
    Dim datArr As New clDatArr
    Dim retTmpBucket As Variant
    ReDim retTmpBucket(1 To MAX_ROW, 1 To datColE - datColS + 1)
    Dim datBySheet As Variant
    Dim rowBySheet As Long
    Dim colBySheet As Long
    Dim bRet  As Boolean
    Dim shName As Variant
    Dim indexRow As Long
    Dim lastRow As Long
    
    indexRow = 1
    For Each shName In shNames
        'シートの有無をチェック
        If sh.existSheet(wb, shName) Then
            'シート内のDataを取得
            bRet = sh.getDataAsArray(wb, shName, datRowS, 0, datColS, datColE, datBySheet, rowBySheet, colBySheet)
            If bRet Then
                '取得したDataをbucketに追加
                bRet = datArr.addArray(datBySheet, indexRow, retTmpBucket, lastRow)
                If Not bRet Then
                    'オーバーフローする可能性があるので処理中止
                    Exit Function
                End If
                indexRow = lastRow + 1
            End If
        End If
    Next
    
    '返り値を作成
    bRet = datArr.formatArray(retTmpBucket, lastRow, colBySheet, dat)
    row = lastRow
    combineSheets = True
End Function
