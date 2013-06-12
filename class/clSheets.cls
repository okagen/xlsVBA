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

'==================================================
'複数シートの中身を結合して、Arrayに格納
'複数シートのデータ部分(Enum shCondで指定している)を結合し、Arrayに格納する。
'  [i]shNames   検索対象のシート名が入ったCollection
'  [o]dat       結合したデータ
'  [o]row       最終行数
'--------------------------------------------------
Public Function combineSheets(ByVal shNames As Collection, _
                                    ByRef dat As Variant, _
                                    ByRef row As Long) As Boolean
    Dim sh As New clSheet
    Dim datArr As New clDatArr
    Dim retTmpBucket(1 To shCond.data_max_row, 1 To shCond.data_last_col - shCond.data_start_col + 1) As Variant
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
        If sh.existSheet(shName) Then
            'シート内のDataを取得
            bRet = sh.getAllDataAsArray(shName, datBySheet, rowBySheet, colBySheet)
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
