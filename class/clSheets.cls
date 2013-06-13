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
'無視するシート名を引数で受けて、Book内の検索対象Sheetの名前Collectionを作成
'  [i]ignoreSheetNames   無視するシート名が入ったCollection
'  [o]targetSheets       検索対象のシート名が入ったCollection
'--------------------------------------------------
Public Function getTargetSheets(ByVal ignoreSheetNames As Collection, _
                                ByRef targetSheets As Collection) As Boolean
                                
    Dim xlsheet As Worksheet
    Dim ignoreSheetName As Variant
    Dim bTgt As Boolean
    
    For Each xlsheet In Worksheets
        
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
   
    If targetSheets.Count < 1 Then
        getTargetSheets = False
    Else
        getTargetSheets = True
    End If
    
End Function

'==================================================
'複数シートの中のデータを結合して、Arrayに格納
'  [i]shNames   検索対象のシート名が入ったCollection
'  [i]datRowS   データ領域のスタート行
'  [i]datColS   データ領域の最初の列（一番左の列)
'  [i]datColE   データ領域の最後の列（一番右の列)
'  [o]dat       結合したデータ
'  [o]row       最終行数
'--------------------------------------------------
Public Function combineSheets(ByVal shNames As Collection, _
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
        If sh.existSheet(shName) Then
            'シート内のDataを取得
            bRet = sh.getAllDataAsArray(shName, datRowS, datColS, datColE, datBySheet, rowBySheet, colBySheet)
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
