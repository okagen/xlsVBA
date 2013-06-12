VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDatArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================
'2次元配列(arr)を、(newRow, newCol)の2次元配列に整形して返す。
'値が入らない要素はEmptyになる。
'  [i]arr       整形する配列
'  [i]newRow    整形後の行数（レコード数）
'  [i]newCol    整形後の列（フィールド数）
'  [o]dat       整形後の配列
'--------------------------------------------------
Public Function formatArray(ByVal arr As Variant, _
                            ByVal newRow As Long, _
                            ByVal newCol As Long, _
                            ByRef dat As Variant _
                            ) As Boolean
                            
    '値を設定する為の、実際のループ処理数を取得
    Dim loopRow As Long
    Dim loopCol As Long
    If UBound(arr, 1) > newRow Then
        loopRow = newRow
    Else
        loopRow = UBound(arr, 1)
    End If
    If UBound(arr, 2) > newCol Then
        loopCol = newCol
    Else
        loopCol = UBound(arr, 2)
    End If
    
    '2次元配列を整形
    Dim retTmpBucket As Variant
    ReDim retTmpBucket(1 To newRow, 1 To newCol)
    Dim i As Long
    Dim j As Long
    For i = 1 To loopRow Step 1
        For j = 1 To loopCol Step 1
            retTmpBucket(i, j) = arr(i, j)
        Next j
    Next i

    dat = retTmpBucket
    formatArray = True
                            
End Function


'==================================================
'2次元配列(arr)を、2次元配列(dat)に追加して返す
'datに対して、indexRowで指定された数字を最初のレコードの添え字としてarrを追加する。
'  [i]arr        追加する配列
'  [i]indexRow    追加する際の先頭レコードの添え字
'  [i/o]dat       追加される親配列
'  [o]row         追加後の最終レコードの添え字
'--------------------------------------------------
Public Function addArray(ByVal arr As Variant, _
                            ByVal indexRow As Long, _
                            ByRef dat As Variant, _
                            ByRef row As Long) As Boolean
    '項目数（列数）が同じかチェック
    Dim col As Long
    col = UBound(dat, 2)
    If col <> UBound(arr, 2) Then
        Debug.Print "err ::: no much columns. [func:combineArray()]" & Now
        addArray = False
        Exit Function
    End If
    
    'datにarrを入れてもオーバーフローしないかチェック
    Dim rowDat As Long
    Dim rowA As Long
    rowDat = UBound(dat, 1)
    rowA = UBound(arr, 1)
    If rowDat < rowA + indexRow - 1 Then
        Debug.Print "err ::: not enough size. [func:combineArray()]" & Now
        addArray = False
        Exit Function
    End If
    
    'arrの中身をdatに入れる
    Dim i As Long
    Dim j As Long
    For i = 1 To rowA Step 1
        For j = 1 To col
            dat(indexRow + i - 1, j) = arr(i, j)
        Next j
    Next i
    row = indexRow + rowA - 1
    addArray = True
End Function

