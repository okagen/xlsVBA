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
'2次元配列で、あるレコード(行)のすべての要素(列)がEmptlyの場合、削除する
'  [i]arr           対象の配列
'  [o]newArr         Emptyレコード削除後の配列
'--------------------------------------------------
Public Function removeEmptyRecord(ByVal arr As Variant, _
                                                            ByRef newArr As Variant _
                                                            ) As Boolean
    Dim bRet As Boolean
    Dim dimension As Long
    Dim tmpArr As Variant
    Dim orgRow As Long
    Dim orgCol As Long
    Dim emptyFlg As Boolean
    Dim i As Long
    Dim j As Long
    Dim tmpRow As Long
    
    '配列の次元数を取得 2次元配列でない場合はエラー
    bRet = getDimension(arr, dimension)
    If bRet = False Or dimension <> 2 Then
        Debug.Print "err ::: dimension ->" & dimension & " [func:removeEmptyRecord()]" & Now
        removeEmptyRecord = False
        Exit Function
    End If
    
    '返り値用の配列の領域を生成
    orgRow = UBound(arr, 1)
    orgCol = UBound(arr, 2)
    ReDim tmpArr(1 To orgRow, 1 To orgCol)
    
    '返り値用の配列にデータを設定
    tmpRow = 1
    For i = 1 To orgRow Step 1
        '各要素に値が設定されているかをチェック
        emptyFlg = True
        For j = 1 To orgCol Step 1
            If Not IsEmpty(arr(i, j)) Then
                emptyFlg = False
            End If
        Next j
        
        '値が一つでも入っていた場合は返り値用の配列にコピー
        If Not emptyFlg Then
            For j = 1 To orgCol Step 1
                tmpArr(tmpRow, j) = arr(i, j)
            Next j
            tmpRow = tmpRow + 1
        End If
    Next i
    
    '不要なレコードを削除
    bRet = formatArray(tmpArr, tmpRow - 1, orgCol, tmpArr)
    If bRet = False Then
        Debug.Print "err ::: cannot format the array. [func:removeEmptyRecord()]" & Now
        removeEmptyRecord = False
        Exit Function
    End If
    
    newArr = tmpArr
    removeEmptyRecord = True
End Function


'==================================================
'2次元配列(arr)の指定列に、1列挿入し値を埋める。処理後の2次元配列は1列増える。
'  [i]arr           列を追加する配列
'  [i]newRow  挿入列の添え字
'  [i]newCol    列の値
'  [o]newArr   追加後の配列
'--------------------------------------------------
Public Function insertColIntoArray(ByVal arr As Variant, _
                                                        ByVal colIndex As Long, _
                                                        ByVal colVal As String, _
                                                        ByRef newArr As Variant _
                                                        ) As Boolean
    Dim bRet As Boolean
    Dim dimension As Long
    Dim retArr As Variant
    Dim newRow As Long
    Dim newCol As Long

    '配列の次元数を取得 2次元配列でない場合はエラー
    bRet = getDimension(arr, dimension)
    If bRet = False Or dimension <> 2 Then
        Debug.Print "err ::: dimension ->" & dimension & " [func:insertColIntoArray()]" & Now
        insertColIntoArray = False
        Exit Function
    End If
    
     'arrの列数+1よりも挿入列の指定添え字が大きい場合エラー
    If colIndex > UBound(arr, 2) + 1 Then
        Debug.Print "err ::: colIndex(" & colIndex & ") is too big [func:insertColIntoArray()] " & Now
        insertColIntoArray = False
        Exit Function
    End If
     
    '返り値用の配列を作成
    ReDim retArr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + 1)
    
    '返り値用の配列にデータを設定
    For newRow = 1 To UBound(arr, 1) Step 1
        For newCol = 1 To UBound(arr, 2) + 1 Step 1
            If newCol = colIndex Then
                retArr(newRow, newCol) = colVal
            ElseIf newCol > colIndex Then
                retArr(newRow, newCol) = arr(newRow, newCol - 1)
            Else
                retArr(newRow, newCol) = arr(newRow, newCol)
            End If
        Next newCol
    Next newRow
    
    newArr = retArr
    insertColIntoArray = True
End Function

'==================================================
'配列(arr)の次元数を求める
'  [i]arr           列を追加する配列
'  [o]dimension         追加後の配列
'--------------------------------------------------
Private Function getDimension(ByVal arr As Variant, _
                                                    ByRef retDim As Long _
                                                    ) As Boolean
    Dim TempData As Variant
    Dim i As Long

    On Error Resume Next
    Do While Err.Number = 0
        i = i + 1
        TempData = UBound(arr, i)
    Loop
    On Error GoTo 0
    
    retDim = i - 1
    getDimension = True
End Function

'==================================================
'2次元配列(arr)を、(newRow, newCol)の2次元配列に整形して返す。
'値が入らない要素はEmptyになる。
'  [i]arr       整形する配列
'  [i]colIndex    整形後の行数（レコード数）
'  [i]colVal    整形後の列（フィールド数）
'  [o]newArr       整形後の配列
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

