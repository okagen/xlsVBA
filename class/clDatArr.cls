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
'Collectionの中身を2次元配列に入れなおす。
'  [i]coll    Collection
'  [i]isRow 行用か列用 true=行(1,1)(1,2)・・・
'                               false=列(1,1)(2,1)・・・
'  [o]arr    変換後の配列
'--------------------------------------------------
Public Function cnvCollToArr(ByVal coll As Collection, _
                                            ByVal isRow As Boolean, _
                                            ByRef arr As Variant _
                                            ) As Boolean
    Dim retArr As Variant
    Dim i As Long
    
    If isRow = True Then
        ReDim retArr(1 To 1, 1 To coll.count)
        For i = 1 To coll.count Step 1
            retArr(1, i) = coll(i)
        Next i
    Else
        ReDim retArr(1 To coll.count, 1 To 1)
        For i = 1 To coll.count Step 1
            retArr(i, 1) = coll(i)
        Next i
    End If
    
    arr = retArr
    cnvCollToArr = True
End Function

'==================================================
'2次元配列を1行ずつCollectionに入れなおす。
'  [i]arr      対象の配列
'  [o]coll    変換後のCollection
'--------------------------------------------------
Public Function cnvArrToColl(ByVal arr As Variant, _
                                            ByRef coll As Collection _
                                            ) As Boolean
    Dim bRet As Boolean
    Dim dimension As Long
    Dim orgRow As Long
    Dim orgCol As Long
    Dim retColl As New Collection
    Dim i As Long
    Dim j As Long
    
    '配列の次元数を取得 2次元配列でない場合はエラー
    bRet = getDimension(arr, dimension)
    If bRet = False Or dimension <> 2 Then
        Debug.Print "err ::: dimension ->" & dimension & " [func:cnvArrToColl()]" & Now
        cnvArrToColl = False
        Exit Function
    End If
    
    orgRow = UBound(arr, 1)
    orgCol = UBound(arr, 2)
    
    For i = 1 To orgRow Step 1
        For j = 1 To orgCol Step 1
            retColl.Add (arr(i, j))
        Next j
    Next i
    
    Set coll = retColl
    cnvArrToColl = True
End Function

'==================================================
'2次元配列の中に、同じレコード(行)が存在した場合、一つを残して他のレコードを削除する処理を追加する。
'  [i]arr           対象の配列
'  [o]newArr    行の重複削除後の配列
'  [i]ignoreCol レコードの重複をチェックす際に無視する要素(列)の添え字
'--------------------------------------------------
Public Function removeDuplication(ByVal arr As Variant, _
                                                    ByRef newArr As Variant, _
                                                    Optional ByVal ignoreCol As Long = 0 _
                                                    ) As Boolean
    Dim bRet As Boolean
    Dim dimension As Long
    Dim orgRow As Long
    Dim orgCol As Long
    Dim tmpArr As Variant
    Dim dic As Variant
    Dim i As Long
    Dim j As Long
    Dim tmpRow As Long
    Dim word As String
    Dim oneLine As String
                                                    
    '配列の次元数を取得 2次元配列でない場合はエラー
    bRet = getDimension(arr, dimension)
    If bRet = False Or dimension <> 2 Then
        Debug.Print "err ::: dimension ->" & dimension & " [func:removeDuplication()]" & Now
        removeDuplication = False
        Exit Function
    End If
    
    '返り値用の配列の領域を生成
    orgRow = UBound(arr, 1)
    orgCol = UBound(arr, 2)
    ReDim tmpArr(1 To orgRow, 1 To orgCol)
    
    '重複するレコードを無視して、新しいArrayを作成する
    Set dic = CreateObject("Scripting.Dictionary")
    tmpRow = 1
    For i = 1 To orgRow Step 1
        'Empty地は" "に変換し、レコードを１文字列に変換。要素は"-"でつなぐ
        oneLine = ""
        For j = 1 To orgCol Step 1
            'ignoreColと一致した場合は無視する
            If ignoreCol <> j Then
                If IsEmpty(arr(i, j)) Then
                    word = " "
                Else
                    word = arr(i, j)
                End If
                oneLine = oneLine & "-" & word
            End If
        Next j
        
        'Dictionaryオブジェクトを使う。同じ値がなければ返り値用Arrayに追加
        If Not dic.Exists(oneLine) Then
            dic.Add oneLine, oneLine
            For j = 1 To orgCol Step 1
                tmpArr(tmpRow, j) = arr(i, j)
            Next j
            tmpRow = tmpRow + 1
        End If
    Next i
    
    '不要なレコードを削除
    bRet = formatArray(tmpArr, tmpRow - 1, orgCol, tmpArr)
    If bRet = False Then
        Debug.Print "err ::: cannot format the array. [func:removeDuplication()]" & Now
        removeDuplication = False
        Exit Function
    End If
    
    newArr = tmpArr
    removeDuplication = True
End Function

'==================================================
'2次元配列で、あるレコード(行)のすべての要素(列)がEmptlyの場合、削除する
'  [i]arr           対象の配列
'  [o]newArr         Emptyレコード削除後の配列
'--------------------
'メソッド事態が返す値を bool -> integer に変更 2017.2.15
' 0=処理エラー
' 1=処理成功
' 2=一つも
'--------------------------------------------------
Public Function removeEmptyRecord(ByVal arr As Variant, _
                                    ByRef newArr As Variant _
                                ) As Integer
    Dim bRet As Boolean
    Dim dimension As Long
    Dim tmpArr As Variant
    Dim orgRow As Long
    Dim orgCol As Long
    Dim emptyFlg As Boolean
    Dim i As Long
    Dim j As Long
    Dim tmpRow As Long
    Dim allEmptyFLg As Boolean
    allEmptyFLg = True
    
    
    '配列の次元数を取得 2次元配列でない場合はエラー
    bRet = getDimension(arr, dimension)
    If bRet = False Or dimension <> 2 Then
        Debug.Print "err ::: dimension ->" & dimension & " [func:removeEmptyRecord()]" & Now
        removeEmptyRecord = 0
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
            allEmptyFLg = False
        End If
    Next i
    
    If allEmptyFLg = False Then
        '不要なレコードを削除
        bRet = formatArray(tmpArr, tmpRow - 1, orgCol, tmpArr)
        If bRet = False Then
            Debug.Print "err ::: cannot format the array. [func:removeEmptyRecord()]" & Now
            removeEmptyRecord = 0
            Exit Function
        End If
        
        newArr = tmpArr
        removeEmptyRecord = 1
        Exit Function
    Else
        '全てのレコードがEmptyの場合
        removeEmptyRecord = 2
        Exit Function
    End If

End Function

'==================================================
'2次元配列(arr)の指定列を削除。処理後の2次元配列は1列減る。
'  [i]arr           削除元配列
'  [i]colIndex  削除列の添え字
'  [o]newArr   削除後の配列
'--------------------------------------------------
Public Function removeColFromArray(ByVal arr As Variant, _
                                                        ByVal colIndex As Long, _
                                                        ByRef newArr As Variant _
                                                        ) As Boolean
    Dim bRet As Boolean
    Dim dimension As Long
    Dim retArr As Variant
    Dim i As Long
    Dim j As Long

     '配列の次元数を取得 2次元配列でない場合はエラー
    bRet = getDimension(arr, dimension)
    If bRet = False Or dimension <> 2 Then
        Debug.Print "err ::: dimension ->" & dimension & " [func:removeColFromArray()]" & Now
        removeColFromArray = False
        Exit Function
    End If
    
    'arrの列数よりも削除列の指定添え字が大きい場合エラー
    If colIndex > UBound(arr, 2) Then
        Debug.Print "err ::: colIndex(" & colIndex & ") is too big [func:removeColFromArray()] " & Now
        removeColFromArray = False
        Exit Function
    ElseIf colIndex < 1 Then
        Debug.Print "err ::: colIndex(" & colIndex & ") is not avairable [func:removeColFromArray()] " & Now
        removeColFromArray = False
        Exit Function
    End If
     
    '返り値用の配列を作成
    ReDim retArr(1 To UBound(arr, 1), 1 To UBound(arr, 2) - 1)
    
    '返り値用の配列にデータを設定
    For i = 1 To UBound(arr, 1) Step 1
        For j = 1 To UBound(arr, 2) Step 1
            If j = colIndex Then
                'なにもしない
            ElseIf j < colIndex Then
                retArr(i, j) = arr(i, j)
            ElseIf j > colIndex Then
                retArr(i, j - 1) = arr(i, j)
            End If
        Next j
    Next i
    
    newArr = retArr
    removeColFromArray = True
End Function

'==================================================
'2次元配列(arr)の指定列に、1列挿入し値を埋める。処理後の2次元配列は1列増える。
'  [i]arr           列を追加する配列
'  [i]colIndex  挿入列の添え字
'  [i]colVal    列の値
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
    Do While ERR.Number = 0
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
'indexRowで指定された数字をarr最初のレコードの添え字としてdatに追加する。
'  [i]arr        追加する配列
'  [i]indexRow    追加する際の先頭レコードの添え字
'  [i/o]dat       追加される親配列
'  [o]row         追加後の最終レコードの添え字
'--------------------------------------------------
Public Function addArray(ByVal arr As Variant, _
                            ByVal indexRow As Long, _
                            ByRef dat As Variant, _
                            ByRef row As Long) As Boolean
    Dim rowArr As Long
    Dim colArr As Long
    Dim rowDat As Long
    Dim colDat As Long
    
    rowArr = UBound(arr, 1)
    colArr = UBound(arr, 2)
    rowDat = UBound(dat, 1)
    colDat = UBound(dat, 2)
    
    'datにarrを入れてもオーバーフローしないかチェック
    If rowDat < rowArr + indexRow - 1 Then
        Debug.Print "err ::: not enough ROW size. [func:combineArray()]" & Now
        addArray = False
        Exit Function
    End If
    If colDat < colArr Then
        Debug.Print "err ::: not enough COLUMN size. [func:combineArray()]" & Now
        addArray = False
        Exit Function
    End If
    
    'arrの中身をdatに入れる
    Dim i As Long
    Dim j As Long
    For i = 1 To rowArr Step 1
        For j = 1 To colArr
            dat(indexRow + i - 1, j) = arr(i, j)
        Next j
    Next i
    row = indexRow + rowArr - 1
    addArray = True
End Function



'==================================================
'2つの2次元配列を行(row)方向に足し合わせる
'  [i]arr1        配列1
'  [i]arr2        配列2
'  [i/o]dat     足した後の配列
'  [o]row        足した後の最終レコードの添え字
'--------------------------------------------------
Public Function add2Arrays(ByVal arr1 As Variant, _
                                            ByVal arr2 As Variant, _
                                            ByRef dat As Variant, _
                                            ByRef row As Long) As Boolean
    Dim rowArr1, colArr1, rowArr2, colArr2, newRow, newCol As Long
    Dim datTmp As Variant
    
    rowArr1 = UBound(arr1, 1)
    colArr1 = UBound(arr1, 2)
    rowArr2 = UBound(arr2, 1)
    colArr2 = UBound(arr2, 2)
    
    newRow = rowArr1 + rowArr2
    If colArr1 > colArr2 Then
      newCol = colArr1
    Else
      newCol = colArr2
    End If
    ReDim datTmp(1 To newRow, 1 To newCol)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To rowArr1 Step 1
        For j = 1 To colArr1
            datTmp(i, j) = arr1(i, j)
        Next j
    Next i
    
    For i = 1 To rowArr2 Step 1
        For j = 1 To colArr2
            datTmp(rowArr1 + i, j) = arr2(i, j)
        Next j
    Next i
    
    dat = datTmp
    row = newRow
End Function


