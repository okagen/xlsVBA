Attribute VB_Name = "verify"
Option Explicit

'最大行番号
Public Const MAX_ROW = 65536

Enum shCond

    'データの領域はA6からスタート
    datRowS = 6
    datColS = 1
    
    'データの領域はG（7列）まで
    datColE = 7

End Enum

'*** clSheets内メソッド ***
'==================================================
Sub verify_clSheets_getTargetSheets()
    Dim shs As New clSheets
    Dim ignoreNames As New Collection
    Dim targetNames As New Collection
    Dim bRet As Boolean
    Dim name As Variant
    
    '=======================
    'Sheet names to ignore
    ignoreNames.Add ("tool")
    ignoreNames.Add ("$")
    ignoreNames.Add ("ugl-")
    '=======================
    
    'get target sheet names
    bRet = shs.getTargetSheets(ignoreNames, targetNames)
    
    For Each name In targetNames
        Debug.Print "result ::: done " & name & " |" & Now
    Next
    
End Sub

'==================================================
Sub verify_clSheets_conbineSheets()
    Dim names As New Collection
    Dim sh As New clSheet
    Dim shs As New clSheets
    Dim bRet As Boolean
    Dim dat As Variant
    Dim row As Long
    Dim arrRow As Long
    Dim arrCol As Long
    
    '=======================
    'The Sheet names to test
    names.Add ("sample1")
    names.Add ("sample2")
    names.Add ("sample3")
    '=======================
    
    'get data in sheets
    bRet = shs.combineSheets(names, shCond.datRowS, shCond.datColS, shCond.datColE, dat, row)
    
    'get number of row and column from array object.
    arrRow = UBound(dat, 1)
    arrCol = UBound(dat, 2)
    
    If bRet = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(arrRow, arrCol)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub



'*** clSheet内メソッド ***
'==================================================
Sub verify_clSheet_setFilter()
    Dim name As String
    Dim sh As New clSheet
    Dim tgtFields As Variant
    Set tgtFields = CreateObject("Scripting.Dictionary")
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    tgtFields.Add 1, "F45N"
    tgtFields.Add 2, "*Table*"
    tgtFields.Add 3, "13"
    '=======================
    
    'set filter on the sheet
    sh.setFilter name, shCond.datRowS, shCond.datColS, shCond.datColE, tgtFields
    
End Sub


'==================================================
Sub verify_clSheet_getRowDataVLookUp()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim ret As Boolean
    Dim str As String
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 1
    str = "45"
    '=======================
    
    'get data in the sheet
    ret = sh.getRowDataVLookUp(name, shCond.datRowS, shCond.datColS, shCond.datColE, col, str, dat, row)
    
    If ret = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, 7)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clSheet_getColDataAsArray()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim ret As Boolean
    Dim allowDup As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 1
    allowDup = False
    '=======================
    
    'get data in the sheet
    ret = sh.getColDataAsArray(name, shCond.datRowS, col, allowDup, dat, row)
    
    If ret = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, 1)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clSheet_getAllDataAsArray()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim ret As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    '=======================
    
    'get data in the sheet
    ret = sh.getAllDataAsArray(name, shCond.datRowS, shCond.datColS, shCond.datColE, dat, row, col)
    
    If ret = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, col)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clSheet_initSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "$verify"
    '=======================
    
    bRet = sh.existSheet(name)
    
    If bRet Then
        sh.initSheet (name)
        Debug.Print "result ::: initSheet done-->" & name & " |" & Now
    Else
        Debug.Print "result ::: err-->" & name & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_newSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim newName As String
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    '=======================
    
    'get new sheet name
    newName = sh.newSheet(name)
    
    Debug.Print "result ::: sheet name is-->" & newName & " |" & Now
End Sub

'==================================================
Sub verify_clSheet_existSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample5"
    '=======================
    
    'check existance of the sheet
    bRet = sh.existSheet(name)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & name & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & name & " |" & Now
    End If
    
End Sub
