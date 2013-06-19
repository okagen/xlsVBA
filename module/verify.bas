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


'*** clAxCtrl内メソッド ***
'==================================================
Sub verify_clAxCtrl_putChkBoxesV()

    Dim name As String
    Dim rowS As Long
    Dim colVal As Long
    Dim colCtrl As Long
    Dim count As Long
    Dim sh As New clSheet
    Dim ax As New clAxCtrl
    Dim bRet As Boolean
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "$verify"
    rowS = 3
    colVal = 4
    colCtrl = 3
    count = 20
    
    Set wb = ThisWorkbook
    '=======================
    
    bRet = sh.existSheet(wb, name)
    
    If bRet Then
        bRet = sh.initSheet(wb, name)
    Else
        bRet = sh.newSheet(wb, name)
    End If
    
    'put check boxes on the seet
    bRet = ax.putChkBoxesV(wb, name, rowS, colVal, colCtrl, count)
    Debug.Print "result ::: putChkBoxesV done-->" & name & " |" & Now

End Sub

'*** clDatArr内メソッド ***
'==================================================
Sub verify_clDatArr_insertColIntoArray()
    Dim datBucket(1 To 10, 1 To 10)
    Dim row As Long
    Dim col As Long
    Dim datArr As New clDatArr
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim colIndex As Long
    Dim val As String
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook

    '=======================
    'create Array for test
    For row = 1 To 10 Step 1
        For col = 1 To 10 Step 1
            datBucket(row, col) = "org(" & row & "," & col & ")"
        Next col
    Next row
    
    colIndex = 3
    val = "Value col=" & colIndex
    '=======================
    
    bRet = datArr.insertColIntoArray(datBucket, colIndex, val, newDat)
    If Not bRet Then
        Exit Sub
    End If
    
    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clDatArr_formatArray()
    Dim datBucket(1 To 100, 1 To 100)
    Dim datArr As New clDatArr
    Dim sh As New clSheet
    Dim row As Long
    Dim col As Long
    Dim newRow As Long
    Dim newCol As Long
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook
    Dim bRet As Boolean

    '=======================
    'create Array for test
    For row = 1 To 100 Step 1
        For col = 1 To 100 Step 1
            datBucket(row, col) = "data in Bucket(" & row & "," & col & ")"
        Next col
    Next row
    
    newRow = 10
    newCol = 12
    '=======================
    
    bRet = datArr.formatArray(datBucket, newRow, newCol, newDat)

    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clDatArr_addArray()
    Dim datA As Variant
    Dim datB As Variant
    Dim datBucket As Variant
    Dim row As Long
    Dim col As Long
    Dim datArr As New clDatArr
    Dim lastIndex As Long
    Dim wb As Workbook
    Dim retRow As Long
    Dim retCol As Long
    Dim bRet As Boolean
    Dim sh As New clSheet
    
    ReDim datA(1 To 5, 1 To 10)
    ReDim datB(1 To 5, 1 To 10)
    ReDim datBucket(1 To 100, 1 To 10)
    
    '=======================
    'create Array for test
    For row = 1 To 5 Step 1
        For col = 1 To 10 Step 1
            datA(row, col) = "data_A(" & row & "," & col & ")"
            datB(row, col) = "data_B(" & row & "," & col & ")"
        Next col
    Next row
    '=======================
    
    bRet = datArr.addArray(datA, 1, datBucket, lastIndex)
    bRet = datArr.addArray(datB, lastIndex + 1, datBucket, lastIndex)
    
    retRow = UBound(datBucket, 1)
    retCol = UBound(datBucket, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = datBucket
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'*** clFiles内メソッド ***
'==================================================
Sub verify_clFiles_getWorkbookObj()
    Dim file As String
    Dim wb As Workbook
    Dim fls As New clFiles
    Dim bRet As Boolean
    
    '=======================
    'The file name for test
    file = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master\インターテクノ\BC-10 ver1.00.xls"
    '=======================
    
    Application.ScreenUpdating = False
    
    'get Workbook object
    bRet = fls.getWorkbookObj(file, wb)
    
    If bRet Then
        wb.Close savechanges:=False
    End If

    Application.ScreenUpdating = True

End Sub

'==================================================
Sub verify_clFiles_getFolderAndFileNameArr()
    Dim path As String
    Dim col As New Collection
    Dim dat As Variant
    Dim sh As New clSheet
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim row As Long
    Dim i As Long
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    path = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master"
    '=======================
    
    'set filter on the sheet
    bRet = fls.getAllXlsFilePathCol(path, col)
    
    'set filter on the sheet
    bRet = fls.getFolderAndFileNameArr(col, dat)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(col.count, 2)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clFiles_getAllXlsFilePathCol()
    Dim path As String
    Dim dat As New Collection
    Dim sh As New clSheet
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim i As Long
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    path = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master"
    '=======================
    
    'set filter on the sheet
    bRet = fls.getAllXlsFilePathCol(path, dat)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            For i = 1 To dat.count
                .Range(Cells(i, 1), Cells(i, 1)).Value = dat(i)
            Next
            
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
    Dim wb As Workbook
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    tgtFields.Add 1, "F45N"
    tgtFields.Add 2, "*Table*"
    tgtFields.Add 3, "13"
    
    Set wb = ThisWorkbook
    '=======================
    
    'set filter on the sheet
    bRet = sh.setFilter(wb, name, shCond.datRowS, shCond.datColS, shCond.datColE, tgtFields)
    
End Sub


'==================================================
Sub verify_clSheet_getRowDataVLookUp()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim bRet As Boolean
    Dim str As String
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 1
    str = "45"
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.getRowDataVLookUp(wb, name, shCond.datRowS, shCond.datColS, shCond.datColE, col, str, dat, row)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
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
    Dim allowDup As Boolean
    Dim wb As Workbook
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 1
    allowDup = False
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.getColDataAsArray(wb, name, shCond.datRowS, col, allowDup, dat, row)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
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
    Dim wb As Workbook
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.getAllDataAsArray(wb, name, shCond.datRowS, shCond.datColS, shCond.datColE, dat, row, col)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
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
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "sample5"
    
    Set wb = ThisWorkbook
    '=======================
    
    bRet = sh.existSheet(wb, name)
    
    If bRet Then
        bRet = sh.initSheet(wb, name)
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
    Dim wb As Workbook
     
    '=======================
    'The Sheet name for test
    name = "sample5"
    
    Set wb = ThisWorkbook
    '=======================
    
    'get new sheet name
    newName = sh.newSheet(wb, name)
    
    Debug.Print "result ::: sheet name is-->" & newName & " |" & Now
End Sub

'==================================================
Sub verify_clSheet_existSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    
    Set wb = ThisWorkbook
    '=======================
    
    'check existance of the sheet
    bRet = sh.existSheet(wb, name)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & name & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & name & " |" & Now
    End If
    
End Sub



'*** clSheets内メソッド ***
'==================================================
Sub verify_clSheets_getTargetSheets()
    Dim shs As New clSheets
    Dim ignoreNames As New Collection
    Dim targetNames As New Collection
    Dim bRet As Boolean
    Dim name As Variant
    Dim wb As Workbook
    
    '=======================
    'Sheet names to ignore
    ignoreNames.Add ("tool")
    ignoreNames.Add ("$")
    ignoreNames.Add ("ugl-")
    
    Set wb = ThisWorkbook
    '=======================
    
    'get target sheet names
    bRet = shs.getTargetSheets(wb, ignoreNames, targetNames)
    
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
    Dim wb As Workbook
    
    '=======================
    'The Sheet names to test
    names.Add ("sample1")
    names.Add ("sample2")
    names.Add ("sample3")
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in sheets
    bRet = shs.combineSheets(wb, names, shCond.datRowS, shCond.datColS, shCond.datColE, dat, row)
    
    'get number of row and column from array object.
    arrRow = UBound(dat, 1)
    arrCol = UBound(dat, 2)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(arrRow, arrCol)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub
