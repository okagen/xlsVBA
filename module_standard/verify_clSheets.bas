Attribute VB_Name = "verify_clSheets"
Option Explicit
Option Base 1


'==================================================
Sub verify_clSheet_hideAndShowSheetsWithPrefix()
    '事前準備：ダミーのシートを持つ、ダミーのファイルを作成。
    Dim dummySheets As Variant
    Dim dummyWb As Workbook
    Dim bRet As Boolean
    dummySheets = Array("$dummy1", "$dummy2", "$$dummy3", "$$dummy4", "$$dummy5")
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets, dummyWb)
    
    '=======================
    Dim shs As New clSheets
    bRet = shs.hideSheetsWithPrefix(dummyWb, "$", "$$")
    bRet = shs.hideSheetsWithPrefix(dummyWb, "$$", "")
    '=======================
    '=======================
    bRet = shs.showSheetsWithPrefix(dummyWb, "$", "$$")
    bRet = shs.showSheetsWithPrefix(dummyWb, "$")
    '=======================
    
    If bRet = True Then
        Debug.Print "result ::: done " & " |" & Now
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clSheet_convAllCellsOnSheetsToValues()
    '事前準備：ダミーのシートを持つ、ダミーのファイルを作成。
    Dim sh As New clSheet
    Dim shs As New clSheets
    Dim bRet As Boolean
    Dim dummyArr(10, 10) As Variant
    Dim i As Integer, j As Integer
    For i = 1 To 10 Step 1
        For j = 1 To 10 Step 1
            dummyArr(i, j) = "=" & i & "+" & j
        Next j
    Next i
    'initialize the sheet to verification
    bRet = sh.initSheet(ThisWorkbook, "$verify1")
    bRet = sh.initSheet(ThisWorkbook, "$verify2")
    'plot all data on the $verify sheet
    With ThisWorkbook.sheets("$verify1")
        .Select
        .Range(.Cells(1, 1), .Cells(UBound(dummyArr, 1), UBound(dummyArr, 2))) = dummyArr
    End With
    With ThisWorkbook.sheets("$verify2")
        .Select
        .Range(.Cells(1, 1), .Cells(UBound(dummyArr, 1), UBound(dummyArr, 2))) = dummyArr
    End With
    
    '=======================
    Dim sheets As Variant
    sheets = Array("$verify1")
    bRet = shs.convAllCellsOnSheetsToValues(ThisWorkbook, sheets)
    '=======================
    
    If bRet = True Then
        Debug.Print "result ::: done " & " |" & Now
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clSheets_deleteUnSpecifiedSheets()
    '事前準備：ダミーのシートを持つ、ダミーのファイルを作成。
    Dim dummySheets As Variant
    Dim dummyWb As Workbook
    Dim bRet As Boolean
    dummySheets = Array("dummy1", "dummy2", "dummy3", "dummy4", "dummy5")
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets, dummyWb)
    
    '=======================
    Dim shs As New clSheets
    Dim remainSheets As Variant
    remainSheets = Array("dummy2", "dummy3", "dummy4")
    bRet = shs.deleteUnSpecifiedSheets(dummyWb, remainSheets)
    '=======================
    
    If bRet Then
        Debug.Print "result ::: exist-->" & CStr(bRet) & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & CStr(bRet) & " |" & Now
    End If
    
End Sub
