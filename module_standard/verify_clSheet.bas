Attribute VB_Name = "verify_clSheet"
Option Explicit
Option Base 1

'==================================================
Sub verify_clSheet_convAllCellsOnSheetToValues()
    '事前準備：$verifyシートを作って、セルに適当な値を設定。
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet As Boolean
    Dim dummyArr(10, 10) As Variant
    Dim i As Integer, j As Integer
    For i = 1 To 10 Step 1
        For j = 1 To 10 Step 1
            dummyArr(i, j) = "=" & i & "+" & j
        Next j
    Next i
    'initialize the sheet to verification
    bRet = sh.initSheet(ThisWorkbook, "$verify")
    'plot all data on the $verify sheet
    With ThisWorkbook.sheets("$verify")
        .Select
        .Range(.Cells(1, 1), .Cells(UBound(dummyArr, 1), UBound(dummyArr, 2))) = dummyArr
    End With
    
    '=======================
    bRet = sh.convAllCellsOnSheetToValues(ThisWorkbook, "$verify")
    '=======================
    Set sh = Nothing
    
    If bRet = True Then
        Debug.Print "result ::: done " & " |" & Now
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clSheet_getDataAsArray()
    '事前準備：$verify1シートを作って、セルに適当な値を設定。
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet As Boolean
    Dim dummyArr(10, 10) As Variant
    Dim i As Integer, j As Integer
    For i = 1 To 10 Step 1
        For j = 1 To 10 Step 1
           dummyArr(i, j) = "dat_" & i & "_" & j
        Next j
    Next i
    'initialize the sheet to verification
    bRet = sh.initSheet(ThisWorkbook, "$verify1")
    'plot all data on the $verify sheet
    With ThisWorkbook.sheets("$verify1")
        .Select
        .Range(.Cells(1, 1), .Cells(UBound(dummyArr, 1), UBound(dummyArr, 2))) = dummyArr
    End With
    
    '=======================
    Dim dat As Variant
    Dim r As Long, c As Long
    'get data in the sheet
    bRet = sh.getDataAsArray(ThisWorkbook, "$verify1", 1, 5, 1, 7, dat, r, c)
   '=======================
     
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(ThisWorkbook, "$verify2")
        'plot all data on the $verify sheet
        With ThisWorkbook.sheets("$verify2")
            .Select
            .Range(.Cells(1, 1), .Cells(UBound(dat, 1), UBound(dat, 2))) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
    
    Set sh = Nothing
End Sub

'==================================================
Sub verify_clSheet_initSheet()
    '=======================
    Dim name As String
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet As Boolean
    name = "SampleSheetForTest"
    bRet = sh.initSheet(ThisWorkbook, name)
    Set sh = Nothing
    '=======================
    
    If bRet Then
        Debug.Print "result ::: initSheet done-->" & name & " |" & Now
    Else
        Debug.Print "result ::: err-->" & name & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_newSheet()
    '事前準備：'ダミーファイルを作成。
    Dim dummySheets As Variant
    Dim dummyWb As Workbook
    Dim bRet As Boolean
    dummySheets = Array()
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets, dummyWb)
    
    '=======================
    Dim name As String, name1 As String, name2 As String, name3 As String
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet1 As Boolean, bRet2 As Boolean, bRet3 As Boolean
    name = "SampleSheetForTest"
    bRet1 = sh.newSheet(dummyWb, name, name1)
    bRet2 = sh.newSheet(dummyWb, name, name2)
    bRet3 = sh.newSheet(dummyWb, name, name3)
    Set sh = Nothing
    '=======================
    
    If bRet1 And bRet2 And bRet3 Then
        Debug.Print "result ::: newSheet done-->" & CStr(name1) & " and " & CStr(name2) & " and " & CStr(name3) & " |" & Now
    Else
        Debug.Print "result ::: err-->" & name & " |" & Now
    End If
    
End Sub


'==================================================
Sub verify_clSheet_copySheet()
    '事前準備：'コピー元のシートを持つ、ダミーファイルを作成。
    Dim dummySheets As Variant
    Dim dummyWb As Workbook
    Dim bRet As Boolean
    Dim dummyArr(10, 10) As Variant
    Dim i As Integer, j As Integer
    dummySheets = Array("ToBeCopied")
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets, dummyWb)
    For i = 1 To 10 Step 1
        For j = 1 To 10 Step 1
           dummyArr(i, j) = "dat_" & i & "_" & j
        Next j
    Next i
    With dummyWb.sheets("ToBeCopied")
        .Select
        .Range(.Cells(1, 1), .Cells(UBound(dummyArr, 1), UBound(dummyArr, 2))) = dummyArr
    End With
    
    '=======================
    Dim actualName1 As String, actualName2 As String
    Dim sh As clSheet
    Set sh = New clSheet
    bRet = sh.copySheet(dummyWb, "ToBeCopied", "ToBeCopied", actualName1)
    bRet = sh.copySheet(dummyWb, "ToBeCopied", "ToBeCopied", actualName2)
    Set sh = Nothing
    '=======================
    
    If bRet Then
        Debug.Print "result ::: copySheet done-->ToBeCopied to " & CStr(actualName1) & " and " & CStr(actualName2) & Now
    Else
        Debug.Print "result ::: err |" & Now
    End If
End Sub


'==================================================
Sub verify_clSheet_existModule()
    Dim moName As String
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    moName = "clFiles"
    '=======================
    
    'check existance of the module.
    bRet = sh.existModule(wb, moName)
    Set sh = Nothing
    
    If bRet Then
        Debug.Print "result ::: exist-->" & moName & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & moName & " |" & Now
    End If
    
End Sub


'==================================================
Sub verify_clSheet_existSheet()
    Dim shName As String
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    shName = "Sheet1"
    '=======================
    
    'check existance of the sheet
    bRet = sh.existSheet(wb, shName)
    Set sh = Nothing
    
    If bRet Then
        Debug.Print "result ::: exist-->" & shName & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & shName & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_existSheetWithWildCardCharacter()
    Dim shName As String
    Dim sh As clSheet
    Set sh = New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    shName = "Sheet*"
    '=======================
    
    'check existance of the sheet
    Dim shNames As New Collection
    bRet = sh.existSheetWithWildCardCharacter(wb, shName, shNames)
    Set sh = Nothing
    
    If bRet Then
        Debug.Print "result ::: exist-->" & shNames.count & " sheets as " & shName & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & shName & " |" & Now
    End If
    
End Sub
