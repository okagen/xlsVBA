Attribute VB_Name = "verify_clSheet"
Option Explicit
Option Base 1

'==================================================
Sub verify_clSheet_getDataAsArray()
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim dummyArr(10, 10) As Variant
    Dim i As Integer, j As Integer
    '=======================
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
End Sub

'==================================================
Sub verify_clSheet_initSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    
    '=======================
    name = "SampleSheetForTest"
    Set wb = ThisWorkbook
    '=======================
    
    bRet = sh.initSheet(wb, name)
    
    If bRet Then
        Debug.Print "result ::: initSheet done-->" & name & " |" & Now
    Else
        Debug.Print "result ::: err-->" & name & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_newSheet()
    Dim name As String, name1 As String, name2 As String, name3 As String
    Dim sh As New clSheet
    Dim bRet1 As Boolean, bRet2 As Boolean, bRet3 As Boolean
    Dim wb As Workbook
    
    '=======================
    name = "SampleSheetForTest"
    Set wb = ThisWorkbook
    '=======================
    
    bRet1 = sh.newSheet(wb, name, name1)
    bRet2 = sh.newSheet(wb, name, name2)
    bRet3 = sh.newSheet(wb, name, name3)
    
    If bRet1 And bRet2 And bRet3 Then
        Debug.Print "result ::: newSheet done-->" & CStr(name1) & " and " & CStr(name2) & " and " & CStr(name3) & " |" & Now
    Else
        Debug.Print "result ::: err-->" & name & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_existModule()
    Dim moName As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    moName = "clFiles"
    '=======================
    
    'check existance of the module.
    bRet = sh.existModule(wb, moName)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & moName & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & moName & " |" & Now
    End If
    
End Sub


'==================================================
Sub verify_clSheet_existSheet()
    Dim shName As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    shName = "Sheet1"
    '=======================
    
    'check existance of the sheet
    bRet = sh.existSheet(wb, shName)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & shName & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & shName & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_existSheetWithWildCardCharacter()
    Dim shName As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    shName = "Sheet*"
    '=======================
    
    'check existance of the sheet
    Dim shNames As New Collection
    bRet = sh.existSheetWithWildCardCharacter(wb, shName, shNames)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & shNames.count & " sheets as " & shName & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & shName & " |" & Now
    End If
    
End Sub
