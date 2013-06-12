Attribute VB_Name = "verify"
Option Explicit

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
    str = "F45N"
    '=======================
    
    'get data in the sheet
    ret = sh.getRowDataVLookUp(name, col, str, dat, row)
    
    If ret = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, 7)) = dat
            Debug.Print "result ::: done " & Now
        End With
    Else
        Debug.Print "result ::: no data" & Now
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
    ret = sh.getColDataAsArray(name, col, allowDup, dat, row)
    
    If ret = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, 1)) = dat
            Debug.Print "result ::: done " & Now
        End With
    Else
        Debug.Print "result ::: no data" & Now
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
    ret = sh.getAllDataAsArray(name, dat, row, col)
    
    If ret = True Then
        'initialize the sheet to verification
        sh.initSheet ("$verify")
        'plot all data on the $verify sheet
        With Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, col)) = dat
            Debug.Print "result ::: done " & Now
        End With
    Else
        Debug.Print "result ::: no data" & Now
    End If
End Sub


'==================================================
Sub verify_clSheet_newSheet()
    Dim name As String
    Dim sh As New clSheet
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    '=======================
    
    'get data in the sheet
    sh.newSheet (name)
    
    Debug.Print "result ::: done " & Now
End Sub


