Attribute VB_Name = "verify_clSheet"
Option Explicit
Option Base 1


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
    
    bRet = sh.initSheet(wb, name)
    
    If bRet Then
        Debug.Print "result ::: initSheet done-->" & name & " |" & Now
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
