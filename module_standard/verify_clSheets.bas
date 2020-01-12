Attribute VB_Name = "verify_clSheets"
Option Explicit
Option Base 1

'==================================================
Sub verify_clSheets_deleteUnSpecifiedSheets()

    'ダミーのシートを持つ、ダミー①のファイルを作成。
    Dim dummySheets As Variant
    Dim dummyWb As Workbook
    Dim bRet As Boolean
    dummySheets = Array("dummy1", "dummy2", "dummy3", "dummy4", "dummy5")
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets, dummyWb)
    
    Dim shs As New clSheets
    Dim remainSheets As Variant
    '=======================
    remainSheets = Array("dummy2", "dummy3", "dummy4")
    '=======================
    
    'check existance of the module.
    bRet = shs.deleteUnSpecifiedSheets(dummyWb, remainSheets)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & CStr(bRet) & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & CStr(bRet) & " |" & Now
    End If
    
End Sub
