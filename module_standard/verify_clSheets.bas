Attribute VB_Name = "verify_clSheets"
Option Explicit
Option Base 1

'==================================================
Sub verify_clSheets_deleteUnSpecifiedSheets()

    Dim shs As New clSheets
    Dim remainSheets As Variant
    Dim bRet As Boolean
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    '=======================
    remainSheets = Array("R02’†Œ‹‰Ê_‘Œê", "H29¬Œ‹‰Ê_‘ŒêA", "$—ÌˆæŠÏ“__R02’†_‘Œê", "$—ÌˆæŠÏ“__H29¬_‘ŒêA")

    '=======================
    
    'check existance of the module.
    bRet = shs.deleteUnSpecifiedSheets(wb, remainSheets)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & CStr(bRet) & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & CStr(bRet) & " |" & Now
    End If
    
End Sub
