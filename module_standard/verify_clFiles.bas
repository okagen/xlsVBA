Attribute VB_Name = "verify_clFiles"
Option Explicit
Option Base 1


'==================================================
Sub verify_clFiles_exportModules()
    Dim bRet As Boolean
    Dim tgtSheets As Variant
    Dim tgtModules As Variant
    Dim toPath As String
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim fl As New clFiles
    
    'export class modules.
    '=======================
    tgtModules = Array("clArr", "clFiles", "clSheet", "clSheets")
    toPath = "C:\my_work\GitHub\excel_vba\module_class"
    '=======================
    bRet = fl.exportModules(wb, tgtModules, toPath)
    Debug.Print "result ::: verify_clFiles_exportModules done -->" & CStr(bRet) & " |" & Now

    'export standard modules.
    '=======================
    tgtModules = Array("verify_clFiles", "verify_clFiles", "verify_clSheet", "verify_clSheets")
    toPath = "C:\my_work\GitHub\excel_vba\module_standard"
    '=======================
    bRet = fl.exportModules(wb, tgtModules, toPath)
    Debug.Print "result ::: verify_clFiles_exportModules done -->" & CStr(bRet) & " |" & Now


End Sub



'==================================================
Sub verify_clFiles_copySheetsAndModulesIntoNewFileThenSave()
    Dim bRet As Boolean
    Dim tgtSheets As Variant
    Dim tgtStdModules As Variant
    Dim tgtClsModules As Variant
    Dim cmpKind As Integer
    Dim toPath As String, fileName As String
    Dim wb As Workbook, wbNew As Workbook
    Set wb = ThisWorkbook
    Dim fl As New clFiles
    
    '=======================
    tgtSheets = Array("R02’†Œ‹‰Ê_‘Œê", "H29¬Œ‹‰Ê_‘ŒêA", "$—ÌˆæŠÏ“__R02’†_‘Œê", "$—ÌˆæŠÏ“__H29¬_‘ŒêA")
    tgtStdModules = Array("verify", "verify_clFiles")
    tgtClsModules = Array("clFiles", "clSheet")
    toPath = wb.Path
    fileName = "verify_clFiles_copySheetsIntoNewFile"
    '=======================
    
    'put check boxes on the seet
    bRet = fl.copySheetsAndModulesIntoNewFileThenSave(wb, tgtSheets, tgtStdModules, tgtClsModules, toPath, fileName, wbNew)
    Debug.Print "result ::: verify_clFiles_copySheetsAndModulesIntoNewFileThenSave done -->" & CStr(bRet) & " |" & Now

End Sub

