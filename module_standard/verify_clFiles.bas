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
    tgtModules = Array("verify_clFiles", "verify_clSheet", "verify_clSheets")
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
    
    'ダミーのシートを持つ、ダミー①のファイルを作成。
    Dim dummySheets1 As Variant
    Dim dummyWb1 As Workbook
    dummySheets1 = Array("dummy1", "dummy2", "dummy3", "dummy4", "dummy5")
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets1, dummyWb1)
    
    'ダミーのシートを持つ、ダミー②のファイルを作成。
    Dim dummySheets2 As Variant
    Dim dummyWb2 As Workbook
    dummySheets2 = Array("test1", "test2", "test3")
    bRet = verify_clFiles_makeDummyExcelFileWithDummySheets(dummySheets2, dummyWb2)
    
    'ダミー①から、ダミー②へ、シートをいくつかコピー
    '=======================
    tgtSheets = Array("dummy1", "dummy3", "dummy5")
    tgtStdModules = Array()
    tgtClsModules = Array()
    bRet = fl.copySheetsAndModulesIntoNewFileThenSave(dummyWb1, dummyWb2, tgtSheets, tgtStdModules, tgtClsModules)
    Debug.Print "result ::: verify_clFiles_copySheetsAndModulesIntoNewFileThenSave done -->" & CStr(bRet) & " |" & Now
    '=======================
    
    'Thisworkbookから、ダミー②へ、モジュールをいくつかコピー
    '=======================
    tgtSheets = Array()
    tgtStdModules = Array("verify", "verify_clFiles")
    tgtClsModules = Array("clFiles", "clSheet")
    bRet = fl.copySheetsAndModulesIntoNewFileThenSave(ThisWorkbook, dummyWb2, tgtSheets, tgtStdModules, tgtClsModules)
    Debug.Print "result ::: verify_clFiles_copySheetsAndModulesIntoNewFileThenSave done -->" & CStr(bRet) & " |" & Now
    '=======================
    
    
'    tgtStdModules = Array("verify", "verify_clFiles")
'    tgtClsModules = Array("clFiles", "clSheet")
'    toPath = wb.Path
'    fileName = "verify_clFiles_copySheetsIntoNewFile1"
'
    
'    'Save the new book as ...
'    Dim objFso As Object
'    Dim newWbPath As String
'    Set objFso = CreateObject("Scripting.FileSystemObject")
'         newWbPath = objFso.BuildPath(toPath, fileName)
'    Set objFso = Nothing
'
'    On Error Resume Next
'    If bRet2 = False And bRet3 = False Then
'        Call wbNew.SaveAs(fileName:=newWbPath, FileFormat:=xlOpenXMLWorkbook)
'    Else
'        Call wbNew.SaveAs(fileName:=newWbPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled)
'    End If
'    On Error GoTo 0
End Sub


'==================================================
Private Function verify_clFiles_makeDummyExcelFileWithDummySheets(ByVal sheets As Variant, _
                                                                                        ByRef dummyWb As Workbook) As Boolean
    Workbooks.Add
    Set dummyWb = Application.ActiveWorkbook
    
    Dim sh As New clSheet
    Dim i As Long
    For i = 1 To UBound(sheets) Step 1
        bRet = sh.initSheet(dummyWb, sheets(i))
    Next i
    
    verify_clFiles_makeDummyExcelFileWithDummySheets = True
End Function
