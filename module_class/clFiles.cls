VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'=============================
'指定されたモジュールを指定したディレクトリへエクスポートする。
'  [i]wb                   エクスポート元のExcelブックオブジェクト
'  [i]tgtModules       エクスポート対象のモジュール名の配列
'  [i]toPath             エクスポート先のExcelブックオブジェクト
'  [o]copiedModules      エクスポートが成功したモジュール名前群
'  関数の返り値：一つ以上エクスポートに成功したらTrue
'--------------------------------------------------
Public Function exportModules( _
                                        ByVal wb As Workbook, _
                                        ByVal tgtModules As Variant, _
                                        ByVal toPath As String _
                                        ) As Boolean
    Dim tgtMo As String, prg As String
    Dim i As Long
    Dim sh As New clSheet
    Dim ext As String
    Dim bRet As Boolean
    
    bRet = False
    For i = 1 To UBound(tgtModules) Step 1
        tgtMo = tgtModules(i)
        If sh.existModule(wb, tgtMo) Then
        
            With wb.VBProject.VBComponents(tgtMo)
                    '指定されたモジュールの拡張子を取得
                    Select Case .Type
                    Case vbext_ct_StdModule
                        ext = "bas"
                    Case vbext_ct_ClassModule, vbext_ct_Document
                        ext = "cls"
                    Case vbext_ct_MSForm
                        ext = "frm"
                    End Select
                    
                    'Export the module.
                    Dim objFso As Object
                    Dim exportFilePath As String
                    Set objFso = CreateObject("Scripting.FileSystemObject")
                         exportFilePath = objFso.BuildPath(toPath, tgtMo & "." & ext)
                    Set objFso = Nothing
                    .Export exportFilePath
                    
                    bRet = True
            End With
        End If
    Next i
    exportModules = bRet
End Function


'=============================
'指定されたシート、モジュールを、新しいブックの中にコピーし保存する。
'マクロが含まれる場合xlsmで保存する。
'  [i]wb                   コピー元のExcelブックオブジェクト
'  [i]tgtSheets           コピー対象のシート名の配列
'  [i]tgtStdModules           コピー対象の標準モジュール名の配列
'  [i]tgtClsModules           コピー対象のクラスモジュール名の配列
'  [i]toPath                   コピー先のExcelブックオブジェクト
'  [i]fileName                ファイル名
'  [o]wbNew      保存したExcelファイルのオブジェクト
'  関数の返り値：一つ以上コピーに成功したらTrue
'--------------------------------------------------
Public Function copySheetsAndModulesIntoNewFileThenSave( _
                                        ByVal wb As Workbook, _
                                        ByVal tgtSheets As Variant, _
                                        ByVal tgtStdModules As Variant, _
                                        ByVal tgtClsModules As Variant, _
                                        ByVal toPath As String, _
                                        ByVal fileName As String, _
                                        ByRef wbNew As Workbook _
                                        ) As Boolean
    Workbooks.Add
    Set wbNew = Application.ActiveWorkbook
    
    Dim bRet1 As Boolean, bRet2 As Boolean, bRet3 As Boolean
    Dim copiedSheets As New Collection, copiedModules As New Collection
    bRet1 = copySheets(wb, tgtSheets, wbNew, copiedSheets)
    bRet2 = copyModules(wb, tgtStdModeles, vbext_ct_StdModule, wbNew, copiedModules)
    bRet3 = copyModules(wb, tgtClsModeles, vbext_ct_ClassModule, wbNew, copiedModules)

    'Save the new book as ...
    Dim objFso As Object
    Dim newWbPath As String
    Set objFso = CreateObject("Scripting.FileSystemObject")
         newWbPath = objFso.BuildPath(toPath, fileName)
    Set objFso = Nothing
    
    On Error Resume Next
    If bRet2 = False And bRet3 = False Then
        Call wbNew.SaveAs(fileName:=newWbPath, FileFormat:=xlOpenXMLWorkbook)
    Else
        Call wbNew.SaveAs(fileName:=newWbPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    End If
    On Error GoTo 0
    
    copySheetsAndModulesIntoNewFileThenSave = bRet1 Or bRet2 Or bRet3
End Function


'=============================
'指定されたモジュールを、別のブックの中にコピーする。
'  [i]wbFrom                コピー元のExcelブックオブジェクト
'  [i]tgtModules           コピー対象のモジュール名の配列
'  [i]wbTo                   コピー先のExcelブックオブジェクト
'  [o]copiedModules      コピーが成功したモジュール名前群
'  関数の返り値：一つ以上コピーに成功したらTrue
'--------------------------------------------------
Public Function copyModules(ByVal wbFrom As Workbook, _
                                            ByVal tgtModules As Variant, _
                                            ByVal cmpKind As Integer, _
                                            ByRef wbTo As Workbook, _
                                            ByRef copiedModules As Collection _
                                           ) As Boolean
    Dim tgtMo As String, prg As String
    Dim i As Long
    Dim sh As New clSheet
    Dim cntLines As Long
    
    For i = 1 To UBound(tgtModules) Step 1
        tgtMo = tgtModules(i)
        If sh.existModule(wbFrom, tgtMo) Then
            With wbFrom.VBProject.VBComponents(tgtMo).CodeModule
                prg = .Lines(1, .CountOfLines)
            End With
            
            With wbTo.VBProject.VBComponents.Add(cmpKind)
                '「変数の宣言を強制する」オプションが有効の場合、Option Explicitの行が自動的にできるため、あったら削除。
                cntLines = .CodeModule.CountOfLines
                If cntLines > 0 Then
                    .CodeModule.DeleteLines 1, cntLines
                End If
                .CodeModule.AddFromString prg
                .name = tgtMo
            End With
            copiedModules.Add (tgtMo)
        End If
    Next i

    If copiedModules.count > 0 Then
        copyModules = True
    Else
        copyModules = False
    End If

End Function

'=============================
' ブック内の指定されたシートを、別のブックの中にコピーする。
'  [i]wbFrom                コピー元のExcelブックオブジェクト
'  [i]tgtSheets             コピー対象のシート名の配列
'  [i]wbTo                   コピー先のExcelブックオブジェクト
'  [o]copiedSheets      コピーが成功したシートの名前群
'  関数の返り値：一つ以上コピーに成功したらTrue
'--------------------------------------------------
Public Function copySheets(ByVal wbFrom As Workbook, _
                                                ByVal tgtSheets As Variant, _
                                                ByVal wbTo As Workbook, _
                                                ByRef copiedSheets As Collection _
                                                ) As Boolean
    Dim i As Long
    Dim tgtSh As String
    Dim sh As clSheet
    Set sh = New clSheet
    
    For i = 1 To UBound(tgtSheets) Step 1
        tgtSh = tgtSheets(i)
        If sh.existSheet(wbFrom, tgtSh) Then
            wbFrom.Sheets(tgtSh).Copy After:=wbTo.Sheets(wbTo.Sheets.count)
            copiedSheets.Add (tgtSh)
        End If
    Next i
    
    If copiedSheets.count > 0 Then
        copySheets = True
    Else
        copySheets = False
    End If
End Function


'----------------------------------------------------------------------------------
'↓↓↓↓ 見直し中

'=============================
'指定フォルダ内のファイルを、別のフォルダに新しいファイル名でコピーする。
'  [i]fromPath            コピー元のファイルがあるフォルダのパス
'  [i]toPath                コピー先のフォルダのパス
'  [i]dicFileNames      コピーの前後のファイル名が設定されたハッシュテーブルCreateObject("Scripting.Dictionary")で生成
'--------------------------------------------------
Public Function copyFiles(ByVal fromPath As String, _
                                        ByVal toPath As String, _
                                        ByVal dicFileNames As Variant _
                                        ) As Boolean
    Dim sRet As String
    Dim i As Long
    Dim FSO As Object
    Dim fromFile As String
    Dim toFile As String
    Dim keys As Variant
    
    'ディレクトリが存在するかチェック
    sRet = Dir(fromPath, vbDirectory)
    If sRet = "" Then
        Debug.Print "err ::: directory does not exist.(" & fromPath & ")  [func:copyFiles()]" & Now
        copyFiles = False
        Exit Function
    End If
    sRet = Dir(toPath, vbDirectory)
    If sRet = "" Then
        Debug.Print "err ::: directory does not exist.(" & toPath & ")  [func:copyFiles()]" & Now
        copyFiles = False
        Exit Function
    End If
        
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'dicFileNamesのkey部分(コピー元のファイル名)を配列として取得
    keys = dicFileNames.keys()
    
    'keyの数分ループして新ファイル名で上書きコピーする
    For i = 0 To UBound(keys) Step 1
        fromFile = FSO.BuildPath(fromPath, keys(i))
        toFile = FSO.BuildPath(toPath, dicFileNames(keys(i)))
        FSO.copyFile fromFile, toFile, True
    Next i
    
    Set FSO = Nothing
    copyFiles = True
End Function

'=============================
'ファイル名を受けてworkbookオブジェクトを取得
'  [i]file  Excelのフルパス
'  [o]wb    workbookオブジェクト
'  *** 取得したwbオブジェクトは本関数の外でwb.closeする
'--------------------------------------------------
Public Function getWorkbookObj(ByVal file As String, _
                        ByRef wb As Workbook) As Boolean
    Dim ret As Workbook
    Dim wbTmp As Workbook
    Dim buf As String
    
    'ファイルの存在確認
    buf = Dir(file)
    If buf = "" Then
        getWorkbookObj = False
        Exit Function
    End If
    
    '既に開いていた場合
    For Each wbTmp In Workbooks
        If wbTmp.name = buf Then
            Set wb = wbTmp
            getWorkbookObj = True
            Exit Function
        End If
    Next wbTmp
    
    Set wb = GetObject(file)
    getWorkbookObj = True
    
End Function

'=============================
'フルパスのCollectionを受け、ファイル名と保存されているフォルダ名のCollectionを返す
'  [i]fullPaths Excelのフルパスを蓄積したCollection
'  [o]folders       ファイル名と保存されているフォルダ名の2次元Array
'  [o]filenames       ファイル名と保存されているフォルダ名の2次元Array
'--------------------------------------------------
Public Function getFolderAndFileNameColl(ByVal fullPaths As Collection, _
                                                                ByRef folders As Collection, _
                                                                ByRef filenames As Collection _
                                                                ) As Boolean
    If fullPaths.count < 1 Then
        getFolderAndFileNameColl = False
        Exit Function
    End If
    
    Dim retFolders As New Collection
    Dim retFiles As New Collection
    Dim item As Variant
    Dim file As String
    Dim folder As String
    Dim bRet As Boolean
  
    For Each item In fullPaths
        bRet = getFolderAndFileName(item, True, folder, file)
        If bRet Then
            retFolders.Add (folder)
            retFiles.Add (file)
        End If
    Next item
    
    Set folders = retFolders
    Set filenames = retFiles
    getFolderAndFileNameColl = True
End Function

'=============================
'フルパスのCollectionを受け、ファイル名と保存されているフォルダ名の2次元Arrayを返す
'  [i]fullPaths Excelのフルパスを蓄積したCollection
'  [o]dat       ファイル名と保存されているフォルダ名の2次元Array
'--------------------------------------------------
Public Function getFolderAndFileNameArr(ByVal fullPaths As Collection, _
                                        ByRef dat As Variant) As Boolean
    Dim retTmp As Variant
    Dim fullPath As String
    Dim file As String
    Dim folder As String
    Dim i As Long
    Dim bRet As Boolean
    
    If fullPaths.count < 1 Then
        getFolderAndFileNameArr = False
        Exit Function
    End If
    
    ReDim retTmp(1 To fullPaths.count, 1 To 2)
    For i = 1 To fullPaths.count Step 1
        fullPath = fullPaths(i)
        bRet = getFolderAndFileName(fullPath, True, folder, file)
        If bRet Then
            retTmp(i, 1) = folder
            retTmp(i, 2) = file
        End If
    Next i
    
    dat = retTmp
    getFolderAndFileNameArr = True
End Function

'=============================
'ファイルのフルパスを受け、ファイル名と保存されているフォルダ名を返す
'  [i]fullPaths     Excelのフルパスを蓄積したCollection
'  [i]extflg        拡張子の要否 true=拡張子あり false=拡張子なし
'  [o]folder        フォルダ名
'  [o]file            ファイル名 拡張子あり
'--------------------------------------------------
Public Function getFolderAndFileName(ByVal fullPath As String, _
                                                            ByVal extFlg As Boolean, _
                                                                ByRef folder As String, _
                                                                ByRef file As String _
                                                                ) As Boolean
    Dim FSO As Object
    Dim bRet As Boolean
    Dim dot As Long
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'ファイルがない場合エラー
    bRet = FSO.FileExists(fullPath)
    If Not bRet Then
        Debug.Print "err ::: file does not exist.(" & fullPath & ")  [func:insertColIntoArray()]" & Now
        getFolderAndFileName = False
        Exit Function
    End If
    
    file = FSO.GetFile(fullPath).name
    '拡張子を除いたファイル名の取得
    If extFlg = False Then
        dot = InStrRev(file, ".")
        file = Left(file, dot - 1)
    End If
    
    folder = FSO.getFolder(FSO.GetFile(fullPath).ParentFolder).name
    
    Set FSO = Nothing
    getFolderAndFileName = True
End Function

'=============================
'指定したディレクトリ配下にある全Excelファイルのフルパス取得
'  [i]path      検索するディレクトリのパス
'  [o]ret       Excelのフルパスを蓄積したCollection
'--------------------------------------------------
Public Function getAllXlsFilePathCol(ByVal Path As String, _
                                    ByRef ret As Collection) As Boolean
    Dim FSO As Object
    Dim folder As Variant
    Dim file As Variant
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each folder In FSO.getFolder(Path).SubFolders
        Call getAllXlsFilePathCol(folder.Path, ret)
    Next folder
    For Each file In FSO.getFolder(Path).files
        If InStr(file.Type, "Excel") > 0 Then
            ret.Add (file.Path)
        End If
    Next file
    Set FSO = Nothing
    
    If ret.count < 1 Then
        getAllXlsFilePathCol = False
    Else
        getAllXlsFilePathCol = True
    End If
End Function

'=============================
'Workbookと同ディレクトリにある指定した拡張子のファイルのファイル名取得
'  [o]ret       ファイル名を蓄積したCollection
'--------------------------------------------------
Public Function getAllFileNameCol(ByVal ext As String, _
                                    ByRef ret As Collection) As Boolean
    Dim FSO As Object
    Dim folder As Variant
    Dim file As Variant
    
    Dim Path As String
    Path = ThisWorkbook.Path
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each folder In FSO.getFolder(Path).SubFolders
        Call getAllXlsFilePathCol(folder.Path, ret)
    Next folder
    For Each file In FSO.getFolder(Path).files
        If InStr(1, file.name, ext, vbTextCompare) > 0 Then
            ret.Add (file.name)
        End If
    Next file
    Set FSO = Nothing
    
    If ret.count < 1 Then
        getAllFileNameCol = False
    Else
        getAllFileNameCol = True
    End If
End Function
                            
