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

'==================================================
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

'==================================================
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
        bRet = getFolderAndFileName(fullPath, folder, file)
        If bRet Then
            retTmp(i, 1) = folder
            retTmp(i, 2) = file
        End If
    Next i
    
    dat = retTmp
    getFolderAndFileNameArr = True
End Function

'==================================================
'ファイルのフルパスを受け、ファイル名と保存されているフォルダ名を返す
'  [i]fullPaths     Excelのフルパスを蓄積したCollection
'  [o]folder        フォルダ名
'  [o]file             ファイル名
'--------------------------------------------------
Public Function getFolderAndFileName(ByVal fullPath As String, _
                                                                ByRef folder As String, _
                                                                ByRef file As String _
                                                                ) As Boolean
    Dim FSO As Object
    Dim bRet As Boolean
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'ファイルがない場合エラー
    bRet = FSO.FileExists(fullPath)
    If Not bRet Then
        Debug.Print "err ::: file does not exist.(" & fullPath & ")  [func:insertColIntoArray()]" & Now
        getFolderAndFileName = False
        Exit Function
    End If
    
    file = FSO.GetFile(fullPath).name
    folder = FSO.getFolder(FSO.GetFile(fullPath).ParentFolder).name
    
    Set FSO = Nothing
    getFolderAndFileName = True
End Function

'==================================================
'指定したディレクトリ配下にある全Excelファイルのフルパス取得
'  [i]path      検索するディレクトリのパス
'  [o]ret       Excelのフルパスを蓄積したCollection
'--------------------------------------------------
Public Function getAllXlsFilePathCol(ByVal path As String, _
                                    ByRef ret As Collection) As Boolean
    Dim FSO As Object
    Dim folder As Variant
    Dim file As Variant
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each folder In FSO.getFolder(path).SubFolders
        Call getAllXlsFilePathCol(folder.path, ret)
    Next folder
    For Each file In FSO.getFolder(path).files
        If InStr(file.Type, "Excel") > 0 Then
            ret.Add (file.path)
        End If
    Next file
    Set FSO = Nothing
    
    If ret.count < 1 Then
        getAllXlsFilePathCol = False
    Else
        getAllXlsFilePathCol = True
    End If
End Function


                            
