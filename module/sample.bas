Attribute VB_Name = "sample"
Option Explicit

'====================
'インプットシート生成
'====================
Sub inputSh_main(ByVal rootPath As String)
    Dim bRet As Boolean
    Dim dealers As Variant
    Dim dealers2 As Variant
    Dim col As Long
    Dim row As Long
    Dim sh As New clSheet
    Dim da As New clDB
    Dim rootDir As String
    
    'ディーラ名を取得し、DBシートへ保存
    bRet = sh.getDataAsArray(ThisWorkbook, TOOL, 21, 21, 5, 11, dealers, row, col)
    
    '2次元配列の行列を入れ替え
    dealers2 = Application.WorksheetFunction.Transpose(dealers)
    bRet = da.setDataArr(dbnum.dealer_name, dealers2)
    
    '必要なディレクトリを作成
    bRet = inputSh_makeDir(rootDir)
    
    'パーツ構成マスターファイルを配布用に名前を変えてインプットシートフォルダへコピー
    bRet = copyConfmasterFile(rootDir)
    
    'ディーラー毎にインプットシートをコピー
    bRet = dealFilesToDealers(rootDir)
End Sub

'ディーラー毎にインプットシートをコピー
Private Function dealFilesToDealers(ByVal parentDir As String)
    
    'どのファイルをどのディーラに配布するかの表(配布マトリクス)の状態を取得
    Dim dealMatrix As Variant
    Dim row As Long
    Dim col As Long
    Dim sh As New clSheet
    Dim bRet As Boolean
    bRet = sh.getDataAsArray(ThisWorkbook, TOOL, 22, 26, 5, 11, dealMatrix, row, col)
    
    'DBシート内の配布用インプットシート名の一覧とディーラ名を取得
    Dim newNames As Collection
    Dim dealers As Collection
    Dim db As New clDB
    bRet = db.getDataColl(dbnum.confmaster_newName, newNames)
    bRet = db.getDataColl(dbnum.dealer_name, dealers)
    
    '配布用インプットシートの保存先のリストをDBシートに作成
    Dim bVal As Boolean
    Dim i As Long
    Dim j As Long
    Dim inpShPath As String
    Dim FSO As Object
    Dim fromFiles As New Collection
    Dim toFiles As New Collection
    Dim fromFolder As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    fromFolder = FSO.BuildPath(parentDir, inpSh)

    For i = 1 To UBound(dealMatrix, 1) Step 1
        For j = 1 To UBound(dealMatrix, 2) Step 1
            bVal = dealMatrix(i, j)
            If bVal Then
                inpShPath = FSO.BuildPath(parentDir, dealers(j))
                inpShPath = inpShPath & "\" & newNames(i)
                toFiles.Add (inpShPath)
                fromFiles.Add (fromFolder & "\" & newNames(i))
            End If
        Next j
    Next i
    
    bRet = db.setDataColl(dbnum.inpSh_fromPath, fromFiles)
    bRet = db.setDataColl(dbnum.inpSh_toPath, toFiles)

    For i = 1 To fromFiles.count Step 1
            FSO.copyFile fromFiles(i), toFiles(i), True
    Next i
    Set FSO = Nothing

End Function

'パーツ構成マスターファイルを配布用に名前を変えてインプットシートフォルダへコピー
Private Function copyConfmasterFile(ByVal parentDir As String) As Boolean
    Dim bRet As Boolean
    Dim db As New clDB
    Dim sh As New clSheet
    Dim confmasterColl As Collection
    Dim newNames As Variant
    Dim lastRow As Long
    Dim da As New clDatArr
    '$toolシート上の配布時ファイル名を取得
    bRet = sh.getColDataAsArray(ThisWorkbook, TOOL, _
                                                toolSh.rowUL, toolSh.colUL + 1, _
                                                True, newNames, lastRow)
    bRet = da.removeEmptyRecord(newNames, newNames)
    'DBシートへペースト
    bRet = db.setDataArr(dbnum.confmaster_newName, newNames)
    'オリジナルのパーツ構成マスターのリストを取得
    Dim orgFiles As New Collection
    Dim newFileNames As New Collection
    bRet = db.getDataColl(dbnum.confmaster_orgPath, orgFiles)
    bRet = db.getDataColl(dbnum.confmaster_newName, newFileNames)
    '全シートを入れておくためのフォルダを作成
    Dim newFolderPath As String
    Dim di As New clDir
    bRet = di.createFolder(parentDir, inpSh, newFolderPath)
    'コピー
    Dim item As Variant
    Dim FSO As Object
    Dim newPath As String
    Dim i As Long
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For i = 1 To orgFiles.count Step 1
            newPath = FSO.BuildPath(newFolderPath, newFileNames(i))
            FSO.copyFile orgFiles(i), newPath, True
    Next i
    Set FSO = Nothing
End Function

'デスクトップにインプットシート保存用のフォルダを作成
'その中にディーラーごとのフォルダを生成する
Function inputSh_makeDir(ByRef rootDir As String) As Boolean
    Dim bRet As Boolean
    Dim di As New clDir
    Dim parentDir As String
    Dim i As Long
    Dim newPath As String
    'デスクトップに親フォルダ(フォルダ名：InputSheets)を作る
    Dim WSH As Variant
    Dim dtp As String
    Set WSH = CreateObject("Wscript.Shell")
    dtp = WSH.SpecialFolders("Desktop")
    Set WSH = Nothing
    bRet = di.createFolder(dtp, ROOT, parentDir)
    'シート：TOOL内のディーラー名を読み取って、ディーラごとのフォルダを親フォルダ(フォルダ名：InputSheets)内に作成
    Dim dealers As New Collection
    Dim da As New clDB
    bRet = da.getDataColl(dbnum.dealer_name, dealers)
    For i = 1 To dealers.count Step 1
        bRet = di.createFolder(parentDir, dealers(i), newPath)
    Next i
    rootDir = parentDir
End Function

'====================
'パーツマスターシート生成処理　メイン関数
'====================
Sub createPartsMasterSheet(ByVal rootPath As String, _
                                            ByVal sheetName As String)
    Application.ScreenUpdating = False
    'パーツ構成マスターファイルの一覧を更新
    Call updateConfMasterList(rootPath)
    'パーツ構成マスターファイルからデータを取得
    Dim retTmpBucket As Variant
    Call getDataInConfMaster(retTmpBucket)
    'パーツマスターシートの初期化
    Dim sh As New clSheet
    Dim bRet As Boolean
    bRet = sh.initSheet(ThisWorkbook, PMASTER)
    'ヘッダー部分を書き込み
    Call setHeader(sheetName)
    'データ部分
    Dim dat As Variant
    Call setPartsListData(sheetName, retTmpBucket)
    Application.ScreenUpdating = True
End Sub

'パーツリスト内容をパーツリストシートに書き込む
Private Sub setPartsListData(ByVal sheetName As String, _
                                            ByVal dat As Variant)
    Dim bRet As Boolean
    Dim sh As New clSheet
    Dim row As Long
    Dim col As Long
    row = UBound(dat, 1)
    col = UBound(dat, 2)
    With ThisWorkbook.Sheets(sheetName)
        'データ設定
        .Range(Cells(2, 1), Cells(row + 1, col)) = dat
        'UGL設定部分の色変更
        .Range(Cells(2, col + 1), Cells(row + 1, col + 4)).Interior.Color = RGB(153, 255, 153)
        '罫線
        With .Range(Cells(2, 1), Cells(row + 1, col + 4))
           With .Borders(xlInsideHorizontal)
               .LineStyle = xlDash
               .Weight = xlThin
               .ColorIndex = xlAutomatic
           End With
           With .Borders(xlEdgeBottom)
               .LineStyle = xlContinuous
               .Weight = xlThin
               .ColorIndex = xlAutomatic
           End With
        End With
    End With
End Sub

'パーツリストのヘッダーをパーツリストに書き込む
Private Sub setHeader(ByVal sheetName As String)
    Dim header As Variant
     With ThisWorkbook.Sheets(sheetName)
         .Select
         header = Array("メーカー", "パーツNo", "パーツ名称", _
                                         "適用Region", "適用規格", "Remarks", _
                                         "UGL備考", "UGL変更履歴", _
                                         "UGL販売価格", "UGL管理No")
         With .Range("A1:J1")
             .Value = header
             .Font.Color = RGB(255, 255, 255)
             .Font.Bold = True
             .HorizontalAlignment = xlCenter
         End With
         .Range("A1:F1").Interior.Color = RGB(128, 128, 128)
         .Range("G1:J1").Interior.Color = RGB(0, 102, 0)
     End With
End Sub

'パーツ構成マスターファイルからデータを取得
'パーツ構成マスターファイルを巡回し、データを取得しArrayに設定する。
Private Sub getDataInConfMaster(ByRef dat As Variant)
    Dim bRet As Boolean
    Dim filesColl As Collection
    Dim foldersColl As Collection
    Dim db As New clDB
    Dim i As Long
    Dim index As Long
    Dim lastIndex As Long
    Dim ignoreNames As New Collection
    Dim targetNames As New Collection
    Dim wb As Workbook
    Dim fls As New clFiles
    Dim shs As New clSheets
    Dim sh As New clSheet
    Dim shName
    Dim row As Long
    Dim col As Long
    Dim da As New clDatArr
    Dim retTmpBucket As Variant
    ReDim retTmpBucket(1 To MAX_ROW, 1 To MAX_COL)
    
    'DBシートからパーツ構成ファイルを取得
    bRet = db.getDataColl(dbnum.confmaster_orgPath, filesColl)
    'DBシートからパーツ構成ファイルのフォルダ名を取得
    bRet = db.getDataColl(dbnum.confmaster_foldername, foldersColl)
    '=======================
    'Sheet names to ignore
    ignoreNames.Add ("tool")
    ignoreNames.Add ("$")
    ignoreNames.Add ("ugl-")
    '=======================
    index = 1
    For i = 1 To filesColl.count Step 1
        'workbookオブジェクトを取得
        bRet = fls.getWorkbookObj(filesColl(i), wb)
        
        '検索対象のシート名を取得
        Set targetNames = New Collection
        bRet = shs.getTargetSheets(wb, ignoreNames, targetNames)
        
        'シートからデータを取得
        For Each shName In targetNames
            'データ取得
            bRet = sh.getDataAsArray(wb, shName, confmasterSh.datRowS, 0, _
                                                            confmasterSh.datColS, confmasterSh.datColE, _
                                                            dat, row, col)
            '1行目(Ref.No）を削除
            bRet = da.removeColFromArray(dat, 1, dat)
            'Emptyだけのレコードを削除
            bRet = da.removeEmptyRecord(dat, dat)
            'まったく同じレコードがあれば削除。
            bRet = da.removeDuplication(dat, dat)
            '1列目にフォルダ名を挿入
            bRet = da.insertColIntoArray(dat, 1, foldersColl(i), dat)
            '取得したデータをbucketに入れる
            bRet = da.addArray(dat, index, retTmpBucket, lastIndex)
            index = lastIndex + 1
        Next shName
        'オブジェクトを閉じる
        wb.Close savechanges:=False
    Next i
    'バケツの不要なエリアを削除
    bRet = da.formatArray(retTmpBucket, lastIndex, UBound(dat, 2), retTmpBucket)
    dat = retTmpBucket
End Sub

'パーツ構成マスターファイルの一覧を更新
'パーツマスターシートを生成する際、パーツ構成マスターファイルを巡回する。その際に
'DBシートへパーツ構成マスターファイルのフルパスのリストを作成すると同時に、
'配布用インプットシートの表の「パーツ構成マスターオリジナルファイル」のリストを更新する。
Private Sub updateConfMasterList(ByVal rootPath As String)
    Dim filesColl As New Collection
    Dim foldersColl As New Collection
    Dim filenamesColl As New Collection
    Dim bRet As Boolean
    Dim fls As New clFiles
    Dim filenamesArr As Variant
    Dim filenamesCount As Long
    Dim db As New clDB
    Dim wb As Workbook
    Dim sh As New clSheet
    Dim lastRow As Long
    Dim ax As New clAxCtrl
    'パーツ構成マスタファイルのフルパスを取得
    bRet = fls.getAllXlsFilePathCol(rootPath, filesColl)
    'フルパスからフォルダ名と、ファイル名を取得
    bRet = fls.getFolderAndFileNameColl(filesColl, foldersColl, filenamesColl)
    'それぞれをDBシートへ出力(後から使うため、DBシートへ保存)
    bRet = db.setDataColl(dbnum.confmaster_orgPath, filesColl)
    bRet = db.setDataColl(dbnum.confmaster_foldername, foldersColl)
    bRet = db.setDataColl(dbnum.confmaster_filename, filenamesColl)
    'リスト部分全体の初期化
    Set wb = ThisWorkbook
    With wb.Sheets(TOOL).Range(Cells(toolSh.rowUL, toolSh.colUL), _
                    Cells(toolSh.rowLR, toolSh.colLR))
        .Clear
        .Interior.Color = RGB(0, 32, 96)
    End With
    '既に配置されているチェックボックスを削除
    bRet = sh.deleteObjectInRange(ThisWorkbook, TOOL, _
                                                    toolSh.rowUL, toolSh.colUL, _
                                                    toolSh.rowLR, toolSh.colLR)
    'パーツ構成マスターオリジナルファイルリスト表示
    bRet = db.getDataArr(dbnum.confmaster_filename, filenamesArr)
    lastRow = UBound(filenamesArr)
    With wb.Sheets(TOOL).Range(Cells(toolSh.rowUL, toolSh.colUL), _
                           Cells(toolSh.rowUL + lastRow - 1, toolSh.colUL))
        .Value = filenamesArr
        .Interior.Color = RGB(255, 255, 255)
        .Font.Size = 12
    End With
    '配布時ファイル名部分の設定
    With wb.Sheets(TOOL).Range(Cells(toolSh.rowUL, toolSh.colUL + 1), _
                           Cells(toolSh.rowUL + lastRow - 1, toolSh.colUL + 1))
        .Interior.Color = RGB(255, 255, 153)
        .Font.Size = 12
    End With
    'ディーラ情報領域の設定
    Dim i As Long
    For i = 0 To 6 Step 1
        bRet = ax.putChkBoxesV(ThisWorkbook, TOOL, _
                                            toolSh.rowUL, toolSh.colUL + 2 + i, _
                                            toolSh.colUL + 2 + i, lastRow)
    Next i
    With wb.Sheets(TOOL).Range(Cells(toolSh.rowUL, toolSh.colUL + 2), _
                    Cells(toolSh.rowUL + lastRow - 1, toolSh.colLR))
        'フォントの色を背景色に合わせる
        .Font.Color = RGB(0, 32, 96)
        '条件付き書式を設定　セルの値がTrueだったら黄色く塗りつぶす
        .FormatConditions.Add(Type:=xlCellValue, _
                Operator:=xlGreaterEqual, Formula1:=True).Interior.Color = vbYellow
    End With
    '罫線を作図
    With wb.Sheets(TOOL).Range(Cells(toolSh.rowUL, toolSh.colUL), _
                    Cells(toolSh.rowUL + lastRow - 1, toolSh.colLR - 1))
            With .Borders(xlEdgeTop)
                .LineStyle = xlDash
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlDash
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlDash
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
    End With
End Sub
