Attribute VB_Name = "verify"
Option Explicit

'*** clAxCtrl内メソッド ***
'==================================================
Sub verify_clAxCtrl_putChkBoxesV()

    Dim name As String
    Dim rowS As Long
    Dim colVal As Long
    Dim colCtrl As Long
    Dim count As Long
    Dim sh As New clSheet
    Dim ax As New clAxCtrl
    Dim bRet As Boolean
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "$verify"
    rowS = 3
    colVal = 4
    colCtrl = 3
    count = 20
    
    Set wb = ThisWorkbook
    '=======================
    
    bRet = sh.existSheet(wb, name)
    
    If bRet Then
        bRet = sh.initSheet(wb, name)
    Else
        bRet = sh.newSheet(wb, name)
    End If
    
    'put check boxes on the seet
    bRet = ax.putChkBoxesV(wb, name, rowS, colVal, colCtrl, count)
    Debug.Print "result ::: putChkBoxesV done-->" & name & " |" & Now

End Sub

'*** clDatArr内メソッド ***
'==================================================
Sub verify_clDatArr_removeDuplication()
    Dim datBucket As Variant
    Dim row As Long
    Dim col As Long
    Dim sh As New clSheet
    Dim datArr As New clDatArr
    Dim bRet As Boolean
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook

    '=======================
    'create Array for test
    ReDim datBucket(1 To 15, 1 To 15)
    
    For row = 1 To UBound(datBucket, 1) Step 1
        For col = 1 To UBound(datBucket, 2) Step 1
            datBucket(row, col) = "org(" & row Mod 3 & "," & col Mod 3 & ")"
            datBucket(row, 2) = row
        Next col
    Next row
    '=======================
    
    bRet = datArr.removeDuplication(datBucket, newDat, 2)
    If Not bRet Then
        Exit Sub
    End If
    
    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
    
End Sub


'==================================================
Sub verify_clDatArr_removeEmptyRecord()
    Dim datBucket As Variant
    Dim row As Long
    Dim col As Long
    Dim sh As New clSheet
    Dim datArr As New clDatArr
    Dim bRet As Boolean
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook

    '=======================
    'create Array for test
    ReDim datBucket(1 To 15, 1 To 15)
    
    For row = 1 To UBound(datBucket, 1) Step 2
        For col = 1 To UBound(datBucket, 2) Step 2
            datBucket(row, col) = "org(" & row Mod 3 & "," & col Mod 3 & ")"
        Next col
    Next row
    '=======================
    
    bRet = datArr.removeEmptyRecord(datBucket, newDat)
    If Not bRet Then
        Exit Sub
    End If
    
    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clDatArr_removeColFromArray()
    Dim datBucket(1 To 10, 1 To 10)
    Dim row As Long
    Dim col As Long
    Dim datArr As New clDatArr
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim colIndex As Long
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook

    '=======================
    'create Array for test
    For row = 1 To 10 Step 1
        For col = 1 To 10 Step 1
            datBucket(row, col) = "org(" & row & "," & col & ")"
        Next col
    Next row
    
    colIndex = 0
    '=======================
    
    bRet = datArr.removeColFromArray(datBucket, colIndex, newDat)
    If Not bRet Then
        Exit Sub
    End If
    
    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
    
End Sub
'==================================================
Sub verify_clDatArr_insertColIntoArray()
    Dim datBucket(1 To 10, 1 To 10)
    Dim row As Long
    Dim col As Long
    Dim datArr As New clDatArr
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim colIndex As Long
    Dim val As String
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook

    '=======================
    'create Array for test
    For row = 1 To 10 Step 1
        For col = 1 To 10 Step 1
            datBucket(row, col) = "org(" & row & "," & col & ")"
        Next col
    Next row
    
    colIndex = 3
    val = "Value col=" & colIndex
    '=======================
    
    bRet = datArr.insertColIntoArray(datBucket, colIndex, val, newDat)
    If Not bRet Then
        Exit Sub
    End If
    
    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clDatArr_formatArray()
    Dim datBucket(1 To 100, 1 To 100)
    Dim datArr As New clDatArr
    Dim sh As New clSheet
    Dim row As Long
    Dim col As Long
    Dim newRow As Long
    Dim newCol As Long
    Dim newDat As Variant
    Dim retRow As Long
    Dim retCol As Long
    Dim wb As Workbook
    Dim bRet As Boolean

    '=======================
    'create Array for test
    For row = 1 To 100 Step 1
        For col = 1 To 100 Step 1
            datBucket(row, col) = "data in Bucket(" & row & "," & col & ")"
        Next col
    Next row
    
    newRow = 10
    newCol = 12
    '=======================
    
    bRet = datArr.formatArray(datBucket, newRow, newCol, newDat)

    retRow = UBound(newDat, 1)
    retCol = UBound(newDat, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = newDat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clDatArr_addArray()
    Dim datA As Variant
    Dim datB As Variant
    Dim datBucket As Variant
    Dim row As Long
    Dim col As Long
    Dim datArr As New clDatArr
    Dim lastIndex As Long
    Dim wb As Workbook
    Dim retRow As Long
    Dim retCol As Long
    Dim bRet As Boolean
    Dim sh As New clSheet
    
    ReDim datA(1 To 5, 1 To 10)
    ReDim datB(1 To 5, 1 To 10)
    ReDim datBucket(1 To 100, 1 To 10)
    
    '=======================
    'create Array for test
    For row = 1 To 5 Step 1
        For col = 1 To 10 Step 1
            datA(row, col) = "data_A(" & row & "," & col & ")"
            datB(row, col) = "data_B(" & row & "," & col & ")"
        Next col
    Next row
    '=======================
    
    bRet = datArr.addArray(datA, 1, datBucket, lastIndex)
    bRet = datArr.addArray(datB, lastIndex + 1, datBucket, lastIndex)
    
    retRow = UBound(datBucket, 1)
    retCol = UBound(datBucket, 2)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(retRow, retCol)) = datBucket
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clDatArr_cnvCollToArr()

    Dim coll As New Collection
    Dim i As Long
    Dim arrR As Variant
    Dim arrC As Variant
    Dim da As New clDatArr
    Dim bRet As Boolean
    Dim wb As Workbook
    Dim sh As New clSheet
    Dim isR_Y As Boolean
    Dim isR_N As Boolean
    
    '=======================
    'arguments
    For i = 1 To 10 Step 1
        coll.Add ("dat" & i)
    Next i
    isR_Y = True
    isR_N = False
    '=======================
    
    bRet = da.cnvCollToArr(coll, isR_Y, arrR)
    bRet = da.cnvCollToArr(coll, isR_N, arrC)

    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(1, coll.count)) = arrR
            .Range(Cells(3, 1), Cells(coll.count + 2, 1)) = arrC
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clDatArr_cnvArrToColl()
    Dim datBucket As Variant
    Dim datArr As New clDatArr
    Dim sh As New clSheet
    Dim row As Long
    Dim col As Long
    Dim wb As Workbook
    Dim bRet As Boolean
    Dim i As Long

    '=======================
    'create Array for test
    ReDim datBucket(1 To 10, 1 To 10)
    For row = 1 To UBound(datBucket, 1) Step 1
        For col = 1 To UBound(datBucket, 2) Step 1
            datBucket(row, col) = "data(" & row & "," & col & ")"
        Next col
    Next row
    
    Dim coll As New Collection
    '=======================
    
    bRet = datArr.cnvArrToColl(datBucket, coll)

    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            For i = 1 To coll.count Step 1
                .Range(Cells(i, 1), Cells(i, 1)) = coll(i)
            Next i
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'*** clDB内メソッド ***
'==================================================
Sub verify_cldB_initDB()
    Dim db As New clDB
    Dim bRet As Boolean
    
    bRet = db.initDB()
End Sub

'==================================================
Sub verify_cldB_setDataColl()
    Dim db As New clDB
    Dim i As Long
    Dim coll As New Collection
    Dim bRet As Boolean
    
    For i = 1 To 5 Step 1
        coll.Add ("db:" & i)
    Next i
    
    bRet = db.setDataColl(2, coll)
End Sub

'==================================================
Sub verify_cldB_setDataArr()
    Dim db As New clDB
    Dim i As Long
    Dim arr As Variant
    Dim bRet As Boolean
    
    ReDim arr(1 To 10, 1 To 1)
    For i = 1 To 10 Step 1
        arr(i, 1) = "db:" & i
    Next i
    
    bRet = db.setDataArr(2, arr)
End Sub

'==================================================
Sub verify_cldB_getDataColl()
    Dim db As New clDB
    Dim coll As New Collection
    Dim bRet As Boolean
    
    bRet = db.getDataColl(2, coll)
    
    bRet = db.setDataColl(3, coll)
End Sub

'==================================================
Sub verify_cldB_getDataArr()
    Dim db As New clDB
    Dim arr As Variant
    Dim bRet As Boolean
    
    bRet = db.getDataArr(2, arr)
    
    bRet = db.setDataArr(3, arr)
End Sub

'*** clDir内メソッド ***
'==================================================
Sub verify_clDir_createFolder()
    Dim di As New clDir
    Dim WSH As Variant
    Set WSH = CreateObject("Wscript.Shell")
    Dim parent As String
    Dim folder As String
    Dim newPath As String
    Dim bRet As Boolean
    
    '=======================
    'arguments
    parent = WSH.SpecialFolders("Desktop")
    folder = "sample"
    '=======================
    
    bRet = di.createFolder(parent, folder, newPath)
    
    If bRet Then
        Debug.Print "result ::: done -> " & newPath & " |" & Now
    Else
        Debug.Print "err ::: cannot create the folder ->" & folder & " |" & Now
    End If
End Sub

'*** clFiles内メソッド ***
'==================================================
Sub verify_clFiles_copyFiles()
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim fromPath As String
    Dim toPath As String
    Dim dicFileName As Variant
    Set dicFileName = CreateObject("Scripting.Dictionary")
    
    '=======================
    'arguments
    fromPath = "C:\Users\10007434\Desktop\InputSheets\$InputSheets"
    toPath = "C:\Users\10007434\Desktop\InputSheets\ディーラA"
    dicFileName.Add "BC-10 ver1.00.xls", "BC10.xls"
    dicFileName.Add "F-45N ver2.00.xlsx", "F45N.xls"
    '=======================
    
    bRet = fls.copyFiles(fromPath, toPath, dicFileName)
    
    Debug.Print "result ::: done " & " |" & Now
End Sub

'==================================================
Sub verify_clFiles_getWorkbookObj()
    Dim file As String
    Dim wb As Workbook
    Dim fls As New clFiles
    Dim bRet As Boolean
    
    '=======================
    'The file name for test
    file = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master\インターテクノ\BC-10 ver1.00.xls"
    '=======================
    
    Application.ScreenUpdating = False
    
    'get Workbook object
    bRet = fls.getWorkbookObj(file, wb)
    
    If bRet Then
        wb.Close savechanges:=False
    End If

    Application.ScreenUpdating = True

End Sub

'==================================================
Sub verify_clFiles_getFolderAndFileNameColl()
    Dim Path As String
    Dim pathColl As New Collection
    Dim folderColl As New Collection
    Dim nameColl As New Collection
    Dim sh As New clSheet
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim i As Long
    Dim wb As Workbook
    Dim item As Variant
   
    '=======================
    'The Sheet name for test
    Path = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master"
    '=======================
    
    'set filter on the sheet
    bRet = fls.getAllXlsFilePathCol(Path, pathColl)
    
    'set filter on the sheet
    bRet = fls.getFolderAndFileNameColl(pathColl, folderColl, nameColl)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            i = 1
            For Each item In folderColl
                .Range(Cells(i, 1), Cells(i, 1)) = item
                i = i + 1
            Next item
            i = 1
            For Each item In nameColl
                .Range(Cells(i, 2), Cells(i, 2)) = item
                i = i + 1
            Next item
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clFiles_getFolderAndFileNameArr()
    Dim Path As String
    Dim col As New Collection
    Dim dat As Variant
    Dim sh As New clSheet
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim row As Long
    Dim i As Long
    Dim wb As Workbook

    
    '=======================
    'The Sheet name for test
    Path = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master"

    '=======================
    
    'set filter on the sheet
    bRet = fls.getAllXlsFilePathCol(Path, col)
    
    'set filter on the sheet
    bRet = fls.getFolderAndFileNameArr(col, dat)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(col.count, 2)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clFiles_getFolderAndFileName()
    Dim Path As String
    Dim folder As String
    Dim file As String
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim extFlg As Boolean
    
    '=======================
    'The Sheet name for test
    Path = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master\タイヨウ\F-45N ver2.00.xlsx"
    extFlg = False
    '=======================
    
    'set filter on the sheet
    bRet = fls.getFolderAndFileName(Path, extFlg, folder, file)
    
    Debug.Print "result ::: folder -> "; folder & "   file -> " & file & " |" & Now
    
End Sub

'==================================================
Sub verify_clFiles_getAllXlsFilePathCol()
    Dim Path As String
    Dim dat As New Collection
    Dim sh As New clSheet
    Dim fls As New clFiles
    Dim bRet As Boolean
    Dim i As Long
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    Path = "C:\Users\10007434\Desktop\my prj\excel_vba\sample_config_master"
    '=======================
    
    'set filter on the sheet
    bRet = fls.getAllXlsFilePathCol(Path, dat)
    
    If bRet = True Then
        Set wb = ThisWorkbook
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            For i = 1 To dat.count
                .Range(Cells(i, 1), Cells(i, 1)).Value = dat(i)
            Next
            
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub



'*** clSheet内メソッド ***
'==================================================
Sub verify_clSheet_setFilter()
    Dim name As String
    Dim sh As New clSheet
    Dim tgtFields As Variant
    Set tgtFields = CreateObject("Scripting.Dictionary")
    Dim wb As Workbook
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    tgtFields.Add 1, "F45N"
    tgtFields.Add 2, "*Table*"
    tgtFields.Add 3, "13"
    
    Set wb = ThisWorkbook
    '=======================
    
    'set filter on the sheet
    bRet = sh.setFilter(wb, name, 6, 1, 7, tgtFields)
    
End Sub


'==================================================
Sub verify_clSheet_getRowDataVLookUp()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim bRet As Boolean
    Dim str As String
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 1
    str = "45"
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.getRowDataVLookUp(wb, name, 6, 1, 7, col, str, dat, row)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, 7)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub

'==================================================
Sub verify_clSheet_deleteColData()
    Dim name As String
    Dim sh As New clSheet
    Dim col As Long
    Dim wb As Workbook
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 2
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.deleteColData(wb, name, 6, col)
    wb.Sheets(name).Select
End Sub


'==================================================
Sub verify_clSheet_getColDataAsArray()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim allowDup As Boolean
    Dim wb As Workbook
    Dim bRet As Boolean
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    col = 1
    allowDup = False
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.getColDataAsArray(wb, name, 6, col, allowDup, dat, row)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, 1)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


'==================================================
Sub verify_clSheet_getAllDataAsArray()
    Dim name As String
    Dim sh As New clSheet
    Dim dat As Variant
    Dim col As Long
    Dim row As Long
    Dim wb As Workbook
    Dim bRet As Boolean
    Dim lastRow As Long
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    lastRow = 10
    Set wb = ThisWorkbook
    '=======================
    
    'get data in the sheet
    bRet = sh.getAllDataAsArray(wb, name, 6, lastRow, 1, 7, dat, row, col)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(row, col)) = dat
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
    'The Sheet name for test
    name = "sample5"
    
    Set wb = ThisWorkbook
    '=======================
    
    bRet = sh.existSheet(wb, name)
    
    If bRet Then
        bRet = sh.initSheet(wb, name)
        Debug.Print "result ::: initSheet done-->" & name & " |" & Now
    Else
        Debug.Print "result ::: err-->" & name & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_newSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim newName As String
    Dim wb As Workbook
     
    '=======================
    'The Sheet name for test
    name = "sample5"
    
    Set wb = ThisWorkbook
    '=======================
    
    'get new sheet name
    newName = sh.newSheet(wb, name)
    
    Debug.Print "result ::: sheet name is-->" & newName & " |" & Now
End Sub

'==================================================
Sub verify_clSheet_existSheet()
    Dim name As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    
    '=======================
    'The Sheet name for test
    name = "sample1"
    
    Set wb = ThisWorkbook
    '=======================
    
    'check existance of the sheet
    bRet = sh.existSheet(wb, name)
    
    If bRet Then
        Debug.Print "result ::: exist-->" & name & " |" & Now
    Else
        Debug.Print "result ::: N/A-->" & name & " |" & Now
    End If
    
End Sub

'==================================================
Sub verify_clSheet_deleteObjectInRange()
    Dim name As String
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim wb As Workbook
    Dim rowUL As Long
    Dim colUL As Long
    Dim rowLR As Long
    Dim colLR As Long
    
    '=======================
    'function's arguments for test
    Set wb = ThisWorkbook
    name = "$verify"
    rowUL = 1
    colUL = 1
    rowLR = 16
    colLR = 10
    '=======================
    
    'check existance of the sheet
    wb.Worksheets(name).Select
    bRet = sh.deleteObjectInRange(wb, name, rowUL, colUL, rowLR, colLR)
   
End Sub


'*** clSheets内メソッド ***
'==================================================
Sub verify_clSheets_getTargetSheets()
    Dim shs As New clSheets
    Dim ignoreNames As New Collection
    Dim targetNames As New Collection
    Dim bRet As Boolean
    Dim name As Variant
    Dim wb As Workbook
    
    '=======================
    'Sheet names to ignore
    ignoreNames.Add ("tool")
    ignoreNames.Add ("$")
    ignoreNames.Add ("ugl-")
    
    Set wb = ThisWorkbook
    '=======================
    
    'get target sheet names
    bRet = shs.getTargetSheets(wb, ignoreNames, targetNames)
    
    For Each name In targetNames
        Debug.Print "result ::: done " & name & " |" & Now
    Next
    
End Sub

'==================================================
Sub verify_clSheets_conbineSheets()
    Dim names As New Collection
    Dim sh As New clSheet
    Dim shs As New clSheets
    Dim bRet As Boolean
    Dim dat As Variant
    Dim row As Long
    Dim arrRow As Long
    Dim arrCol As Long
    Dim wb As Workbook
    
    '=======================
    'The Sheet names to test
    names.Add ("sample1")
    names.Add ("sample2")
    names.Add ("sample3")
    
    Set wb = ThisWorkbook
    '=======================
    
    'get data in sheets
    bRet = shs.combineSheets(wb, names, 6, 1, 7, dat, row)
    
    'get number of row and column from array object.
    arrRow = UBound(dat, 1)
    arrCol = UBound(dat, 2)
    
    If bRet = True Then
        'initialize the sheet to verification
        bRet = sh.initSheet(wb, "$verify")
        'plot all data on the $verify sheet
        With wb.Sheets("$verify")
            .Select
            .Range(Cells(1, 1), Cells(arrRow, arrCol)) = dat
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub
