VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ExcelのあるシートをDBとして扱うためのクラス
'DBシート名はグローバル定数に設定

'=============================
'DBシートを作成。無ければ作る、あれば何もしない。
'--------------------------------------------------
Public Function initDB() As Boolean
    Dim sh As New clSheet
    Application.ScreenUpdating = False
    If Not sh.existSheet(ThisWorkbook, TOOLDB) Then
        ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets(Worksheets.count)
        ActiveSheet.name = TOOLDB
    End If
    Application.ScreenUpdating = True
    initDB = True
End Function

'=============================
'DBシートに値を設定。
'  [i]col  設定先の列番号
'  [i]datColl  設定するデータ
'--------------------------------------------------
Public Function setDataColl(ByVal col As Long, _
                                            ByVal datColl As Collection _
                                            ) As Boolean
    Call initDB
    
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim i As Long
    
    '既存データがあれば削除
    bRet = sh.deleteColData(ThisWorkbook, TOOLDB, 1, col)
    If bRet Then
        'データ設定
        With ThisWorkbook.Sheets(TOOLDB)
            For i = 1 To datColl.count Step 1
                '.Select
                .Range(.Cells(i, col), .Cells(i, col)) = datColl(i)
            Next i
        End With
    End If
    setDataColl = bRet
End Function

'=============================
'DBシートに値を設定。
'  [i]col  設定先の列番号
'  [i]datArr  設定するデータ
'--------------------------------------------------
Public Function setDataArr(ByVal col As Long, _
                                ByVal datArr As Variant _
                                ) As Boolean
    Call initDB
    
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim lastRow As Long
    
    '既存データがあれば削除
    bRet = sh.deleteColData(ThisWorkbook, TOOLDB, 1, col)
    If bRet Then
        lastRow = UBound(datArr, 1)
        g_dbSh.Range(Cells(1, col), Cells(lastRow, col)) = datArr
    End If
    setDataArr = bRet
End Function

'=============================
'DBシートから値をCollectionで取得。
'  [i]col  設定先の列番号
'  [o]datColl  取得データ
'--------------------------------------------------
Public Function getDataColl(ByVal col As Long, _
                                        ByRef datColl As Collection) As Boolean
    Call initDB
                                    
    Dim sh As New clSheet
    Dim retColl As New Collection
    Dim i As Long
    Dim getVal As Variant
    
    Application.ScreenUpdating = False
    
    'データ取得
    i = 1
    With ThisWorkbook.Sheets(TOOLDB)
        While True
            getVal = .Range(.Cells(i, col), .Cells(i, col)).Value
            If IsEmpty(getVal) Then
                GoTo WNEXT
            Else
                retColl.Add (getVal)
                i = i + 1
            End If
        Wend
    End With
WNEXT:
    Application.ScreenUpdating = True
    
    If retColl.count > 0 Then
        Set datColl = retColl
        getDataColl = True
    Else
        getDataColl = False
    End If
End Function

'=============================
'DBシートから値をArrayで取得。
'  [i]col  設定先の列番号
'  [o]datArr  取得データ
'--------------------------------------------------
Public Function getDataArr(ByVal col As Long, _
                                    ByRef datArr As Variant) As Boolean
    Call initDB
    Dim retArr As Variant
    Dim lastRow As Long
    
    'データ取得
    With ThisWorkbook.Sheets(TOOLDB)
        lastRow = .Range(.Cells(65536, col), .Cells(65536, col)).End(xlUp).row
        retArr = .Range(.Cells(1, col), .Cells(lastRow, col))
    End With
    
    If UBound(retArr, 1) > 0 Then
        datArr = retArr
        getDataArr = True
    Else
        getDataArr = False
    End If
End Function

