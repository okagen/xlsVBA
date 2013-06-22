Attribute VB_Name = "g_list"
Option Explicit

'グローバル変数、定数の設定用モジュール
'イニシャライズが必要なグローバル変数、定数に関してはThisWorkbookクラスにてイニシャライズする

Public Const MAX_ROW As Long = 65536     '最大行番号
Public Const MAX_COL As Long = 100          '最大列番号
Public Const TOOL As String = "$tool"           'ツールシート名
Public Const INPSH As String = "$allInptSheets" '全インプットシートを保存するフォルダの名前
Public Const PMASTER As String = "$PartsMaster" 'パーツマスターシート名
Public g_dealers As New Collection                      'ディーラー名のCollection
Public g_desktop As String                          'デスクトップパス

'DBシートのフォーマット情報
Public Const TOOLDB As String = "$$$db"
Enum dbnum
    confmaster_orgPath = 1
    confmaster_foldername = 2
    confmaster_filename = 3
End Enum

'パーツ構成マスターシートのフォーマット情報
Enum confmasterSh
    'データの領域はN3からスタート
    datRowS = 6
    datColS = 14
    'データの領域はS（19列）まで
    datColE = 19
End Enum

'$toolシートのフォーマット情報
 Enum toolSh
 '$toolシート内のインプットシートリストの範囲
    rowUL = 22
    colUL = 3
    rowLR = 36
    colLR = 12
End Enum

Public Sub init()
    Dim sh As New clSheet
    Dim da As New clDatArr
    Dim bRet As Boolean
    
    'ディーラ名を取得
    Dim dealers As Variant
    Dim col As Long
    Dim row As Long
    bRet = sh.getAllDataAsArray(ThisWorkbook, TOOL, 21, 21, 5, 11, dealers, row, col)
    bRet = da.cnvArrToColl(dealers, g_dealers)
    
    '配布用シートに含める列を選択するためのリストボックスの初期設定
    Dim lstBx As MSForms.ListBox
    Set lstBx = ThisWorkbook.Sheets(TOOL).ListBox_addColumn
    With lstBx
        If .ListCount < 1 Then
            .AddItem ("UGL備考")
            .AddItem ("UGL変更履歴")
            .AddItem ("UGL販売価格")
            .AddItem ("UGL管理No")
        End If
    End With
    
    'デスクトップに親フォルダ(フォルダ名：InputSheets)を作る
    Dim WSH As Variant
    Set WSH = CreateObject("Wscript.Shell")
    g_desktop = WSH.SpecialFolders("Desktop")
    Set WSH = Nothing
End Sub


