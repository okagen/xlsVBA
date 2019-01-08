VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clAxCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================
'ActiveXコントロールのcheckBoxを指定列に指定数配置する。
'checkBoxの値と、指定列のセルの値とリンクさせた状態。
'  [i]wb        対象のWorkbook
'  [i]shName    checkBoxを配置するシート名
'  [i]rowSt      配置する最初の行
'  [i]colVal    リンクするセルの列
'  [i]colCtrl   checkBoxを配置する列
'  [i]count     配置する数
'--------------------------------------------------
Public Function putChkBoxesV(ByVal wb As Workbook, _
                                ByVal shName As String, _
                                ByVal rowSt As Long, _
                                ByVal colVal As Long, _
                                ByVal colCtrl As Long, _
                                ByVal Count As Long) As Boolean
    Dim objChkBox As OLEObject
    Dim i As Long
    Dim rngCtrl As Range
    Dim rngVal As Range
    Dim chkBoxHeight As Integer
    Dim chkBoxWidth As Integer
    
    'checkBoxのサイズを指定
    chkBoxHeight = 15
    chkBoxWidth = 15
    
    '画面更新を一時的にOFF
    Application.ScreenUpdating = False
    
    With wb.Worksheets(shName)
        .Select
    
        For i = 1 To Count
            'checkBoxを配置するセルを取得
            Set rngCtrl = .Range(.Cells(rowSt + i - 1, colCtrl), .Cells(rowSt + i - 1, colCtrl))
            
            'checkBoxの値と連動させるセルを取得
            Set rngVal = .Range(.Cells(rowSt + i - 1, colVal), .Cells(rowSt + i - 1, colVal))
            
            With .OLEObjects.Add(ClassType:="Forms.CheckBox.1")
                '下の1行を使ってオブジェクトの名前を設定すると、この関数を実行後、他の処理が動かなくなる。
                '.name = "chkBox_" & i
                'セルの中央に配置
                .Left = rngCtrl.Left + rngCtrl.Width / 2 - chkBoxWidth / 2
                .Top = rngCtrl.Top + rngCtrl.Height / 2 - chkBoxHeight / 2
                .Height = 16
                .Width = 15
                .LinkedCell = Replace(rngVal.Address, "$", "")
                .Object.Caption = ""
                .Object.Value = False
            End With
        Next
    
    End With
    
    Set rngCtrl = Nothing
    Set rngVal = Nothing

    '画面更新をONに戻す
    Application.ScreenUpdating = True
    
    putChkBoxesV = True
End Function

