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
'  [i]shName    checkBoxを配置するシート名
'  [i]rowS      配置する最初の行
'  [i]colVal    リンクするセルの列
'  [i]colCtrl   checkBoxを配置する列
'  [i]count     配置する数
'--------------------------------------------------
Public Function putChkBoxesV(ByVal shName As String, _
                                ByVal rowS As Long, _
                                ByVal colVal As Long, _
                                ByVal colCtrl As Long, _
                                ByVal count As Long) As Boolean
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
    
    With Worksheets(shName)
        .Select
    
        For i = 1 To count
            'checkBoxを配置するセルを取得
            Set rngCtrl = .Range(.Cells(rowS + i - 1, colCtrl), .Cells(rowS + i - 1, colCtrl))
            
            'checkBoxの値と連動させるセルを取得
            Set rngVal = .Range(.Cells(rowS + i - 1, colVal), .Cells(rowS + i - 1, colVal))
            
            With .OLEObjects.Add(ClassType:="Forms.CheckBox.1")
                .name = "chkBox_" & i
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

    '画面更新をONに戻す
    Application.ScreenUpdating = True
    
    putChkBoxesV = True
End Function

