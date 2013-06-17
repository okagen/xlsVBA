Attribute VB_Name = "ctrls"
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
Sub putChkBoxes(ByVal shName As String, _
                ByVal rowS As Long, _
                ByVal colVal As Long, _
                ByVal colCtrl As Long, _
                ByVal count As Long)
                
    Dim objChkBox As OLEObject
    Dim i As Long
    Dim rngCtrl As Range
    Dim rngVal As Range
    
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
                .name = "chk_" & i
                .Left = rngCtrl.Left
                .Top = rngCtrl.Top
                .LinkedCell = Replace(rngVal.Address, "$", "")
                .Object.Caption = "chk_" & i
                .Object.Value = False
            End With
        Next
    
    End With

    '画面更新をONに戻す
    Application.ScreenUpdating = True

End Sub
