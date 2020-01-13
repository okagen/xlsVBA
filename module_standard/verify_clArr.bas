Attribute VB_Name = "verify_clArr"
Option Explicit
Option Base 1

'==================================================
Sub verify_clDatArr_cnvCollToArr()

    Dim coll As New Collection
    Dim i As Long
    Dim arrR As Variant
    Dim arrC As Variant
    Dim da As New clArr
    Dim bRet As Boolean
    Dim wb As Workbook
    Dim sh As New clSheet
    Dim isR_Y As Boolean
    Dim isR_N As Boolean
    
    '=======================
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
        With wb.sheets("$verify")
            .Select
            .Range(.Cells(1, 1), .Cells(UBound(arrR, 1), UBound(arrR, 2))) = arrR
            .Range(.Cells(3, 1), .Cells(UBound(arrC, 1) + 2, UBound(arrC, 2))) = arrC
            Debug.Print "result ::: done " & " |" & Now
        End With
    Else
        Debug.Print "result ::: no data" & " |" & Now
    End If
End Sub


