VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================
'VLOOKUP�֐���p���āA�����V�[�g���Q�Ƃ��ăf�[�^���擾����
'  [i]wb                   �����Ώ�Workbook
'  [i]targetSheets    �����ΏۃV�[�g����Collection
'  [i]key                   �����L�[
'  [i]keyCol               �����V�[�g�� �L�[�������Ă����
'  [i]dataCol         �����V�[�g�� �l�̗�
'  [o]ret                   ��������
'--------------------------------------------------
Public Function getDataFromSheetsByVlookup(ByVal wb As Workbook, _
                                                                        ByRef targetSheets As Collection, _
                                                                        ByVal key As String, _
                                                                        ByVal keyCol As Long, _
                                                                        ByVal dataCol As Long, _
                                                                        ByRef ret As String _
                                                                        ) As Boolean
    Dim sh As New clSheet
    Dim bRet As Boolean
    Dim rngTxt As String
    Dim i As Long
    Dim result As Variant
    Dim rng As Range
    
    '2�̗�ԍ���Range��\���A���t�@�x�b�g�ɕϊ�����
    bRet = sh.colNo2Rng(keyCol, dataCol, rngTxt)
    
    '�V�[�g�𐔕����[�v
    For i = 1 To targetSheets.Count Step 1
        With wb.Worksheets(targetSheets(i))
            On Error Resume Next
            Set rng = .Range(rngTxt)
            result = WorksheetFunction.VLookup( _
                                   key, _
                                   rng, _
                                   dataCol - keyCol + 1, _
                                   False)
            If ERR <> 0 Then result = ""
        End With
        
        '�l���擾�ł�����֐��𔲂���
        If result <> "" Then
            ret = result
            getDataFromSheetsByVlookup = True
            Exit Function
        End If
    Next i
    
    getDataFromSheetsByVlookup = False
End Function

'==================================================
'��������V�[�g���������Ŏ󂯂āABook���̌����Ώ�Sheet�̖��OCollection���쐬
'  [i]wb                �Ώۂ�Workbook
'  [i]ignoreSheetNames  ��������V�[�g����������Collection
'  [o]targetSheets      �����Ώۂ̃V�[�g����������Collection
'--------------------------------------------------
Public Function getTargetSheets(ByVal wb As Workbook, _
                                ByVal ignoreSheetNames As Collection, _
                                ByRef targetSheets As Collection) As Boolean
                                
    Dim xlsheet As Worksheet
    Dim ignoreSheetName As Variant
    Dim bTgt As Boolean
    
    For Each xlsheet In wb.Worksheets
        
        bTgt = True
        
        For Each ignoreSheetName In ignoreSheetNames
            '�ΏۂȂ�΁Acollection�ɒǉ�
            If Left(xlsheet.name, Len(ignoreSheetName)) = ignoreSheetName Then
                bTgt = False
            End If
        Next
        
        If bTgt = True Then
            targetSheets.Add (xlsheet.name)
        End If
    Next
   
    If targetSheets.Count < 1 Then
        getTargetSheets = False
    Else
        getTargetSheets = True
    End If
    
End Function

'==================================================
'�����V�[�g�̒��̃f�[�^���������āAArray�Ɋi�[
'  [i]wb        �Ώۂ�Workbook
'  [i]shNames   �����Ώۂ̃V�[�g����������Collection
'  [i]datRowS   �f�[�^�̈�̃X�^�[�g�s
'  [i]datColS   �f�[�^�̈�̍ŏ��̗�i��ԍ��̗�)
'  [i]datColE   �f�[�^�̈�̍Ō�̗�i��ԉE�̗�)
'  [o]dat       ���������f�[�^
'  [o]row       �ŏI�s��
'--------------------------------------------------
Public Function combineSheets(ByVal wb As Workbook, _
                                ByVal shNames As Collection, _
                                ByVal datRowS As Long, _
                                ByVal datColS As Long, _
                                ByVal datColE As Long, _
                                ByRef dat As Variant, _
                                ByRef row As Long) As Boolean
    Dim sh As New clSheet
    Dim datArr As New clDatArr
    Dim retTmpBucket As Variant
    ReDim retTmpBucket(1 To MAX_ROW, 1 To datColE - datColS + 1)
    Dim datBySheet As Variant
    Dim rowBySheet As Long
    Dim colBySheet As Long
    Dim bRet  As Boolean
    Dim shName As Variant
    Dim indexRow As Long
    Dim lastRow As Long
    
    indexRow = 1
    For Each shName In shNames
        '�V�[�g�̗L�����`�F�b�N
        If sh.existSheet(wb, shName) Then
            '�V�[�g����Data���擾
            bRet = sh.getDataAsArray(wb, shName, datRowS, 0, datColS, datColE, datBySheet, rowBySheet, colBySheet)
            If bRet Then
                '�擾����Data��bucket�ɒǉ�
                bRet = datArr.addArray(datBySheet, indexRow, retTmpBucket, lastRow)
                If Not bRet Then
                    '�I�[�o�[�t���[����\��������̂ŏ������~
                    Exit Function
                End If
                indexRow = lastRow + 1
            End If
        End If
    Next
    
    '�Ԃ�l���쐬
    bRet = datArr.formatArray(retTmpBucket, lastRow, colBySheet, dat)
    row = lastRow
    combineSheets = True
End Function