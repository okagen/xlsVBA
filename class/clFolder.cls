VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================
'�t���p�X���w�肵�ăt�H���_���쐬�B
'�����̃t�H���_�����݂����ꍇ�A�t�H���_��������(#)��t���ăJ�E���g�A�b�v�B
'  [i]path      �쐬����t�H���_�̃t���p�X
'  [o]newPath   �쐬�����t�H���_�̃t���p�X
'--------------------------------------------------
Public Function mkFolder(ByVal path As String, ByRef newPath As String) As Boolean
    Dim num As Integer
    Dim flg As Boolean

    newPathTmp = path
    num = 1
    flg = True
    
    Do While flg
        If Dir(newPathTmp, vbDirectory) = "" Then
            flg = False
        Else
            newPathTmp = path & " " & Format(num, "(#)")
            num = num + 1
        End If
    Loop
    
    MkDir newPathTmp
    newPath = newPathTmp
    mkFolder = True
    
End Function
