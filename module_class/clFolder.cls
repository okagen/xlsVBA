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
'フルパスを指定してフォルダを作成。
'同名のフォルダが存在した場合、フォルダ名末尾に(#)を付けてカウントアップ。
'  [i]path      作成するフォルダのフルパス
'  [o]newPath   作成したフォルダのフルパス
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

