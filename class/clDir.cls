VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================
'指定ディレクトリにフォルダを作る。
'同名のフォルダが存在した場合、フォルダ名末尾に(#)を付けてカウントアップ。
'  [i]parentDir    フォルダ作成先ディレクトリのパス
'  [i]name          フォルダ名
'  [o]newDir   作成したフルパス付フォルダ名
'--------------------------------------------------
Public Function createFolder(ByVal parentDir As String, _
                                            ByVal name As String, _
                                            ByRef newDir As String _
                                            ) As Boolean
    If parentDir = "" Or name = "" Then
        createFolder = False
        Debug.Print "err ::: no input parentDir-> (" & parentDir & ")" & _
                            "/name -> (" & name & ") [func:createFolder()]" & Now
        Exit Function
    End If
                                            
    Dim orgName As String
    Dim tmpName As String
    Dim ret As String
    Dim num As Long
    Dim flg As Boolean
    
    'フォルダ名のフルパスを生成
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    orgName = FSO.BuildPath(parentDir, name)
    Set FSO = Nothing

    'フォルダの存在をチェックしてフォルダ名を生成
    tmpName = orgName
    num = 1
    flg = True
    Do While flg
        ret = Dir(tmpName, vbDirectory)
        If ret = "" Then
            flg = False
        Else
            tmpName = orgName & " " & Format(num, "(#)")
            num = num + 1
        End If
    Loop

    'フォルダ作成
    MkDir tmpName
    
    newDir = tmpName
    createFolder = True
End Function


