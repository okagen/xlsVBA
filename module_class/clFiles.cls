VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================
'�w��t�H���_���̃t�@�C�����A�ʂ̃t�H���_�ɐV�����t�@�C�����ŃR�s�[����B
'  [i]fromPath            �R�s�[���̃t�@�C��������t�H���_�̃p�X
'  [i]toPath                �R�s�[��̃t�H���_�̃p�X
'  [i]dicFileNames      �R�s�[�̑O��̃t�@�C�������ݒ肳�ꂽ�n�b�V���e�[�u��CreateObject("Scripting.Dictionary")�Ő���
'--------------------------------------------------
Public Function copyFiles(ByVal fromPath As String, _
                                        ByVal toPath As String, _
                                        ByVal dicFileNames As Variant _
                                        ) As Boolean
    Dim sRet As String
    Dim i As Long
    Dim FSO As Object
    Dim fromFile As String
    Dim toFile As String
    Dim keys As Variant
    
    '�f�B���N�g�������݂��邩�`�F�b�N
    sRet = Dir(fromPath, vbDirectory)
    If sRet = "" Then
        Debug.Print "err ::: directory does not exist.(" & fromPath & ")  [func:copyFiles()]" & Now
        copyFiles = False
        Exit Function
    End If
    sRet = Dir(toPath, vbDirectory)
    If sRet = "" Then
        Debug.Print "err ::: directory does not exist.(" & toPath & ")  [func:copyFiles()]" & Now
        copyFiles = False
        Exit Function
    End If
        
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'dicFileNames��key����(�R�s�[���̃t�@�C����)��z��Ƃ��Ď擾
    keys = dicFileNames.keys()
    
    'key�̐������[�v���ĐV�t�@�C�����ŏ㏑���R�s�[����
    For i = 0 To UBound(keys) Step 1
        fromFile = FSO.BuildPath(fromPath, keys(i))
        toFile = FSO.BuildPath(toPath, dicFileNames(keys(i)))
        FSO.copyFile fromFile, toFile, True
    Next i
    
    Set FSO = Nothing
    copyFiles = True
End Function

'=============================
'�t�@�C�������󂯂�workbook�I�u�W�F�N�g���擾
'  [i]file  Excel�̃t���p�X
'  [o]wb    workbook�I�u�W�F�N�g
'  *** �擾����wb�I�u�W�F�N�g�͖{�֐��̊O��wb.close����
'--------------------------------------------------
Public Function getWorkbookObj(ByVal file As String, _
                        ByRef wb As Workbook) As Boolean
    Dim ret As Workbook
    Dim wbTmp As Workbook
    Dim buf As String
    
    '�t�@�C���̑��݊m�F
    buf = Dir(file)
    If buf = "" Then
        getWorkbookObj = False
        Exit Function
    End If
    
    '���ɊJ���Ă����ꍇ
    For Each wbTmp In Workbooks
        If wbTmp.name = buf Then
            Set wb = wbTmp
            getWorkbookObj = True
            Exit Function
        End If
    Next wbTmp
    
    Set wb = GetObject(file)
    getWorkbookObj = True
    
End Function

'=============================
'�t���p�X��Collection���󂯁A�t�@�C�����ƕۑ�����Ă���t�H���_����Collection��Ԃ�
'  [i]fullPaths Excel�̃t���p�X��~�ς���Collection
'  [o]folders       �t�@�C�����ƕۑ�����Ă���t�H���_����2����Array
'  [o]filenames       �t�@�C�����ƕۑ�����Ă���t�H���_����2����Array
'--------------------------------------------------
Public Function getFolderAndFileNameColl(ByVal fullPaths As Collection, _
                                                                ByRef folders As Collection, _
                                                                ByRef filenames As Collection _
                                                                ) As Boolean
    If fullPaths.Count < 1 Then
        getFolderAndFileNameColl = False
        Exit Function
    End If
    
    Dim retFolders As New Collection
    Dim retFiles As New Collection
    Dim item As Variant
    Dim file As String
    Dim folder As String
    Dim bRet As Boolean
  
    For Each item In fullPaths
        bRet = getFolderAndFileName(item, True, folder, file)
        If bRet Then
            retFolders.Add (folder)
            retFiles.Add (file)
        End If
    Next item
    
    Set folders = retFolders
    Set filenames = retFiles
    getFolderAndFileNameColl = True
End Function

'=============================
'�t���p�X��Collection���󂯁A�t�@�C�����ƕۑ�����Ă���t�H���_����2����Array��Ԃ�
'  [i]fullPaths Excel�̃t���p�X��~�ς���Collection
'  [o]dat       �t�@�C�����ƕۑ�����Ă���t�H���_����2����Array
'--------------------------------------------------
Public Function getFolderAndFileNameArr(ByVal fullPaths As Collection, _
                                        ByRef dat As Variant) As Boolean
    Dim retTmp As Variant
    Dim fullPath As String
    Dim file As String
    Dim folder As String
    Dim i As Long
    Dim bRet As Boolean
    
    If fullPaths.Count < 1 Then
        getFolderAndFileNameArr = False
        Exit Function
    End If
    
    ReDim retTmp(1 To fullPaths.Count, 1 To 2)
    For i = 1 To fullPaths.Count Step 1
        fullPath = fullPaths(i)
        bRet = getFolderAndFileName(fullPath, True, folder, file)
        If bRet Then
            retTmp(i, 1) = folder
            retTmp(i, 2) = file
        End If
    Next i
    
    dat = retTmp
    getFolderAndFileNameArr = True
End Function

'=============================
'�t�@�C���̃t���p�X���󂯁A�t�@�C�����ƕۑ�����Ă���t�H���_����Ԃ�
'  [i]fullPaths     Excel�̃t���p�X��~�ς���Collection
'  [i]extflg        �g���q�̗v�� true=�g���q���� false=�g���q�Ȃ�
'  [o]folder        �t�H���_��
'  [o]file            �t�@�C���� �g���q����
'--------------------------------------------------
Public Function getFolderAndFileName(ByVal fullPath As String, _
                                                            ByVal extFlg As Boolean, _
                                                                ByRef folder As String, _
                                                                ByRef file As String _
                                                                ) As Boolean
    Dim FSO As Object
    Dim bRet As Boolean
    Dim dot As Long
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    '�t�@�C�����Ȃ��ꍇ�G���[
    bRet = FSO.FileExists(fullPath)
    If Not bRet Then
        Debug.Print "err ::: file does not exist.(" & fullPath & ")  [func:insertColIntoArray()]" & Now
        getFolderAndFileName = False
        Exit Function
    End If
    
    file = FSO.GetFile(fullPath).name
    '�g���q���������t�@�C�����̎擾
    If extFlg = False Then
        dot = InStrRev(file, ".")
        file = Left(file, dot - 1)
    End If
    
    folder = FSO.getFolder(FSO.GetFile(fullPath).ParentFolder).name
    
    Set FSO = Nothing
    getFolderAndFileName = True
End Function

'=============================
'�w�肵���f�B���N�g���z���ɂ���SExcel�t�@�C���̃t���p�X�擾
'  [i]path      ��������f�B���N�g���̃p�X
'  [o]ret       Excel�̃t���p�X��~�ς���Collection
'--------------------------------------------------
Public Function getAllXlsFilePathCol(ByVal path As String, _
                                    ByRef ret As Collection) As Boolean
    Dim FSO As Object
    Dim folder As Variant
    Dim file As Variant
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each folder In FSO.getFolder(path).SubFolders
        Call getAllXlsFilePathCol(folder.path, ret)
    Next folder
    For Each file In FSO.getFolder(path).files
        If InStr(file.Type, "Excel") > 0 Then
            ret.Add (file.path)
        End If
    Next file
    Set FSO = Nothing
    
    If ret.Count < 1 Then
        getAllXlsFilePathCol = False
    Else
        getAllXlsFilePathCol = True
    End If
End Function

'=============================
'Workbook�Ɠ��f�B���N�g���ɂ���w�肵���g���q�̃t�@�C���̃t�@�C�����擾
'  [o]ret       �t�@�C������~�ς���Collection
'--------------------------------------------------
Public Function getAllFileNameCol(ByVal ext As String, _
                                    ByRef ret As Collection) As Boolean
    Dim FSO As Object
    Dim folder As Variant
    Dim file As Variant
    
    Dim path As String
    path = ThisWorkbook.path
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each folder In FSO.getFolder(path).SubFolders
        Call getAllXlsFilePathCol(folder.path, ret)
    Next folder
    For Each file In FSO.getFolder(path).files
        If InStr(1, file.name, ext, vbTextCompare) > 0 Then
            ret.Add (file.name)
        End If
    Next file
    Set FSO = Nothing
    
    If ret.Count < 1 Then
        getAllFileNameCol = False
    Else
        getAllFileNameCol = True
    End If
End Function
                            