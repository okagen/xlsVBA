VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================
'アドレス、タイトル、本文を設定しメーラを起動する。（添付ファイルなし）
'  [i]wb          ThisWorkbook
'  [i]toAddr    宛先メールアドレス
'  [i]subj        メールのタイトル
'  [i]bdy         メールの本文
'--------------------------------------------------
Public Function openMailer(ByVal wb As Workbook, _
                                        ByVal toAddr As String, _
                                        ByVal subj As String, _
                                        ByVal bdy As String _
                                        ) As Boolean
    Dim strLink As String
    strLink = "mailto:" & toAddr & "?subject=" & subj & "&body=" & bdy
    wb.FollowHyperlink strLink
    openMailer = True
End Function

'=============================
'アドレス、タイトル、本文、添付ファイルを設定しoutlookを起動する
'  [i]toAddr          宛先メールアドレス
'  [i]subj              メールのタイトル
'  [i]bdy               メールの本文
'  [i]attachPath    添付ファイルのフルパス
'--------------------------------------------------
Public Function openOutlook(ByVal toAddr As String, _
                                        ByVal subj As String, _
                                        Optional ByVal bdy As String = "", _
                                        Optional ByVal attachPath As String = "" _
                                        ) As Boolean
    Dim OLK As Outlook.Application
    Dim mailItem As Outlook.mailItem
    
    Set OLK = CreateObject("Outlook.Application")
    Set mailItem = OLK.CreateItem(olMailItem)

    mailItem.To = toAddr
    mailItem.CC = ""
    mailItem.BCC = ""
    mailItem.subject = subj
    mailItem.Importance = olImportanceNormal
    If attachPath <> "" Then
        mailItem.Attachments.Add attachPath
    End If
    mailItem.body = bdy

    mailItem.Display
    Set mailItem = Nothing
    Set OLK = Nothing
    openOutlook = True
End Function

