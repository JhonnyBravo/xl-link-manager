VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YahooSearch
   Caption         =   "Yahoo 検索"
   ClientHeight    =   1240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11180
   OleObjectBlob   =   "YahooSearch.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "YahooSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''
'フォームを閉じる。
'''
Private Sub btnClose_Click()
    Me.Hide
End Sub

'''
'指定した検索ワードを Yahoo にて検索し、検索結果のリンク一覧を Excel へ出力する。
'''
Private Sub btnSearch_Click()
    If Me.txtWord = "" Then
        MsgBox "Keyword を入力してください。"
        Exit Sub
    End If

    'URL と検索ワードを取得。
    Dim strUrl As String
    Dim strWord As String

    strUrl = Sheets("Config").Range("B2").Value
    strWord = Me.txtWord

    'IE 起動。
    Dim objBc As New BrowserController
    Dim objDoc As HTMLDocument

    With objBc
        .openBrowser strUrl, True
        .waitForLoading
        Set objDoc = .getDocument
    End With

    'Yahoo 検索画面へ検索ワードを入力して検索ボタンをクリック。
    Dim objInput As HTMLInputElement
    Dim objButton As HTMLButtonElement

    Set objInput = objDoc.querySelector("input#srchtxt")
    objInput.Value = strWord

    Set objButton = objDoc.querySelector("input#srchbtn")
    objButton.Click

    With objBc
        .waitForLoading
        Set objDoc = .getDocument
    End With

    '検索結果のリンク一覧を Excel へ出力する。
    Dim objElements As IHTMLElementCollection
    Dim objAnchor As HTMLAnchorElement

    Dim objRegExp As New RegExp
    Dim strHref As String

    Dim objLc As CrudRepository
    Dim objLink As link
    Dim strMsg As String

    Set objElements = objDoc.getElementsByTagName("a")
    objRegExp.Pattern = "search.yahoo|btoptout.yahoo|cache.yahoofs|javascript"

    Set objLc = New LinksController
    objLc.deleteRecordAll

    On Error GoTo catch

    For Each objAnchor In objElements
        strHref = objAnchor.getAttribute("href")

        If objRegExp.test(strHref) = False And strHref <> "" Then
            Set objLink = New link

            With objLink
                .title = objAnchor.innerText
                .url = strHref
            End With

            objLc.createRecord objLink
        End If
    Next

    strMsg = "完了しました。"
    GoTo finally
catch:
    strMsg = "エラーが発生しました。" & vbCrLf & Err.Description
    Debug.Print strMsg
finally:
    Set objBc = Nothing
    MsgBox strMsg
End Sub

'''
'フォーム初期化処理。
'''
Private Sub UserForm_Activate()
    Me.txtWord = ""
End Sub
