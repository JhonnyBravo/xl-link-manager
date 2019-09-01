VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinkManager
   Caption         =   "Link Manager"
   ClientHeight    =   2040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11180
   OleObjectBlob   =   "LinkManager.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LinkManager"
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
'リンクデータを新規登録する。
'''
Private Sub btnCreate_Click()
    With Me
        If .lstTitle = "" Or .txtUrl = "" Then
            MsgBox "Title または URL が入力されていません。"
            Exit Sub
        End If
    End With

    Dim objLink As New link

    With objLink
        .title = Me.lstTitle
        .url = Me.txtUrl
    End With

    Dim objLc As CrudRepository

    Set objLc = New LinksController
    objLc.createRecord objLink
End Sub

'''
'リンクデータを削除する。
'''
Private Sub btnDelete_Click()
    Dim objLink As link
    Dim objLc As CrudRepository

    With Me
        If .lstTitle.Value = "" Then
            MsgBox "Title を入力してください。"
            Exit Sub
        End If

        Set objLink = New link
        objLink.title = .lstTitle

        Set objLc = New LinksController
        objLc.deleteRecord objLink
    End With
End Sub

'''
'条件に合致するタイトルを持つリンクデータを検索し、
'フォームの URL 入力欄へ URL を入力する。
'''
Private Sub btnFindByTitle_Click()
    Dim objLc As CrudRepository
    Dim objRange As Range

    With Me
        If .lstTitle.Value = "" Then
            MsgBox "Title を入力してください。"
            Exit Sub
        End If

        Set objLc = New LinksController
        Set objRange = objLc.findByTitle(.lstTitle)

        If objRange Is Nothing = False Then
            .txtUrl.Value = objRange.Cells(1, 3).Value
        End If
    End With
End Sub

'''
'リンクデータを更新する。
'''
Private Sub btnUpdate_Click()
    Dim objRange As Range
    Dim objLink As link
    Dim objLc As CrudRepository

    With Me
        If .lstTitle.Value = "" Or .txtUrl.Value = "" Then
            MsgBox "Title または URL が入力されていません。"
            Exit Sub
        End If

        Set objLink = New link
        objLink.title = .lstTitle.Value
        objLink.url = .txtUrl.Value

        Set objLc = New LinksController
        objLc.updateRecord objLink
    End With
End Sub

'''
'フォーム起動時の初期化処理を実行する。
'''
Private Sub UserForm_Activate()
    'フォーム入力欄の初期化処理
    With Me
        .lstTitle.Value = ""
        .txtUrl.Value = ""
    End With

    'タイトル入力欄のリスト初期化処理
    Dim objTc As New TitlesController

    objTc.initTitles
End Sub
