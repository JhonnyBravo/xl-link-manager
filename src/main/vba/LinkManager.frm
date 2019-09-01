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

Private objSheet As Worksheet

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
            .txtUrl = objRange.Cells(1, 3).Value
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
        If .lstTitle = "" Or .txtUrl = "" Then
            MsgBox "Title または URL が入力されていません。"
            Exit Sub
        End If

        Set objLink = New link
        objLink.title = .lstTitle
        objLink.url = .txtUrl

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

        If .lstTitle.ListCount > 0 Then
            .lstTitle.Clear
        End If
    End With

    'タイトル入力欄のリスト初期化処理
    Dim objRange As Range
    Dim varTitles As Variant
    Dim lngRow As Long

    Dim strStartAddr As String
    Dim strEndAddr As String

    strStartAddr = "B3"

    With objSheet.Range("B2")
        If .Cells(2, 1).Value <> "" Then
            strEndAddr = .Cells(1, 1).End(xlDown).Address
            Set objRange = objSheet.Range(strStartAddr, strEndAddr)

            If objRange.Address(False, False) = strStartAddr Then
                Me.lstTitle.AddItem objRange.Value
                Exit Sub
            End If

            varTitles = objRange

            For lngRow = 1 To UBound(varTitles)
                Me.lstTitle.AddItem varTitles(lngRow, 1)
            Next
        End If
    End With
End Sub

'''
'フォームインスタンス生成時の初期化処理
'''
Private Sub UserForm_Initialize()
    Set objSheet = ActiveWorkbook.Sheets("Links")
End Sub
