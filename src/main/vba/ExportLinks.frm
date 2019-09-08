VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportLinks
   Caption         =   "エクスポート"
   ClientHeight    =   1240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11180
   OleObjectBlob   =   "ExportLinks.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ExportLinks"
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
'外部ファイルへデータをエクスポートする。
'''
Private Sub btnExport_Click()
    If Me.txtPath.Value = "" Then
        MsgBox "ファイル出力先が指定されていません。 Path を確認してください。"
        Exit Sub
    End If

    'コピー対象とするレコードセットを取得する。
    Dim objXlc As CrudRepository
    Dim objRange As Range
    Dim varRecordset As Variant

    Set objXlc = New XlLinksController
    Set objRange = objXlc.findAll

    If objRange Is Nothing Then
        Exit Sub
    End If

    varRecordset = objRange

    '外部ファイルへデータをコピーする。
    Dim lngRow As Long

    Dim objCr As CrudRepository
    Dim objLink As link

    Dim objFso As New FileSystemObject

    Select Case objFso.GetExtensionName(Me.txtPath.Value)
        Case "csv"
            Set objCr = New CsvLinksController
            objCr.deleteRecordAll

            For lngRow = 1 To UBound(varRecordset)
                Set objLink = New link

                With objLink
                    .title = varRecordset(lngRow, 1)
                    .url = varRecordset(lngRow, 3)
                End With

                objCr.createRecord objLink
            Next
        Case "xlsx"
            Set objCr = New XlLinksController2
            objCr.deleteRecordAll

            For lngRow = 1 To UBound(varRecordset)
                Set objLink = New link

                With objLink
                    .title = varRecordset(lngRow, 1)
                    .url = varRecordset(lngRow, 3)
                End With

                objCr.createRecord objLink
            Next
        Case "accdb"
            Set objCr = New AcLinksController
            objCr.deleteRecordAll

            For lngRow = 1 To UBound(varRecordset)
                Set objLink = New link

                With objLink
                    .title = varRecordset(lngRow, 1)
                    .url = varRecordset(lngRow, 3)
                End With

                objCr.createRecord objLink
            Next
    End Select

    MsgBox "完了しました。"
End Sub

'''
'出力先ディレクトリのパスを設定する。
'''
Private Sub lstType_AfterUpdate()
    With Me.lstType
        If .Value = "" Then
            Exit Sub
        End If

        If .Value <> "CSV" And .Value <> "Excel" And .Value <> "Access" Then
            MsgBox "ファイル出力形式が不正です。 Type を確認してください。"
            .Value = ""
            Exit Sub
        End If
    End With

    Dim objDialog As FileDialog
    Dim strPath As String

    Set objDialog = Application.FileDialog(msoFileDialogFolderPicker)

    With objDialog
        .InitialFileName = ActiveWorkbook.Path

        If .Show = -1 Then
            strPath = .SelectedItems(1)

            Select Case Me.lstType.Value
                Case "CSV"
                    strPath = strPath & "\Links.csv"
                Case "Excel"
                    strPath = strPath & "\Links.xlsx"
                Case "Access"
                    strPath = strPath & "\Links.accdb"
            End Select

            Me.txtPath = strPath
            Sheets("Config").Range("B4").Value = strPath
        End If
    End With
End Sub

'フォーム起動時の初期化処理
Private Sub UserForm_Activate()
    Me.txtPath.Value = Sheets("Config").Range("B4").Value

    With Me.lstType
        .Clear
        .AddItem "CSV"
        .AddItem "Excel"
        .AddItem "Access"
    End With
End Sub
