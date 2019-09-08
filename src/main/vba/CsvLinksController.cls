VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvLinksController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CrudRepository

Private objFso As FileSystemObject
Private objStream As TextStream
Private strPath As String

Private Sub Class_Initialize()
    'ファイルパスの取得
    strPath = Sheets("Config").Range("B4").Value
    Set objFso = New FileSystemObject

    If objFso.FileExists(strPath) = False Then
        'ファイルが存在しない場合は新規作成してヘッダーを書き込む。
        Set objStream = objFso.CreateTextFile(strPath)

        With objStream
            .WriteLine "Title,Url"
            .Close
        End With

        Set objStream = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    If objStream Is Nothing = False Then
        objStream.Close
        Set objStream = Nothing
    End If
End Sub

'''
'CSV へレコードを新規登録する。
'
'@param objModel 登録対象とする Link オブジェクトを指定する。
'''
Private Sub CrudRepository_createRecord(objModel As Object)
    Dim objLink As link
    Dim varRecord As Variant
    Dim strRecord As String

    Set objLink = objModel

    If objStream Is Nothing Then
        Set objStream = objFso.OpenTextFile(strPath, ForAppending)
    End If

    varRecord = Array(objLink.title, objLink.url)
    strRecord = Join(varRecord, ",")

    objStream.WriteLine strRecord
End Sub

Private Sub CrudRepository_deleteRecord(objModel As Object)

End Sub

'''
'CSV ファイルからレコードを全件削除する。
'''
Private Sub CrudRepository_deleteRecordAll()
    With objFso
        If .FileExists(strPath) Then
            .DeleteFile strPath

            Set objStream = .CreateTextFile(strPath)

            With objStream
                .WriteLine "Title,Url"
                .Close
            End With

            Set objStream = Nothing
        End If
    End With
End Sub

Private Function CrudRepository_findAll() As Range

End Function

Private Function CrudRepository_findByTitle(strTitle As String) As Range

End Function

Private Sub CrudRepository_updateRecord(objModel As Object)

End Sub
