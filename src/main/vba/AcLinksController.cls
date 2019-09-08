VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcLinksController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CrudRepository

Private objCn As ADODB.Connection

Private Sub Class_Initialize()
    '接続情報の取得
    Dim objCnInfo As New AdoCnInfo
    Dim strCnInfo As String

    With Sheets("Config")
        objCnInfo.provider = .Range("B3").Value
        objCnInfo.dataSource = .Range("B4").Value
        strCnInfo = objCnInfo.cnInfo
    End With

    Dim objFso As New FileSystemObject
    Set objCn = New ADODB.Connection

    If objFso.FileExists(objCnInfo.dataSource) = False Then
        'DB が存在しない場合は新規作成する。
        Dim objCatalog As New ADOX.Catalog
        Dim objCommand As New ADODB.Command

        objCatalog.Create strCnInfo
        objCn.Open strCnInfo

        With objCommand
            .ActiveConnection = objCn
            .CommandText = "CREATE TABLE Links(Title Text(100),Url Text(255));"
            .Execute
        End With
    Else
        '既に DB が存在する場合はそのまま開く。
        objCn.Open objCnInfo.cnInfo
    End If
End Sub

Private Sub Class_Terminate()
    If objCn.State = adStateOpen Then
        'DB 接続が開いている場合は切断する。
        objCn.Close
    End If
End Sub

'''
'Access DB へレコードを新規登録する。
'
'@param objModel 登録対象とする Link オブジェクトを指定する。
'''
Private Sub CrudRepository_createRecord(objModel As Object)
    Dim objLink As link
    Set objLink = objModel

    Dim objCommand As New ADODB.Command

    With objCommand
        .ActiveConnection = objCn
        .CommandText = "INSERT INTO Links (Title, Url) VALUES(?, ?);"
        .Parameters(0).Value = objLink.title
        .Parameters(1).Value = objLink.url
        .Execute
    End With
End Sub

Private Sub CrudRepository_deleteRecord(objModel As Object)

End Sub

'''
'Access DB からレコードを全件削除する。
'''
Private Sub CrudRepository_deleteRecordAll()
    Dim objCommand As New ADODB.Command

    With objCommand
        .ActiveConnection = objCn
        .CommandText = "DELETE FROM Links;"
        .Execute
    End With
End Sub

Private Function CrudRepository_findAll() As Range

End Function

Private Function CrudRepository_findByTitle(strTitle As String) As Range

End Function

Private Sub CrudRepository_updateRecord(objModel As Object)

End Sub
