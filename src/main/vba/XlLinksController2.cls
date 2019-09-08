VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlLinksController2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CrudRepository

Private objExcel As Excel.Application
Private objBook As Workbook
Private objFso As New FileSystemObject

Private Sub Class_Initialize()
    Dim strPath As String

    strPath = Sheets("Config").Range("B4").Value
    Set objExcel = New Excel.Application
'    objExcel.Visible = True 'デバッグ用
    objExcel.Visible = False '本番用

    If objFso.FileExists(strPath) = False Then
        Set objBook = objExcel.Workbooks.Add

        With objBook.Sheets(1)
            .Range("A1").Value = "Title"
            .Range("B1").Value = "Url"
        End With

        objBook.SaveAs strPath
    Else
        Set objBook = objExcel.Workbooks.Open(strPath)
    End If
End Sub

Private Sub Class_Terminate()
    If objBook Is Nothing = False Then
        objBook.Close True
    End If

    If objExcel Is Nothing = False Then
        objExcel.Quit
    End If
End Sub

'''
'Excel ファイルへレコードを新規登録する。
'
'@param objModel 登録対象とする Link オブジェクトを指定する。
'''
Private Sub CrudRepository_createRecord(objModel As Object)
    Dim objLink As link
    Set objLink = objModel

    With objBook.Sheets(1).Range("A1")
        If .Cells(2, 1).Value = "" Then
            .Cells(2, 1).Value = objLink.title
            .Cells(2, 2).Value = objLink.url
        Else
            With .End(xlDown)
                .Cells(2, 1).Value = objLink.title
                .Cells(2, 2).Value = objLink.url
            End With
        End If
    End With
End Sub

Private Sub CrudRepository_deleteRecord(objModel As Object)

End Sub

'''
'Excel ファイルからレコードを全件削除する。
'''
Private Sub CrudRepository_deleteRecordAll()
    Dim strRowEnd As String
    Dim strColEnd As String

    strColEnd = "B2"

    With objBook.Sheets(1).Range("A1")
        If .Cells(2, 1).Value <> "" Then
            strRowEnd = .End(xlDown).Address
            objBook.Sheets(1).Range(strRowEnd, strColEnd).ClearContents
        End If
    End With
End Sub

Private Function CrudRepository_findAll() As Range

End Function

Private Function CrudRepository_findByTitle(strTitle As String) As Range

End Function

Private Sub CrudRepository_updateRecord(objModel As Object)

End Sub
