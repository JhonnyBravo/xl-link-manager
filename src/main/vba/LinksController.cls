VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinksController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CrudRepository

Private objSheet As Worksheet
Private objSelf As CrudRepository

Private Sub Class_Initialize()
    Set objSheet = Sheets("Links")
    Set objSelf = Me
End Sub

'''
'リンク情報を新規登録する。
'
'@param objModel 登録対象とする Link オブジェクトを指定する。
'''
Private Sub CrudRepository_createRecord(objModel As Object)
    Dim objLink As link
    Dim objTc As New TitlesController

    Set objLink = objModel

    With objSheet.Range("B2")
        If .Cells(2, 1).Value = "" Then
            .Cells(2, 1).Value = objLink.title
            .Cells(2, 2).Formula = objLink.link
            .Cells(2, 3).Value = objLink.url
        ElseIf .Cells(2, 1).Value <> "" Then
            With .Cells(1, 1).End(xlDown)
                .Cells(2, 1).Value = objLink.title
                .Cells(2, 2).Formula = objLink.link
                .Cells(2, 3).Value = objLink.url
            End With
        End If
    End With

    objTc.addTitle objLink.title
End Sub

'''
'リンク情報を削除する。
'
'@param objModel 削除対象とする Link オブジェクトを指定する。
'''
Private Sub CrudRepository_deleteRecord(objModel As Object)
    Dim objLink As link
    Dim objRange As Range
    Dim objTc As TitlesController

    Set objLink = objModel
    Set objRange = objSelf.findByTitle(objLink.title)

    If objRange Is Nothing = False Then
        objRange.EntireRow.Delete

        Set objTc = New TitlesController
        objTc.deleteTitle
    End If
End Sub

'''
'リンク情報を全件削除する。
'''
Private Sub CrudRepository_deleteRecordAll()
    Dim strRowEnd As String
    Dim strColEnd As String

    strColEnd = "D3"

    With objSheet.Range("B2")
        If .Cells(2, 1).Value <> "" Then
            strRowEnd = .End(xlDown).Address
            objSheet.Range(strRowEnd, strColEnd).ClearContents
        End If
    End With
End Sub

'''
'@param strTitle 検索対象とするリンク情報のタイトルを指定する。
'@return Range 条件に合致するリンク情報を持つ行の Range オブジェクトを返す。
'''
Private Function CrudRepository_findByTitle(strTitle As String) As Range
    Dim lngRow As Long
    Dim boolResult As Boolean

    lngRow = 1
    boolResult = False

    With objSheet.Range("B3")
        While .Cells(lngRow, 1).Value <> ""
            If .Cells(lngRow, 1).Value = strTitle Then
                Set CrudRepository_findByTitle = .Cells(lngRow, 1)
                boolResult = True
                GoTo finally
            End If

            lngRow = lngRow + 1
        Wend
    End With

finally:
    If boolResult = False Then
        MsgBox strTitle & " が見つかりませんでした。"
    End If
End Function

'''
'リンク情報を更新する。
'
'@param objModel 更新対象とする Link オブジェクトを指定する。
'''
Private Sub CrudRepository_updateRecord(objModel As Object)
    Dim objLink As link
    Dim objRange As Range

    Set objLink = objModel
    Set objRange = objSelf.findByTitle(objLink.title)

    If objRange Is Nothing = False Then
        With objRange
            .Cells(1, 1).Value = objLink.title
            .Cells(1, 2).Formula = objLink.link
            .Cells(1, 3).Value = objLink.url
        End With
    End If
End Sub
