VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''
'@return Range レコードを全件取得する。
'''
Public Function findAll() As Range

End Function

'''
'レコードを検索する。
'
'@param strTitle 検索対象とするタイトルを指定する。
'@return Range 検索条件に合致するセルの Range オブジェクトを返す。
'''
Public Function findByTitle(strTitle As String) As Range

End Function

'''
'レコードを新規登録する。
'
'@param objModel 新規登録するモデルオブジェクトを指定する。
'''
Public Sub createRecord(objModel As Object)

End Sub

'''
'既存レコードを更新する。
'
'@param objModel 更新情報を格納したモデルオブジェクトを指定する。
'''
Public Sub updateRecord(objModel As Object)

End Sub

'''
'レコードを削除する。
'
'@param objModel 削除情報を格納したモデルオブジェクトを指定する。
'''
Public Sub deleteRecord(objModel As Object)

End Sub

'''
'レコードを全件削除する。
'''
Public Sub deleteRecordAll()

End Sub
