VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'''
'データエクスポート用フォームを起動する。
'''
Private Sub btnOpenExportLinks_Click()
    openExportLinks
End Sub

'''
'リンク編集用フォームを起動する。
'''
Private Sub btnOpenForm_Click()
    openForm
End Sub

'''
'Yahoo 検索用フォームを起動する。
'''
Private Sub btnOpenYahooSearch_Click()
    openYahooSearch
End Sub
