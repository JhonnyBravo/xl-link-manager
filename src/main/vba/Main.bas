Attribute VB_Name = "Main"
Option Explicit

'''
'リンク編集用フォームを起動する。
'''
Public Sub openForm()
    LinkManager.Show vbModeless
End Sub

'''
'Yahoo 検索用フォームを起動する。
'''
Public Sub openYahooSearch()
    YahooSearch.Show vbModeless
End Sub

'''
'データエクスポート用フォームを起動する。
'''
Public Sub openExportLinks()
    ExportLinks.Show vbModeless
End Sub
