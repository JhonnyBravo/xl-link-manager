VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TitlesController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''
'lstTitle のリストへタイトルを新規登録する。
'
'@param strTitle 登録対象とするタイトルを指定する。
'''
Public Sub addTitle(strTitle As String)
    LinkManager.lstTitle.AddItem strTitle
End Sub

'''
'lstTitle のリストから現在選択されているタイトルを削除する。
'''
Public Sub deleteTitle()
    With LinkManager.lstTitle
        .RemoveItem .ListIndex
    End With

    With LinkManager
        .lstTitle.Value = ""
        .txtUrl.Value = ""
    End With
End Sub

'''
'lstTitle のリストを初期化する。
'''
Public Sub initTitles()
    With LinkManager.lstTitle
        If .ListCount > 0 Then
            .Clear
        End If
    End With

    Dim objSheet As Worksheet
    Dim objRange As Range

    Dim varTitles As Variant
    Dim lngRow As Long

    Dim strStartAddr As String
    Dim strEndAddr As String

    Set objSheet = Sheets("Links")
    strStartAddr = "B3"

    With objSheet.Range("B2")
        If .Cells(2, 1).Value <> "" Then
            strEndAddr = .Cells(1, 1).End(xlDown).Address
            Set objRange = objSheet.Range(strStartAddr, strEndAddr)

            If objRange.Address(False, False) = strStartAddr Then
                LinkManager.lstTitle.AddItem objRange.Value
                Exit Sub
            End If

            varTitles = objRange

            For lngRow = 1 To UBound(varTitles)
                LinkManager.lstTitle.AddItem varTitles(lngRow, 1)
            Next
        End If
    End With
End Sub
