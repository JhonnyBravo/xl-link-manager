VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowserController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objIe As InternetExplorer
Private objDoc As HTMLDocument

'''
'Web ページの読込みが完了するまで待機する。
'''
Public Sub waitForLoading()
    With objIe
        While .Busy = True Or .readyState < READYSTATE_COMPLETE
            DoEvents
        Wend
    End With
End Sub

'''
'Internet Explorer(以下 IE と表記) を起動し、指定した Web ページを開く。
'
'@param strUrl IE に表示する Web ページの URL を指定する。
'@param boolVisible IE のウィンドウを表示するかどうかを真偽値で指定する。
'   * True: IE のウィンドウを表示する。
'   * False: IE のウィンドウを表示しない。(既定値)
'''
Public Sub openBrowser(strUrl As String, Optional boolVisible As Boolean = False)
    Set objIe = New InternetExplorer

    With objIe
        .Visible = boolVisible
        .navigate strUrl
    End With
End Sub

'''
'IE を終了する。
'''
Public Sub closeBrowser()
    If objIe Is Nothing = False Then
        objIe.Quit
        Set objIe = Nothing
    End If
End Sub

'''
'@return Document 現在表示されている Web ページから Document オブジェクトを取得する。
'''
Public Function getDocument() As HTMLDocument
    Set getDocument = objIe.document
End Function

Private Sub Class_Terminate()
    closeBrowser
End Sub
