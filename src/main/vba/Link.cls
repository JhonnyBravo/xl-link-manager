VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private propTitle As String
Private propUrl As String
Private propLink As String

'''
'@return title リンクのタイトルを返す。
'''
Public Property Get title() As String
    title = propTitle
End Property

'''
'@param strTitle リンクのタイトルとして設定する文字列を指定する。
'''
Public Property Let title(strTitle As String)
    propTitle = strTitle
End Property

'''
'@return url リンクの URL を返す。
'''
Public Property Get url() As String
    url = propUrl
End Property

'''
'@param strUrl リンクの URL として設定する文字列を指定する。
'''
Public Property Let url(strUrl As String)
    propUrl = strUrl
End Property

'''
'@return HYPERLINK 関数を返す。
'''
Public Property Get link() As String
    link = "=HYPERLINK(""" & Me.url & """,""■"")"
End Property
