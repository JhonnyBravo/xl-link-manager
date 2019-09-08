VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoCnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private propProvider As String
Private propDataSource As String

'''
'@return provider ADO の Provider 設定を返す。
'''
Public Property Get provider() As String
    provider = propProvider
End Property

'''
'@param strProvider ADO の Provider 設定を指定する。
'''
Public Property Let provider(strProvider As String)
    propProvider = strProvider
End Property

'''
'@return dataSource 操作対象とする Access DB のパスを返す。
'''
Public Property Get dataSource() As String
    dataSource = propDataSource
End Property

'''
'@param strDataSource 操作対象とする Access DB のパスを指定する。
'''
Public Property Let dataSource(strDataSource As String)
    propDataSource = strDataSource
End Property

'''
'@return cnInfo ADO の接続情報を返す。
'''
Public Property Get cnInfo() As String
    Dim varCnInfo As Variant
    Dim strCnInfo As String

    varCnInfo = Array("Provider=" & provider, "Data Source=" & dataSource)
    strCnInfo = Join(varCnInfo, ";")

    cnInfo = strCnInfo
End Property
