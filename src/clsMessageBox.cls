VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 02/06/2017
'desc: 消息提示

Option Explicit
Dim m_strDefaultTitle As String

'默认的标题名称
Public Property Get DefaultTitle() As String
    DefaultTitle = m_strDefaultTitle
End Property
Public Property Let DefaultTitle(ByVal strTitle As String)
    m_strDefaultTitle = strTitle
End Property

'名称：Info_
'作用：消息提示(提醒)
'参数：strMsg(string) 提示内容          [strTitle](string)标题显示
Public Sub Info_(ByVal strMsg As String, Optional ByVal strTitle As String = vbNullString)
    If strTitle = vbNullString Then strTitle = m_strDefaultTitle
    MsgBox strMsg, vbInformation + vbOKOnly, IIf(strTitle = vbNullString, "提示", strTitle)
End Sub

'名称：Warn_
'作用：消息提示(提醒)
'参数：strMsg(string) 提示内容          [strTitle](string)标题显示
Public Sub Warn_(ByVal strMsg As String, Optional ByVal strTitle As String = vbNullString)
    If strTitle = vbNullString Then strTitle = m_strDefaultTitle
    MsgBox strMsg, vbExclamation + vbOKOnly, IIf(strTitle = vbNullString, "警告", strTitle)
End Sub

'名称：Err_
'作用：消息提示(提醒)
'参数：strMsg(string) 提示内容          [strTitle](string)标题显示
Public Sub Err_(ByVal strMsg As String, Optional ByVal strTitle As String = vbNullString)
    If strTitle = vbNullString Then strTitle = m_strDefaultTitle
    MsgBox strMsg, vbCritical + vbOKOnly, IIf(strTitle = vbNullString, "错误", strTitle)
End Sub
