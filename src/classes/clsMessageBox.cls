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
'desc: ��Ϣ��ʾ

Option Explicit
Dim m_strDefaultTitle As String

'Ĭ�ϵı�������
Public Property Get DefaultTitle() As String
    DefaultTitle = m_strDefaultTitle
End Property
Public Property Let DefaultTitle(ByVal strTitle As String)
    m_strDefaultTitle = strTitle
End Property

'���ƣ�Info_
'���ã���Ϣ��ʾ(����)
'������strMsg(string) ��ʾ����          [strTitle](string)������ʾ
Public Sub Info_(ByVal strMsg As String, Optional ByVal strTitle As String = vbNullString)
    If strTitle = vbNullString Then strTitle = m_strDefaultTitle
    MsgBox strMsg, vbInformation + vbOKOnly, IIf(strTitle = vbNullString, "��ʾ", strTitle)
End Sub

'���ƣ�Warn_
'���ã���Ϣ��ʾ(����)
'������strMsg(string) ��ʾ����          [strTitle](string)������ʾ
Public Sub Warn_(ByVal strMsg As String, Optional ByVal strTitle As String = vbNullString)
    If strTitle = vbNullString Then strTitle = m_strDefaultTitle
    MsgBox strMsg, vbExclamation + vbOKOnly, IIf(strTitle = vbNullString, "����", strTitle)
End Sub

'���ƣ�Err_
'���ã���Ϣ��ʾ(����)
'������strMsg(string) ��ʾ����          [strTitle](string)������ʾ
Public Sub Err_(ByVal strMsg As String, Optional ByVal strTitle As String = vbNullString)
    If strTitle = vbNullString Then strTitle = m_strDefaultTitle
    MsgBox strMsg, vbCritical + vbOKOnly, IIf(strTitle = vbNullString, "����", strTitle)
End Sub