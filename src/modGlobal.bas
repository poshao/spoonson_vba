Attribute VB_Name = "modGlobal"
Option Explicit
'����: 2017/01/12
'����: Byron Gong
'����: ȫ�ֱ���,����
'˵��: ������������ APP_NAME  APP_VERSION

'ȫ����������
'===============================================================================


'ȫ�ֱ�������
'===============================================================================
Dim m_util As clsUtil

'ȫ�ֺ�������
'===============================================================================

'ȫ����������
'===============================================================================
'���߼���
Public Property Get Util() As clsUtil
    If m_util Is Nothing Then
        Set m_util = New clsUtil
    End If
    Set Util = m_util
End Property

'���ú���
'===============================================================================
'��ȡ�汾��
Public Function GetVersion() As String
    GetVersion = APP_VERSION
End Function
'��ȡ�汾����
Public Function GetAppname() As String
    GetAppname = APP_NAME
End Function
'��ȡ��ܰ汾��
Public Function GetFrameVersion() As String
    GetFrameVersion = FRAME_VERSION
End Function


