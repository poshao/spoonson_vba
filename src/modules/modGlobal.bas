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
Dim m_unit As clsUnit

'ȫ�ֺ�������
'===============================================================================

'ȫ����������
'===============================================================================
'���߼���
Public Property Get Unit() As clsUnit
    If m_unit Is Nothing Then
        Set m_unit = New clsUnit
    End If
    Set Unit = m_unit
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


