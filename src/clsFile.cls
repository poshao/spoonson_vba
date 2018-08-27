VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 02/06/2017
'desc: �ļ�����

Option Explicit

'���ƣ�ExistsFile
'���ã�����ļ��Ƿ����
'������strFile(string) �ļ�·��
'����ֵ�� (boolean) �����ļ�����True ���򷵻�False
Public Function ExistsFile(ByVal strFile As String) As Boolean
    If Dir(strFile) = vbNullString Then
        ExistsFile = False
    Else
        ExistsFile = True
    End If
End Function

'���ƣ�GetFilename
'���ã���ȡ������׺�����ļ���
'������strFile(string) �ļ���·�����ļ���
'����ֵ��(string)������׺�����ļ�����
Public Function GetFilename(ByVal strFile As String) As String
    Dim iPos As Integer
    strFile = VBA.Trim$(strFile)
    iPos = InStrRev(strFile, "\")
    If iPos < 1 Then
        GetFilename = strFile
    Else
        GetFilename = VBA.Left$(strFile, iPos - 1)
    End If
End Function

'���ƣ�GetShortFilename
'���ã���ȡ���ú���׺�����ļ���
'������strFile(string)�ļ���·�����ļ���
'����ֵ��(string)��������׺�����ļ�����
'������GetFilename(fun)
Public Function GetShortFilename(ByVal strFile As String) As String
    Dim p_strFilename As String, iPos As Integer
    p_strFilename = GetFilename(strFile)
    iPos = InStrRev(p_strFilename, ".")
    If iPos < 1 Then
        GetShortFilename = p_strFilename
    Else
        GetShortFilename = VBA.Left$(p_strFilename, iPos - 1)
    End If
End Function

'���ƣ�GetFilePath
'���ã���ȡ�ļ�·��
'������strFile(string)�ļ������ļ�·��
'����ֵ��(string)�ļ�·��   ������vbnullstring���·����Ч
Public Function GetFilePath(ByVal strFile As String) As String
    Dim iPos As Integer
    strFile = VBA.Trim$(strFile)
    If VBA.Right$(strFile, 1) = "\" Then
        GetFilePath = strFile
    Else
        iPos = InStrRev(strFile, "\")
        If iPos < 1 Then
            GetFilePath = vbNullString
        Else
            GetFilePath = VBA.Left$(strFile, iPos - 1)
        End If
    End If
End Function

'���ƣ�GetExtensionName
'���ã���ȡ�ļ���׺��
'������strFilename(string)�ļ�����
'����ֵ��(string)�ļ���׺��  ��vbnullstring���޺�׺
'������GetFilename(fun)
Public Function GetExtensionName(ByVal strFile As String) As String
    Dim p_strFilename As String, iPos As String
    strFile = VBA.Trim$(strFile)
    p_strFilename = GetFilename(strFile)
    iPos = InStrRev(p_strFilename, ".")
    If iPos < 1 Then
        GetExtensionName = p_strFilename
    Else
        GetExtensionName = VBA.Mid$(p_strFilename, iPos + 1)
    End If
End Function

'���ƣ�DeleteFile
'���ã�ɾ���ļ�
'������strFile(string)�ļ�·��
'����ֵ��(boolean)
'������ExistsFile
Public Function DeleteFile(ByVal strFile As String) As Boolean
    On Error Resume Next
    DeleteFile = False
    If ExistsFile(strFile) = True Then
        Kill strFile
        If Err.Number = 0 Then
            DeleteFile = True
        End If
    End If
End Function