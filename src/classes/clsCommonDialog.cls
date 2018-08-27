VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 02/07/2017
'desc: �ļ�����

Option Explicit
Dim m_app As Excel.Application

'���ƣ�GetFolderPath
'���ã���ȡ�ļ���·��
'������
'       [strTitle](string) �Ի�����⣬Ĭ��Ϊ"��ѡ���ļ���"
'       [strInitPath](string)��ʼ·����Ĭ��Ϊvbnullstring
'����ֵ��(string)�ļ���·��
Public Function GetFolderPath(Optional ByVal strTitle As String = "��ѡ���ļ���", Optional ByVal strInitPath As String = vbNullString) As String
    Dim f As FileDialog
    Set f = m_app.FileDialog(msoFileDialogFolderPicker)
    f.Title = strTitle
    f.InitialFileName = strInitPath
    f.ButtonName = "Select"
    f.InitialView = msoFileDialogViewDetails
    If f.Show = -1 Then
        GetFolderPath = f.SelectedItems(1)
    End If
End Function

'���ƣ�GetOpenFilename
'���ã���ȡ���ļ���
'������
'       [strFilter](string) ��������ʹ��"|"���ָ�����������ʾ�����Զ����Ӻ�׺��Ĭ��ֵ vbnullstring
'       [strTitle](string)�Ի�����⣬Ĭ��ֵ"��ѡ���ļ�"
'       [strInitPath](string)��ʼ·����Ĭ��Ϊvbnullstring
'       [xMultiSelect](boolean)�Ƿ���Զ�ѡ����������ֵ����
'����ֵ����xMultiSelect����ΪTrue������ֵ�ַ������飬���򷵻�·���ַ�����ֱ�ӹر�ʱ����vbnullstring
Public Function GetOpenFilename(Optional ByVal strFilter As String = vbNullString, Optional ByVal strTitle As String = "��ѡ���ļ�", Optional ByVal strInitPath As String = vbNullString, Optional xMultiSelect As Boolean = False)
    Dim f As FileDialog
    Dim Filter
    Dim i As Integer
    Dim arrFiles() As String
    Set f = m_app.FileDialog(msoFileDialogOpen)
    f.Title = strTitle
    f.Filters.Clear
    If strFilter <> vbNullString Then
        Filter = Split(strFilter, "|")
        For i = LBound(Filter) To UBound(Filter) Step 2
            f.Filters.Add Filter(i), Filter(i + 1)
        Next
    End If
    f.InitialFileName = strInitPath
    f.InitialView = msoFileDialogViewDetails
    f.AllowMultiSelect = xMultiSelect
    If f.Show = -1 Then
        If xMultiSelect = True Then
            ReDim arrFiles(f.SelectedItems.Count - 1) As String
            For i = 0 To f.SelectedItems.Count - 1
                arrFiles(i) = f.SelectedItems(i + 1)
            Next
            GetOpenFilename = arrFiles
        Else
            GetOpenFilename = f.SelectedItems(1)
        End If
    Else
        GetOpenFilename = vbNullString
    End If
End Function

'���ƣ�GetSaveFilename
'���ã���ȡ�����ļ���
'������
'       [strFilter](string) ��������ʹ��"|"���ָ�����������ʾ�����Զ����Ӻ�׺��Ĭ��ֵ vbnullstring
'       [strTitle](string)�Ի�����⣬Ĭ��ֵ"��ѡ���ļ�"
'       [strInitPath](string)��ʼ·����Ĭ��Ϊvbnullstring
'����ֵ������·���ַ�����ֱ�ӹر�ʱ����vbnullstring
Public Function GetSaveFilename(Optional ByVal strFilter As String = vbNullString, Optional ByVal strTitle As String = "��ѡ���ļ�", Optional ByVal strInitPath As String = vbNullString) As String
    Dim f As FileDialog
    Dim Filter
    Dim i As Integer
    Set f = m_app.FileDialog(msoFileDialogSaveAs)
    f.Title = strTitle
    If strFilter <> vbNullString Then
        Filter = Split(strFilter, "|")
        For i = LBound(Filter) To UBound(Filter) Step 2
            f.Filters.Add Filter(i), Filter(i + 1)
        Next
    End If
    f.InitialFileName = strInitPath
    f.InitialView = msoFileDialogViewDetails
    f.AllowMultiSelect = False
    If f.Show = -1 Then
        GetOpenFilename = f.SelectedItems(1)
    Else
        GetOpenFilename = vbNullString
    End If
End Function

Private Sub Class_Initialize()
    Set m_app = Application
End Sub

Private Sub Class_Terminate()
    Set m_app = Nothing
End Sub