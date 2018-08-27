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
'desc: 文件操作

Option Explicit
Dim m_app As Excel.Application

'名称：GetFolderPath
'作用：获取文件夹路径
'参数：
'       [strTitle](string) 对话框标题，默认为"请选择文件夹"
'       [strInitPath](string)初始路径，默认为vbnullstring
'返回值：(string)文件夹路径
Public Function GetFolderPath(Optional ByVal strTitle As String = "请选择文件夹", Optional ByVal strInitPath As String = vbNullString) As String
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

'名称：GetOpenFilename
'作用：获取打开文件名
'参数：
'       [strFilter](string) 过滤器，使用"|"作分隔符，过滤显示名称自动添加后缀，默认值 vbnullstring
'       [strTitle](string)对话框标题，默认值"请选择文件"
'       [strInitPath](string)初始路径，默认为vbnullstring
'       [xMultiSelect](boolean)是否可以多选，决定返回值类型
'返回值：若xMultiSelect设置为True，返回值字符串数组，否则返回路径字符串，直接关闭时返回vbnullstring
Public Function GetOpenFilename(Optional ByVal strFilter As String = vbNullString, Optional ByVal strTitle As String = "请选择文件", Optional ByVal strInitPath As String = vbNullString, Optional xMultiSelect As Boolean = False)
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

'名称：GetSaveFilename
'作用：获取保存文件名
'参数：
'       [strFilter](string) 过滤器，使用"|"作分隔符，过滤显示名称自动添加后缀，默认值 vbnullstring
'       [strTitle](string)对话框标题，默认值"请选择文件"
'       [strInitPath](string)初始路径，默认为vbnullstring
'返回值：返回路径字符串，直接关闭时返回vbnullstring
Public Function GetSaveFilename(Optional ByVal strFilter As String = vbNullString, Optional ByVal strTitle As String = "请选择文件", Optional ByVal strInitPath As String = vbNullString) As String
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
