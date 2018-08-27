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
'desc: 文件操作

Option Explicit

'名称：ExistsFile
'作用：检查文件是否存在
'参数：strFile(string) 文件路径
'返回值： (boolean) 存在文件返回True 否则返回False
Public Function ExistsFile(ByVal strFile As String) As Boolean
    If Dir(strFile) = vbNullString Then
        ExistsFile = False
    Else
        ExistsFile = True
    End If
End Function

'名称：GetFilename
'作用：获取包含后缀名的文件名
'参数：strFile(string) 文件名路径或文件名
'返回值：(string)包含后缀名的文件名称
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

'名称：GetShortFilename
'作用：获取不好含后缀名的文件名
'参数：strFile(string)文件名路径或文件名
'返回值：(string)不包含后缀名的文件名称
'依赖：GetFilename(fun)
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

'名称：GetFilePath
'作用：获取文件路径
'参数：strFile(string)文件名或文件路径
'返回值：(string)文件路径   若返回vbnullstring则该路径无效
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

'名称：GetExtensionName
'作用：获取文件后缀名
'参数：strFilename(string)文件名称
'返回值：(string)文件后缀名  若vbnullstring则无后缀
'依赖：GetFilename(fun)
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

'名称：DeleteFile
'作用：删除文件
'参数：strFile(string)文件路径
'返回值：(boolean)
'依赖：ExistsFile
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
