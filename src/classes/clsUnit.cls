VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 02/06/2017
'desc: 集合库

Const DB_Entry_List As String = "Driver={Microsoft Access Driver (*.mdb)};Dbq=dbpath;Uid=admin;Pwd=;" '数据库连接代码

Option Explicit

Dim m_File As clsFile
Dim m_MsgBox As clsMessageBox
Dim m_Log As clsLog
Dim m_xls As clsExcel
Dim m_MD5 As clsMD5
Dim m_Environ As clsEnvironment
Dim m_CommonDialog As clsCommonDialog

Public Property Get File() As clsFile
    If m_File Is Nothing Then Set m_File = New clsFile
    Set File = m_File
End Property

Public Property Get MessageBox() As clsMessageBox
    If m_MsgBox Is Nothing Then Set m_MsgBox = New clsMessageBox
    Set MessageBox = m_MsgBox
End Property

Public Property Get LOG() As clsLog
    If m_Log Is Nothing Then Set m_Log = New clsLog
    Set LOG = m_Log
End Property

Public Property Get Excel() As clsExcel
    If m_xls Is Nothing Then Set m_xls = New clsExcel
    Set Excel = m_xls
End Property

Public Property Get Md5() As clsMD5
    If m_MD5 Is Nothing Then Set m_MD5 = New clsMD5
    Set Md5 = m_MD5
End Property

Public Property Get Environment() As clsEnvironment
    If m_Environ Is Nothing Then Set m_Environ = New clsEnvironment
    Set Environment = m_Environ
End Property

Public Property Get CommonDialog() As clsCommonDialog
    If m_CommonDialog Is Nothing Then Set m_CommonDialog = New clsCommonDialog
    Set CommonDialog = m_CommonDialog
End Property

'创建数据库操作类
Public Function CreateDatabase() As clsDatabase
    Set CreateDatabase = New clsDatabase
End Function

'名称：GetDbByName
'作用：获取数据库连接
'参数：strName(string) 数据库名称
'返回值： (clsDatabase) 成功打开返回数据库连接,否则 nothing
Public Function GetDbByName(strName As String) As clsDatabase
    Dim db As New clsDatabase, rs As ADODB.Recordset, strSQL As String
    Dim db_out As clsDatabase
    
    db.Connect DB_Entry_List
    strSQL = "select path,passwd from dblist where dbname='" & strName & "'"
    Set rs = db.SelectCommand(strSQL)
    If rs.RecordCount > 0 Then
        Set db_out = New clsDatabase
        db_out.ConnectMDB rs(0), strPasswd:=rs(1)
        Set GetDbByName = db_out
    Else
        MsgBox "数据库【" & strName & "】不存在"
    End If
    db.Close_connection
End Function

'导出框架代码
Public Sub ExportAllCode(Optional strFolder As String = vbNullString)
    Dim Comp As VBComponent
    Dim strFilename As String
    
    If strFolder = vbNullString Then
        strFolder = VBA.Environ$("userprofile") & "\git\VBA_COMMON"
    End If
    '仅导出标准模块及类模块
    For Each Comp In ThisWorkbook.VBProject.VBComponents
        Select Case Comp.Type
        Case vbext_ct_StdModule '标准模块
            strFilename = strFolder & "\" & Comp.Name & ".bas"
            Comp.Export strFilename
            Debug.Print "Export " & Comp.Name & " to " & strFilename
        Case vbext_ct_ClassModule '类模块
            strFilename = strFolder & "\" & Comp.Name & ".cls"
            Comp.Export strFilename
            Debug.Print "Export " & Comp.Name & " to " & strFilename
        End Select
    Next
End Sub

'版本号升级
'修改末位版本号+1
Public Sub UpgradeVersion()
    Dim com As VBComponent
    Dim i As Long, strLine As String
    Dim strVersion As String, strVersionNew As String
    Dim iStart As Integer, iEnd As Integer
    Dim iPointPos As Integer
    
    Set com = ThisWorkbook.VBProject.VBComponents("modDefine")
    For i = 1 To com.CodeModule.CountOfLines
        strLine = com.CodeModule.Lines(i, 1)
        If InStr(strLine, "APP_VERSION") > 0 Then
            iStart = InStr(strLine, """") + 1
            iEnd = InStrRev(strLine, """")
            strVersion = Mid(strLine, iStart, iEnd - iStart)
            iPointPos = InStrRev(strVersion, ".")
            If iPointPos < 1 Then
                Err.Raise 3001, Description:="无效版本号 : " & strVersion
            End If
            strVersionNew = Left(strVersion, iPointPos) & Val(Mid(strVersion, iPointPos + 1)) + 1
            com.CodeModule.ReplaceLine i, Replace(strLine, strVersion, strVersionNew, Count:=1)
            Exit For
        End If
    Next
End Sub

Private Sub Class_Terminate()
    Set m_File = Nothing
    Set m_MsgBox = Nothing
    Set m_Log = Nothing
    Set m_xls = Nothing
    Set m_MD5 = Nothing
    Set m_CommonDialog = Nothing
End Sub
