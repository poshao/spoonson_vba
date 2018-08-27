VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUtil"
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

Public Property Get log() As clsLog
    If m_Log Is Nothing Then Set m_Log = New clsLog
    Set log = m_Log
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

Private Sub Class_Terminate()
    Set m_File = Nothing
    Set m_MsgBox = Nothing
    Set m_Log = Nothing
    Set m_xls = Nothing
    Set m_MD5 = Nothing
    Set m_CommonDialog = Nothing
End Sub
