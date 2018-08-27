VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 10/26/2016
'desc: 数据库操作

Option Explicit

Dim m_conn As ADODB.Connection

'属性
Public Property Get Connection() As ADODB.Connection
    Set Connection = m_conn
End Property

'数据库连接=============================================================
'如果成功连接,返回连接标识,否则Nothing--OK
Public Function Connect(ByVal strConnect As String) As Connection
    If m_conn.State = adStateOpen Then
        m_conn.Close
    End If
    m_conn.Open strConnect
    If m_conn.State = adStateOpen Then
        Set Connect = m_conn
    End If
End Function

'2000版本Access连接--OK
Public Function ConnectMDB(ByVal strMDBFile As String, Optional strUser As String = "admin", Optional strPasswd As String = "") As ADODB.Connection
    Dim strConnect As String
    strConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & strMDBFile & ";Uid=" & strUser & ";Pwd=" & strPasswd & ";"
    'strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strMDBFile & ";Jet OLEDB:Database Password=" & strPasswd & ";"
    Set ConnectMDB = Connect(strConnect)
End Function

'2000版本Access连接--OK
Public Function ConnectMDBByACE(ByVal strMDBFile As String, Optional strUser As String = "admin", Optional strPasswd As String = "") As ADODB.Connection
    Dim strConnect As String
    'strConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & strMDBFile & ";Uid=" & strUser & ";Pwd=" & strPasswd & ";"
    strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strMDBFile & ";Jet OLEDB:Database Password=" & strPasswd & ";"
    Set ConnectMDBByACE = Connect(strConnect)
End Function

'07之后版本Access连接--?
Public Function ConnectACCDB(ByVal strACCDBFile As String, Optional strPasswd As String = vbNullString) As Connection
    Dim strConnect As String
    strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strACCDBFile & ";Jet OLEDB:Database Password=" & strPasswd & ";"
    Set ConnectACCDB = Connect(strConnect)
End Function

'Paradox连接--?
Public Function ConnectParadox(ByVal strPath As String) As Connection
    Dim strConnect As String
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";Extended Properties=Paradox 5.x;"
    Set ConnectParadox = Connect(strConnect)
End Function

'MySQL连接--?
Public Function ConnectMySQL(ByVal strHost As String, ByVal strUser As String, ByVal strPasswd As String, ByVal strDbName As String, Optional ByVal strPort As String = "3306") As Connection
    Dim strConnect As String
    strConnect = "Server=" & strHost & ";Port=" & strPort & ";Database=" & strDbName & ";Uid=" & strUser & ";Pwd=" & strPort & ";"
    Set ConnectMySQL = Connect(strConnect)
End Function

'EXCEL连接(2007-)--?
Public Function ConnectEXCEL(ByVal strFile As String, ByVal bHasHeader As Boolean) As Connection
    Dim strConnect As String
    strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & _
                ";Extended Properties=""Excel 12.0 Xml;HDR=" & IIf(bHasHeader, "YES", "NO") & ";IMEX=1"";"
    Set ConnectEXCEL = Connect(strConnect)
End Function

'select查询(只读)
Public Function SelectCommand(sql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open sql, m_conn, adOpenKeyset, adLockReadOnly
    Set SelectCommand = rs
End Function

'检查是否有记录(主要用于登陆验证)
Public Function ExistsRecord(sql As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = SelectCommand(sql)
    If rs.RecordCount > 0 Then
        ExistsRecord = True
    Else
        ExistsRecord = False
    End If
    rs.Close
    Set rs = Nothing
End Function

'select查询(可更新)
Public Function SelectCommandUpdate(sql As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open sql, m_conn, adOpenKeyset, adLockOptimistic
    Set SelectCommandUpdate = rs
End Function

'insert/update/delete查询
Public Function QueryCommand(sql As String) As Long
    Dim iEffect As Long
    m_conn.Execute sql, iEffect
    QueryCommand = iEffect
End Function

'名称：CopyToExcel
'作用：复制到Excel表格
'参数：
'       rs(Recordset) 记录集
'       rng(Range) 单元格对象
'       [withTitle](boolean) 是否包含列标题
'返回值：无
Public Sub CopyToExcel(rs As ADODB.Recordset, rng As Excel.Range, Optional withTitle As Boolean = True)
    '生成表头
    Dim i As Integer
    If withTitle = True Then
        For i = 0 To rs.Fields.Count - 1
            rng.Offset(0, i).Value = rs.Fields(i).Name
        Next
        rng.Offset(1, 0).CopyFromRecordset rs.Clone
    Else
        rng.CopyFromRecordset rs.Clone
    End If
End Sub

'复制到数组
'名称：CopyToArray
'作用：复制到数组
'参数：
Public Function CopyToArray(rs As ADODB.Recordset, Optional hasHeader As Boolean = False)
    Dim arr()
    Dim i As Long, j As Integer
    rs.MoveFirst
    If hasHeader = True Then
        ReDim arr(rs.RecordCount, rs.Fields.Count - 1)
        For j = 0 To rs.Fields.Count - 1
            arr(0, j) = rs.Fields(j).Name
        Next
        For i = 0 To rs.RecordCount - 1
            For j = 0 To rs.Fields.Count - 1
                arr(i + 1, j) = rs(j)
            Next
            rs.MoveNext
        Next
    Else
        ReDim arr(rs.RecordCount - 1, rs.Fields.Count - 1)
        For i = 0 To rs.RecordCount - 1
            For j = 0 To rs.Fields.Count - 1
                arr(i, j) = rs(j)
            Next
            rs.MoveNext
        Next
    End If
    CopyToArray = arr
End Function

'文件转化为二进制流
Public Function GetBinFromFile(strFilename As String) As Variant
    Dim st As New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.LoadFromFile strFilename
    GetBinFromFile = st.Read
    st.Close
    Set st = Nothing
End Function

'二进制流转化为文件
Public Function SetBinToFile(bin, strFilename As String)
    Dim st As New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.Write bin
    st.SaveToFile strFilename
    st.Close
    Set st = Nothing
End Function

'关闭连接
Public Sub Close_connection()
    If Not (Me.Connection Is Nothing) Then
        Me.Connection.Close
    End If
End Sub

'导出到Excel表格
Public Sub Export_Excel(sheet, sql As String)
    If sheet Is Nothing Then
        Exit Sub
    End If
    Dim rs As ADODB.Recordset
    Dim i As Long
    Set rs = Me.SelectCommand(sql)
    If rs.BOF = True Then
        sheet.Cells(1, 1) = "没有记录"
    Else
        '生成标题
        For i = 0 To rs.Fields.Count - 1
            sheet.Cells(1, i + 1) = rs.Fields(i).Name
        Next
        sheet.Cells(2, 1).CopyFromRecordset rs.Clone
    End If
    rs.Close
End Sub

Public Function GetDbByName(strDbName As String) As clsDatabase
    Dim db As New clsDatabase

End Function

Private Sub Class_Initialize()
    Set m_conn = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    If m_conn.State = adStateOpen Then
        m_conn.Close
        Set m_conn = Nothing
    End If
End Sub

