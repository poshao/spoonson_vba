Attribute VB_Name = "modTest"
'����: 2017/01/12
'����: Byron Gong
'����: ������ģ��

Option Explicit


'ʾ��
'����ΪTest_��ͷ ��������ʱ�� ��Test_201701121234()
'==============================================================================
'Public Sub Test_201701121234()
'  '����:
'  '����:
'  '����˵��:
'  '������:...
'
'End Sub
'==============================================================================
Sub Test_20170503()
    Dim db As clsDatabase, rs As ADODB.Recordset
    Dim strSQL As String
    Dim strFilename As String
    
    strFilename = Unit.CommonDialog.GetOpenFilename("all files|*.*")
    If strFilename = vbNullString Then Exit Sub
    
    Set db = Unit.CreateDatabase()
    db.ConnectEXCEL strFilename, True
    strSQL = "select * from [open so$] WHERE (((Trim([FSC]))<>'') AND ((Trim([AD]))=''));"
    Workbooks.Add
    Set rs = db.SelectCommand(strSQL)
    db.CopyToExcel rs, [A1], True
    db.Close_connection
    Unit.MessageBox.Info_ "OK"
End Sub

Sub Test_20170515_ExportCode()
    Dim strFolder As String
    strFolder = Unit.CommonDialog.GetFolderPath()
    If strFolder <> vbNullString Then
        Unit.ExportAllCode strFolder
        Unit.MessageBox.Info_ "OK"
    End If
End Sub
