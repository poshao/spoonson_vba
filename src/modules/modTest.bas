Attribute VB_Name = "modTest"
'日期: 2017/01/12
'作者: Byron Gong
'描述: 测试用模板

Option Explicit


'示例
'命名为Test_开头 加上日期时间 如Test_201701121234()
'==============================================================================
'Public Sub Test_201701121234()
'  '日期:
'  '功能:
'  '测试说明:
'  '函数体:...
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
