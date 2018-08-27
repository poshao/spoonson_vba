VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInstallUpgradeHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'时间: 2018/08/27
'功能: 安装/升级框架工具
'作者: Byron Gong
'版本: 0.1.0

Option Explicit

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
    MsgBox "版本已更新至 " & strVersionNew
End Sub


'导出所有代码
Public Function ExportAll()
    Dim strFolder As String
    Dim comp As VBComponent
    
    strFolder = Util.CommonDialog.GetFolderPath()
'    Shell "cmd /c del /Q /F " & strFolder & "\*", vbNormalFocus
'    RmDir strFolder
'    MkDir strFolder
    If strFolder = vbNullString Then Exit Function

    '仅导出标准模块及类模块
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
        Case vbext_ct_MSForm '标准窗体
            comp.Export strFolder & "\" & comp.Name & ".frm"
        Case vbext_ct_StdModule '标准模块
            comp.Export strFolder & "\" & comp.Name & ".bas"
        Case vbext_ct_ClassModule '类模块
            comp.Export strFolder & "\" & comp.Name & ".cls"
        End Select
    Next
    ExportAll = True
    Util.MessageBox.Info_ "OK"
End Function

'安装框架
Public Function Install()
    Dim thisComp As VBComponent
    Dim destComp As VBComponent
    Dim destComps As VBComponents
    Dim strTempFilename As String
    
    Dim book As Workbook
    Dim strTip As String
    Dim i As Integer
    Dim strResult As String
    Dim arrBooknames
    
    ReDim arrBooknames(Workbooks.Count)
    For Each book In Workbooks
        i = i + 1
        arrBooknames(i) = book.Name
        strTip = strTip & i & "  " & book.Name & vbCrLf
    Next
    strResult = InputBox(strTip)
    
    If strResult = vbNullString Then
        Exit Function
    End If

    Set destComps = Workbooks(arrBooknames(Val(strResult))).VBProject.VBComponents
    
    For Each thisComp In ThisWorkbook.VBProject.VBComponents
        strTempFilename = vbNullString
        Select Case thisComp.Type
        Case vbext_ct_MSForm '标准窗体
            strTempFilename = ".frm"
        Case vbext_ct_StdModule '标准模块
            strTempFilename = ".bas"
        Case vbext_ct_ClassModule '类模块
            strTempFilename = ".cls"
        End Select
        
        If thisComp.Name = "modLog" Then
            strTempFilename = vbNullString
        End If
        
        If strTempFilename <> vbNullString Then
            On Error Resume Next
            Set destComp = destComps(thisComp.Name)
            If destComp Is Nothing = False Then
                destComps.Remove destComp
            End If
            On Error GoTo 0
            
            strTempFilename = VBA.Environ("temp") & "\spoonson_" & Format(Now, "yyyymmddhhmmss") & strTempFilename
            
            thisComp.Export strTempFilename
            destComps.Import strTempFilename
            Kill strTempFilename
        End If
    Next
    
    '添加引用
    Dim thisRef As Reference
    Dim destRefs As References
    Set destRefs = Workbooks(arrBooknames(Val(strResult))).VBProject.References
    On Error Resume Next
    For Each thisRef In ThisWorkbook.VBProject.References
        destRefs.AddFromFile thisRef.FullPath
    Next
    On Error GoTo 0
    MsgBox "OK"
End Function
