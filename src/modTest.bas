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

'导出代码
Sub Test001_ExportCode()
    Dim hp As New clsInstallUpgradeHelper
    hp.ExportAll
End Sub

Sub Test002_Install()
    Dim hp As New clsInstallUpgradeHelper
    hp.Install
End Sub

Sub Test003_UpgradeVersion()
    Dim hp As New clsInstallUpgradeHelper
    hp.UpgradeVersion
End Sub
