Attribute VB_Name = "modAutoload"
'日期: 2017/01/12
'作者: Byron Gong
'描述: 自动加载模板
'说明:
'     必须设定常量  APP_NAME  APP_VERSION

Option Explicit
'自行修改部分
'===============================================================================
'配置加载工具代码
Private Sub AutoLoad_LoadTools(xls As clsExcel)
'    Dim pp
'    xls.DefaultCommandBar = xls.GetCommandBarByName(APP_NAME & APP_VERSION)
'    Set pp = xls.AddPopup(APP_NAME, parent:=xls.GetMenuRoot)
'    xls.AddPopup "test", parent:=pp
    
    '这里添加工具调用代码
    'xls.AddButtonOrMenu "Test", "AutoLoad_Test",662
End Sub

'加载函数/方法
'Sub AutoLoad_Test()
'  MsgBox "Hello AutoLoad"
'End Sub

'===============================================================================
'以下部分不可更改
'自动加载
Sub auto_open()
    Dim xls As clsExcel
    Set xls = Unit.Excel
    xls.DeleteCommandbar APP_NAME & APP_VERSION
    xls.DeleteControl APP_NAME, xls.GetMenuRoot
    AutoLoad_LoadTools xls
    Set xls = Nothing
End Sub

'自动卸载
Sub auto_close()
    Dim xls As clsExcel
    Set xls = Unit.Excel
    xls.DeleteCommandbar APP_NAME & APP_VERSION
    xls.DeleteControl APP_NAME, xls.GetMenuRoot
    Set xls = Nothing
End Sub
