Attribute VB_Name = "modAutoload"
'����: 2017/01/12
'����: Byron Gong
'����: �Զ�����ģ��
'˵��:
'     �����趨����  APP_NAME  APP_VERSION

Option Explicit
'�����޸Ĳ���
'===============================================================================
'���ü��ع��ߴ���
Private Sub AutoLoad_LoadTools(xls As clsExcel)
'    Dim pp
'    xls.DefaultCommandBar = xls.GetCommandBarByName(APP_NAME & APP_VERSION)
'    Set pp = xls.AddPopup(APP_NAME, parent:=xls.GetMenuRoot)
'    xls.AddPopup "test", parent:=pp
    
    '������ӹ��ߵ��ô���
    'xls.AddButtonOrMenu "Test", "AutoLoad_Test",662
End Sub

'���غ���/����
'Sub AutoLoad_Test()
'  MsgBox "Hello AutoLoad"
'End Sub

'===============================================================================
'���²��ֲ��ɸ���
'�Զ�����
Sub auto_open()
    Dim xls As clsExcel
    Set xls = Unit.Excel
    xls.DeleteCommandbar APP_NAME & APP_VERSION
    xls.DeleteControl APP_NAME, xls.GetMenuRoot
    AutoLoad_LoadTools xls
    Set xls = Nothing
End Sub

'�Զ�ж��
Sub auto_close()
    Dim xls As clsExcel
    Set xls = Unit.Excel
    xls.DeleteCommandbar APP_NAME & APP_VERSION
    xls.DeleteControl APP_NAME, xls.GetMenuRoot
    Set xls = Nothing
End Sub
