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

'��������
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
