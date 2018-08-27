VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 02/06/2017
'desc: ϵͳ����

Option Explicit
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

'���ƣ�TickCount
'���ã���ȡʱ�Ӽ���
'����ֵ��(long)
Public Function TickCount() As Long
    TickCount = GetTickCount
End Function

'���ƣ�LoginName
'���ã���ȡϵͳ��¼��
'����ֵ��(string)��¼��
Public Function LoginName() As String
    LoginName = VBA.Environ$("username")
End Function

'���ƣ�DisplayName
'���ã���ȡ�û���
'����ֵ��(string)�û���
Public Function DisplayName() As String
    Dim objSysInfo As Object
    Set objSysInfo = CreateObject("ADSystemInfo")
    DisplayName = GetObject("LDAP://" & objSysInfo.UserName).DisplayName
    Set objSysInfo = Nothing
End Function

'���ƣ�ComputerName
'���ã���ȡ������
'����ֵ��(string)������
Public Function ComputerName() As String
    ComputerName = VBA.Environ$("computername")
End Function

'���ƣ�GetIPs
'���ã���ȡIP����
'����ֵ��(string())IP��ַ ����������С������ ��û�л�ȡ��IP
Public Function GetIPs() As String()
    Dim objSysInfo As Object, IPInfo As Object
    Dim p_strIPs As String
    
    Set objSysInfo = GetObject("winmgmts:{impersonationLevel=impersonate}//localhost").ExecQuery( _
                            "SELECT index, IPAddress FROM Win32_NetworkAdapterConfiguration")
    For Each IPInfo In objSysInfo
        If Not IsNull(IPInfo.IPAddress) Then
                p_strIPs = p_strIPs & "#" & IPInfo.IPAddress(0)
        End If
    Next
    Set objSysInfo = Nothing
    If Len(p_strIPs) > 0 Then
        GetIPs = Split(Mid(p_strIPs, 2), "#")
    Else
        GetIPs = Split("", "#")
    End If
End Function

Public Function GetPrinterList() As String()
    Dim i As Integer, n As Integer
    Dim ws As Object
    Dim strPrinterList() As String

    Set ws = CreateObject("wscript.network")
    n = ws.EnumPrinterConnections.Count
    ReDim strPrinterList(n \ 2 - 1) As String
    For i = 1 To n - 1 Step 2
         strPrinterList((i - 1) / 2) = ws.EnumPrinterConnections.Item(i) '��ӡ������
    Next
    Set ws = Nothing
    GetPrinterList = strPrinterList
End Function

'���ƣ�TempFolder
'���ã���ȡ��ʱ�ļ���
'����ֵ��(string)�ļ���·��
Public Function TempFolder() As String
    TempFolder = VBA.Environ$("TEMP")
End Function