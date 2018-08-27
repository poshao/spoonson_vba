Attribute VB_Name = "modFrame"
'����: 2017/06/16
'����: Byron Gong
'����: ���ͨ�ú���

Option Explicit
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
    Private Declare PtrSafe Function lStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
#Else
    Private Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
    Private Declare Function lStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
#End If

Public Enum xlBorderFlag
    xlsDiagonalDown = 1
    xlsDiagonalUp = 2
    xlsEdgeLeft = 4
    xlsEdgeTop = 8
    xlsEdgeBottom = 16
    xlsEdgeRight = 32
    xlsInsideVertical = 64
    xlsInsideHorizontal = 128
    xlsOutside = 60
    xlsInside = 192
    xlsAll = 252
    xlsNone = 256
End Enum

'�����Զ��庯��
'���ƣ�N2RMB
'���ã���ֵת��Ϊ���Ĵ�д
'������
'       n(number)��ֵ
'����ֵ��(string)���Ĵ�д
Public Function N2RMB(n) As String
    Dim y As Integer, j As Double, f As Double
    Dim a As String, b As String, c As String
    y = Int(Round(100 * Abs(n)) / 100)
    j = Round(100 * Abs(n) + 0.00001) - y * 100
    f = (j / 10 - Int(j / 10)) * 10
    a = IIf(y < 1, "", Application.Text(y, "[DBNum2]") & "Ԫ")
    b = IIf(j > 9.5, Application.Text(Int(j / 10), "[DBNum2]") & "��", IIf(y < 1, "", IIf(f > 1, "��", "")))
    c = IIf(f < 1, "��", Application.Text(Round(f, 0), "[DBNum2]") & "��")
    N2RMB = IIf(Abs(n) < 0.005, "", IIf(n < 0, "��" & a & b & c, a & b & c))
End Function

'���ƣ�GBK2Big5
'���ã�����ת����
'������
'       strGBK(string) �����ַ���
'����ֵ��(string)�����ַ���
Function GBK2Big5(strGBK As String)
    Dim istrLen As Long, strBIG As String
    istrLen = lStrLen(strGBK)
    strBIG = Space(istrLen)
    LCMapString &H804, &H4000000, strGBK, istrLen, strBIG, istrLen
    GBK2Big5 = strBIG
End Function

'���ƣ�Big52GBK
'���ã�����ת����
'������
'       strBIG(string) �����ַ���
'����ֵ��(string)�����ַ���
Function Big52GBK(strBIG As String)
    Dim istrLen As Long, strGBK As String
    istrLen = lStrLen(strBIG)
    strGBK = Space(istrLen)
    LCMapString &H804, &H4000000, strBIG, istrLen, strGBK, istrLen
    Big52GBK = strGBK
End Function

'����ģ��
Function ImportCodes(Optional strFolder As String = vbNullString)
    Dim fso As New FileSystemObject
    Dim f As File
    If strFolder = vbNullString Then strFolder = "C:\Users\0115289\git\VBA_COMMON"
    For Each f In fso.GetFolder(strFolder).Files
        If f.Type = "CLS File" Or f.Type = "BAS File" Then
            ThisWorkbook.VBProject.VBComponents.Import f.Path
        End If
    Next
End Function
