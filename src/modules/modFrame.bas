Attribute VB_Name = "modFrame"
'日期: 2017/06/16
'作者: Byron Gong
'描述: 框架通用函数

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

'常用自定义函数
'名称：N2RMB
'作用：数值转化为中文大写
'参数：
'       n(number)数值
'返回值：(string)中文大写
Public Function N2RMB(n) As String
    Dim y As Integer, j As Double, f As Double
    Dim a As String, b As String, c As String
    y = Int(Round(100 * Abs(n)) / 100)
    j = Round(100 * Abs(n) + 0.00001) - y * 100
    f = (j / 10 - Int(j / 10)) * 10
    a = IIf(y < 1, "", Application.Text(y, "[DBNum2]") & "元")
    b = IIf(j > 9.5, Application.Text(Int(j / 10), "[DBNum2]") & "角", IIf(y < 1, "", IIf(f > 1, "零", "")))
    c = IIf(f < 1, "整", Application.Text(Round(f, 0), "[DBNum2]") & "分")
    N2RMB = IIf(Abs(n) < 0.005, "", IIf(n < 0, "负" & a & b & c, a & b & c))
End Function

'名称：GBK2Big5
'作用：简体转繁体
'参数：
'       strGBK(string) 简体字符串
'返回值：(string)繁体字符串
Function GBK2Big5(strGBK As String)
    Dim istrLen As Long, strBIG As String
    istrLen = lStrLen(strGBK)
    strBIG = Space(istrLen)
    LCMapString &H804, &H4000000, strGBK, istrLen, strBIG, istrLen
    GBK2Big5 = strBIG
End Function

'名称：Big52GBK
'作用：繁体转简体
'参数：
'       strBIG(string) 繁体字符串
'返回值：(string)简体字符串
Function Big52GBK(strBIG As String)
    Dim istrLen As Long, strGBK As String
    istrLen = lStrLen(strBIG)
    strGBK = Space(istrLen)
    LCMapString &H804, &H4000000, strBIG, istrLen, strGBK, istrLen
    Big52GBK = strGBK
End Function

'导入模块
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
