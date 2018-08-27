Attribute VB_Name = "modGlobal"
Option Explicit
'日期: 2017/01/12
'作者: Byron Gong
'描述: 全局变量,函数
'说明: 必须声明常量 APP_NAME  APP_VERSION

'全局类型声明
'===============================================================================


'全局变量声明
'===============================================================================
Dim m_util As clsUtil

'全局函数声明
'===============================================================================

'全局属性声明
'===============================================================================
'工具集合
Public Property Get Util() As clsUtil
    If m_util Is Nothing Then
        Set m_util = New clsUtil
    End If
    Set Util = m_util
End Property

'内置函数
'===============================================================================
'获取版本号
Public Function GetVersion() As String
    GetVersion = APP_VERSION
End Function
'获取版本名称
Public Function GetAppname() As String
    GetAppname = APP_NAME
End Function
'获取框架版本号
Public Function GetFrameVersion() As String
    GetFrameVersion = FRAME_VERSION
End Function


