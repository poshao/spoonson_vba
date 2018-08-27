VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 12/05
'desc: 日志记录
'version:1.0.1
'
'初始文件夹路径为用户文件夹下的Logs文件夹
'默认文件名为 日期+日志名称+.log

Option Explicit
Dim m_strLogFolder As String
Dim m_strLogName As String
Dim m_strLogPath As String
Dim m_IsDebug As Boolean

'名称：RefreshPath
'作用：重新拼接路径地址
Private Sub RefreshPath()
    m_strLogPath = m_strLogFolder & VBA.Format$(Now(), "yyyymmdd") & m_strLogName & ".log"
End Sub

Public Property Get LogFolder() As String
    LogFolder = m_strLogFolder
End Property
Public Property Let LogFolder(ByVal v As String)
    v = VBA.Trim$(v)
    If v = vbNullString Then Err.Raise 1001, Description:="无效路径(LOG)"
    If VBA.Right$(v, 1) <> "\" Then v = v & "\"
    If Dir(v, vbDirectory) = vbNullString Then
        MkDir v
    End If
    m_strLogFolder = v
    RefreshPath
End Property

Public Property Get Filename() As String
    Filename = m_strLogName
End Property
Public Property Let Filename(ByVal v As String)
    m_strLogName = v
    RefreshPath
End Property

Public Property Get ShowDebug() As Boolean
  ShowDebug = m_IsDebug
End Property
Public Property Let ShowDebug(v As Boolean)
  m_IsDebug = v
End Property

Private Sub Class_Initialize()
    '初始化日志记录路径 用户桌面 Logs文件夹
    Me.LogFolder = VBA.Environ$("userprofile") & "\Logs\"
    Me.Filename = "LogSample"
End Sub

'debug
Public Sub Debug1(strMsg As String)
  If m_IsDebug Then
    WriteMsg "Debug", strMsg
  End If
End Sub

'Information
Public Sub Info(strMsg As String)
  WriteMsg "Info ", strMsg
End Sub
'Warnning
Public Sub Warn(strWarn As String)
  WriteMsg "Warn ", strWarn
End Sub
'Error
Public Sub Error(strErr As String)
  WriteMsg "Error", strErr
End Sub

'write message
Public Sub WriteMsg(strFlag As String, strMsg As String)
  WriteFile "[" & strFlag & "]" & VBA.Format$(Now(), "yyyy/mm/dd HH:mm:ss") & " => " & strMsg
End Sub

'write file
Public Sub WriteFile(txt As String)
  Dim f As Long
  f = FreeFile()
  Open m_strLogPath For Append As #f
  Lock #f
  Print #f, txt
'  Print #f, "[" & flag & "]" & Now() & " => " & strMsg
  Unlock #f
  Close #f
End Sub
