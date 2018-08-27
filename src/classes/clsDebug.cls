VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 11/20/2017 15:48
'desc: ต๗สิ
Option Explicit

'find Control Pos
Public Function FindFather(ctl)
    On Error GoTo EatErr
    If ctl.parent Is Nothing Then
        FindFather = ctl.Name
    Else
        FindFather = FindFather(ctl.parent) & "=>" & ctl.Name
    End If
    Exit Function
EatErr:
    If Err.Number = 438 Then
        FindFather = ""
    Else
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function
