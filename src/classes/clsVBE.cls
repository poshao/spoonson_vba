VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 03/04/2017
'desc: VBE操作类
Option Explicit

Dim m_VBE As VBIDE.VBE

Public Property Get XVBE()
    Set XVBE = m_VBE
End Property
Public Property Set XVBE(v As VBIDE.VBE)
    Set m_VBE = v
End Property

'增加右键弹出按钮
Public Function AddPopupMenu(ByVal strTitle As String, ByVal strRun As String, Optional lFaceID As Long = 0) As CommandBarButton
    Dim btnPopup As CommandBarButton
    Set btnPopup = m_VBE.CommandBars("Code Window").Controls.Add(msoControlButton)
    With btnPopup
        .Caption = strTitle
        .FaceId = lFaceID
        .OnAction = strRun
        .Visible = True
    End With
    Set AddPopupMenu = btnPopup
End Function

Private Sub Class_Initialize()
    If m_VBE Is Nothing Then
        Set m_VBE = Application.VBE
    End If
End Sub

Private Sub Class_Terminate()
    Set m_VBE = Nothing
End Sub
