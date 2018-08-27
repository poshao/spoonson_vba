VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRestful"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub doPost(url As String, strData As String)

    Dim http As New XMLHTTP
    
    With http
        .Open "POST", url
        .setRequestHeader "Content-Length", Len(strData)
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send strData
        Do Until .readyState = 4
            DoEvents
        Loop
        Debug.Print .responseText
    End With
    
    Set http = Nothing
End Sub
