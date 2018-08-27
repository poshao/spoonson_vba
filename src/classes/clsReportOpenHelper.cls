VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportOpenHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'时间: 2018/01/11
'功能: 打开并汇总报表到Excel格式
'作者: Byron Gong
'版本: 0.1.0

Option Explicit
Dim dig As clsCommonDialog

Private Sub Class_Initialize()
    Set dig = New clsCommonDialog
End Sub

'打开汇总OLXR报表
Public Function OpenOLXRs(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath, i As Long
    Dim strTemp As String
    Dim rngDelete As Range
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("OLXR文件|*.txt|所有文件|*.*", "请选择OLXR文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    '汇总OLXR文件
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow:=10, DataType:=xlFixedWidth, _
                        FieldInfo:=Array(Array(0, 1), Array(11, 1), Array(21, 1), Array(30, 1), Array(41, 1), _
                        Array(52, 1), Array(63, 1), Array(74, 1), Array(85, 1), Array(96, 1), Array(136, 1), _
                        Array(148, 1)), TrailingMinusNumbers:=True
        Range("A1:K" & Range("B1048576").End(xlUp).Row).Copy sheet.Range("B1048576").End(xlUp).Offset(1, -1)
        ActiveWorkbook.Close False
    Next
    
    '剔除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = Range("B1048576").End(xlUp).Row To 2 Step -1
        strTemp = Cells(i, "B").Formula
        If Len(strTemp) <> 8 Or VBA.Left(strTemp, 1) = "-" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    OpenOLXRs = True
End Function

'打开汇总S95报表
Public Function OpenS95s(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("S95文件|*.txt|所有文件|*.*", "请选择S95文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow:=9, DataType:=xlFixedWidth, FieldInfo:=Array( _
                    Array(0, 1), Array(7, 1), Array(18, 1), Array(25, 2), Array(51, 1), Array(55, 1), Array(68, _
                    1), Array(77, 2), Array(103, 1), Array(107, 1), Array(118, 1), Array(122, 1), Array(131, 1) _
                    , Array(138, 1), Array(146, 1), Array(155, 1), Array(164, 1), Array(173, 1), Array(182, 1), _
                    Array(203, 1), Array(219, 1), Array(228, 1)), TrailingMinusNumbers:=True
                    
        Range("A1:U" & Range("O1048576").End(xlUp).Row).Copy sheet.Range("O1048576").End(xlUp).Offset(1, -14)
        ActiveWorkbook.Close False
    Next
    
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("O1048576").End(xlUp).Row
        strTemp = Cells(i, "O").Formula
        If VBA.Len(strTemp) < 1 Or strTemp = "CREATEBY" Or VBA.Left(strTemp, 1) = "-" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    OpenS95s = True
End Function

'打开汇总S96报表
Public Function OpenS96s(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("S96文件|*.txt|所有文件|*.*", "请选择S96文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow:=6 _
            , DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, Comma:=True, _
            Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
            3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
            , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
            Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 2), Array(21, 1), Array(22, 1), Array( _
            23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
            Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array( _
            36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1)), TrailingMinusNumbers _
            :=True
        Range("A1:AN" & Range("A1048576").End(xlUp).Row).Copy sheet.Range("A1048576").End(xlUp).Offset(1, 0)
        ActiveWorkbook.Close False
    Next
        
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("A1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "A").Formula)
        If VBA.Len(strTemp) < 1 Or strTemp = "SHIPDOC" Or VBA.Left(strTemp, 1) = "*" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    OpenS96s = True
End Function

'打开472D报表
Public Function Open472Ds(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("472D文件|*.txt|所有文件|*.*", "请选择472D文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                        xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
                        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
                        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
                        , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), _
                        Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array( _
                        28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), _
                        Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array( _
                        41, 1), Array(42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1)), _
                        TrailingMinusNumbers:=True
        Range("A1:AS" & Range("A1048576").End(xlUp).Row).Copy sheet.Range("A1048576").End(xlUp).Offset(1, 0)
        ActiveWorkbook.Close False
    Next
    
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("A1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "A").Formula)
        If strTemp = "Retail#" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    Open472Ds = True
End Function

'打开209D报表
Public Function Open209Ds(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("209D文件|*.txt|所有文件|*.*", "请选择209D文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow _
                        :=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                        ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, _
                        Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
                        3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
                        , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
                        Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array( _
                        23, 1)), TrailingMinusNumbers:=True
        Range("A1:W" & Range("A1048576").End(xlUp).Row).Copy sheet.Range("A1048576").End(xlUp).Offset(1, 0)
        ActiveWorkbook.Close False
    Next
    
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("A1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "A").Formula)
        If strTemp = "TE#" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    Open209Ds = True
End Function

'打开030D报表
Public Function Open030Ds(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("209D文件|*.txt|所有文件|*.*", "请选择209D文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, _
                StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, _
                Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
                3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
                , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1)), _
                TrailingMinusNumbers:=True
        Range("A1:O" & Range("B1048576").End(xlUp).Row).Copy sheet.Range("B1048576").End(xlUp).Offset(1, -1)
        ActiveWorkbook.Close False
    Next
    
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("B1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "A").Formula)
        If strTemp = "BU" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    Open030Ds = True
End Function

'打开013D报表
Public Function Open013Ds(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("212D文件|*.txt|所有文件|*.*", "请选择212D文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
                xlNone, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
                Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
                Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
                , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), _
                Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array( _
                28, 1), Array(29, 1)), TrailingMinusNumbers:=True
        Range("A1:AC" & Range("A1048576").End(xlUp).Row).Copy sheet.Range("A1048576").End(xlUp).Offset(1, 0)
        ActiveWorkbook.Close False
    Next
    
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("A1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "A").Formula)
        If strTemp = "PTC" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    Open013Ds = True
End Function

'打开212D报表
Public Function Open212Ds(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("212D文件|*.txt|所有文件|*.*", "请选择212D文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:=936, StartRow _
            :=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, _
            Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
            3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
            , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
            Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array( _
            23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
            Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array( _
            36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), _
            Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array( _
            49, 1), Array(50, 1), Array(51, 1), Array(52, 1)), TrailingMinusNumbers:=True
        Range("A1:AZ" & Range("B1048576").End(xlUp).Row).Copy sheet.Range("B1048576").End(xlUp).Offset(1, -1)
        ActiveWorkbook.Close False
    Next
    
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("B1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "B").Formula)
        If strTemp = "BO NAME" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    Open212Ds = True
End Function

'打开S93报表
Public Function OpenS93s(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("S93文件|*.xls;*.xlsx|所有文件|*.*", "请选择S93文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, Origin:= _
            936, StartRow:=5, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, Comma:=True, _
            Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array( _
            3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
            , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
            Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1)), _
            TrailingMinusNumbers:=True
        Range("A1:U" & Range("B1048576").End(xlUp).Row).Copy sheet.Range("B1048576").End(xlUp).Offset(1, -1)
        ActiveWorkbook.Close False
    Next
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("B1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "B").Formula)
        If strTemp = "BillTo#" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    OpenS93s = True
End Function

'打开RZV报表
Public Function OpenRZVs(sheet As Worksheet, Optional strPaths) As Boolean
    Dim strPath
    Dim strTemp As String
    Dim rngDelete As Range
    Dim i As Long
    
    If VBA.IsMissing(strPaths) Then
        strPaths = dig.GetOpenFilename("RZV文件|*.txt|所有文件|*.*", "请选择RZV文件", "", True)
        If TypeName(strPaths) = "String" Then Exit Function
    End If
    
    Application.ScreenUpdating = False
    For Each strPath In strPaths
        Workbooks.OpenText Filename:=strPath, _
                Origin:=936, StartRow:=2, DataType:=xlDelimited, TextQualifier:= _
                xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
                , Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
                Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1)), _
                TrailingMinusNumbers:=True
        
        Range("A1:M" & Range("B1048576").End(xlUp).Row).Copy sheet.Range("B1048576").End(xlUp).Offset(1, -1)
        ActiveWorkbook.Close False
    Next
    '删除无效数据
    sheet.Activate
    Rows(1).Delete
    
    For i = 2 To Range("B1048576").End(xlUp).Row
        strTemp = VBA.Trim(Cells(i, "B").Formula)
        If strTemp = "BO" Then
            If rngDelete Is Nothing Then
                Set rngDelete = Cells(i, "A")
            Else
                Set rngDelete = Union(Cells(i, "A"), rngDelete)
            End If
        End If
    Next
    If rngDelete Is Nothing = False Then rngDelete.EntireRow.Delete
    Application.ScreenUpdating = True
    
    OpenRZVs = True
End Function

'打开OLU报表
Public Function OpenOLUs(sheet As Worksheet, Optional strPaths) As Boolean
    
End Function

'打开RZT报表
Public Function OpenRZTs(sheet As Worksheet, Optional strPaths) As Boolean

End Function
