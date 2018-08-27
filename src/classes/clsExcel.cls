VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author: Byron Gong
'date: 10/26/2016 08:50
'desc: 实现Excel相关的基本操作

Option Explicit
'API声明
#If VBA7 And Win64 Then
  Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME As Long = &H40000 '(恢复大小)
Private Const WS_MINIMIZEBOX As Long = &H20000 '(最小化)
Private Const WS_MAXIMIZEBOX As Long = &H10000 '(最大化)

Private Const CLASSNAME As String = "clsExcel"

'内部程序
Dim m_app As Excel.Application 'app对象
Dim m_DefaultParentControl As Object

'属性
Public Property Get CurrentApp() As Excel.Application
  Set CurrentApp = m_app
End Property
Public Property Let CurrentApp(v As Excel.Application)
  Set m_app = v
End Property

Public Property Get DefaultParentControl() As Object
    Set DefaultParentControl = m_DefaultParentControl
End Property
Public Property Let DefaultParentControl(v As Object)
    If v Is Nothing Then
        Err.Raise 2001, Description:="默认工具栏设置无效(Excel)"
    Else
        Set m_DefaultParentControl = v
    End If
End Property

'通过窗体标题获取句柄
Public Function GetHwndByTitle(strTitle As String) As Long
    GetHwndByTitle = FindWindow("ThunderDFrame", strTitle)
End Function

'增加最小化按钮
Public Sub AddMinBox(strTitle As String)
    Dim hwnd As Long, lStyle As Long
    hwnd = GetHwndByTitle(strTitle)
    If hwnd = 0 Then Debug.Print "AddMinBox: can't find hwnd": Exit Sub
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX
    SetWindowLong hwnd, GWL_STYLE, lStyle
End Sub

'增加最大化按钮
Public Sub AddMaxBox(strTitle As String)
    Dim hwnd As Long, lStyle As Long
    hwnd = GetHwndByTitle(strTitle)
    If hwnd = 0 Then Debug.Print "AddMaxBox: can't find hwnd": Exit Sub
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MAXIMIZEBOX
    SetWindowLong hwnd, GWL_STYLE, lStyle
End Sub

'屏幕刷新及错误显示
Public Sub SetScreenUpdate(Optional b As Boolean = True)
    m_app.ScreenUpdating = b
    m_app.DisplayAlerts = b
End Sub

'工具栏======================================================================
'名称：ExistsCommandBar
'作用：检查是否存在工具栏
'参数：strName(string)工具栏名称
'返回值：(CommandBar)如果不存在则返回nothing 否则返回CommandBar对象
Public Function ExistsCommandBar(ByVal strName As String) As CommandBar
    On Error Resume Next
    Dim p_cb As CommandBar
    Set p_cb = m_app.CommandBars(strName)
    If Err.Number = 0 Then
        '找到相应对象
        Set ExistsCommandBar = p_cb
    ElseIf Err.Number = 5 Then
        '未找到相应对象
        Err.Clear
        Set ExistsCommandBar = Nothing
    Else
        '其他错误
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".ExistsCommandBar]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'名称：GetCommandBarByName
'作用：根据名称获取CommandBar对象或者创建一个
'参数：strName(string) 工具栏名称
'返回值：(CommandBar)如果没有找到 返回nothing 否则返回CommandBar对象
Public Function GetCommandBarByName(ByVal strName As String) As CommandBar
    Dim p_cb As CommandBar
    Set p_cb = Me.ExistsCommandBar(strName)
    If p_cb Is Nothing Then
        Set p_cb = m_app.CommandBars.Add(strName)
        p_cb.Visible = True
    End If
    Set GetCommandBarByName = p_cb
End Function

'名称：DeleteCommandbar
'作用：删除工具栏
'参数：strName(string) 工具栏名称
'返回值：(boolean)结果
Public Function DeleteCommandbar(ByVal strName As String) As Boolean
    Dim p_cb As CommandBar
    Set p_cb = Me.ExistsCommandBar(strName)
    If p_cb Is Nothing Then
        DeleteCommandbar = False
    Else
        p_cb.Delete
        DeleteCommandbar = True
    End If
End Function

'名称：GetMenuRoot
'作用：获取菜单工具栏对象
'返回值：(CommandBar)菜单栏对象
Public Function GetMenuRoot() As CommandBar
    Set GetMenuRoot = Me.ExistsCommandBar("Worksheet Menu Bar")
End Function

'名称：AddButton
'作用：增加一个按钮
'参数：
'       strName(string) 显示名称
'       [strAction](string)触发函数 默认vbnullstring
'       [iFaceId](integer)显示图标 默认空白
'       [parent](CommandBar)工具栏对象 默认为Nothing
'返回值：(CommandBarButton)按钮对象
Public Function AddButton(strName As String, Optional strAction As String = vbNullString, Optional iFaceId As Integer = 0, Optional parent = Nothing) As CommandBarButton
    Dim p_button As CommandBarButton
    If parent Is Nothing Then
        Set parent = Me.DefaultParentControl
    End If
    Set p_button = parent.Controls.Add(msoControlButton)
    p_button.Caption = strName
    p_button.OnAction = strAction
    p_button.FaceId = iFaceId
    p_button.style = msoButtonIconAndCaptionBelow
    p_button.Visible = True
    Set AddButton = p_button
End Function

'名称：AddComboBox
'作用：添加一个组合框
'参数：
'       strName(string) 组合框显示名称
'       [width](integer)组合框宽度
'       [parent](CommandBar)工具栏对象 默认为Nothing
'返回值：(CommandBarComboBox)组合框对象
Public Function AddComboBox(strName As String, Optional width As Integer = 50, Optional parent = Nothing) As CommandBarComboBox
    Dim p_combo As CommandBarComboBox
    If parent Is Nothing Then
        Set parent = Me.DefaultParentControl
    End If
    Set p_combo = parent.Controls.Add(msoControlComboBox)
    p_combo.Caption = strName
    p_combo.width = width
    p_combo.Visible = True
    p_combo.style = msoComboLabel
    Set AddComboBox = p_combo
End Function

'名称：AddPopup
'作用：增加一个弹出菜单
'参数：
'       strName(string) 显示名称
'       [strAction](string)触发函数 默认vbnullstring
'       [parent](CommandBar)工具栏对象 默认为Nothing
'返回值：(CommandBarPopup)菜单对象
Public Function AddPopup(strName As String, Optional strAction As String = vbNullString, Optional parent = Nothing) As CommandBarPopup
    Dim p_menu As CommandBarPopup
    If parent Is Nothing Then
        Set parent = Me.GetMenuRoot
    End If
    Set p_menu = parent.Controls.Add(msoControlPopup)
    p_menu.Caption = strName
    p_menu.OnAction = strAction
    p_menu.Visible = True
    Set AddPopup = p_menu
End Function

Public Function AddEdit() As CommandBarControl
    Debug.Print "本功能尚未实现:AddEdit"
End Function

Public Function AddDropdown() As CommandBarControl
    Debug.Print "本功能尚未实现:AddDropdown"
End Function

'名称：GetControl
'作用：获取控件对象
'参数：
'       strName(string)控件名称
'       [parent]父对象
'返回值：(CommandBarControl)控件对象 如果没有找到返回Nothing
Public Function GetControl(strName As String, Optional parent = Nothing) As CommandBarControl
    On Error Resume Next
    Dim p_control As CommandBarControl
    If parent Is Nothing Then
        Set parent = Me.DefaultParentControl
    End If
    Set p_control = parent.Controls(strName)
    If Err.Number = 0 Then
        Set GetControl = p_control
    ElseIf Err.Number = 5 Then
        '没有找到对象
        Err.Clear
        Set GetControl = Nothing
    Else
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".GetControl]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'名称：DeleteControl
'作用：根据名称删除控件
'参数：
'       strControlName(string)控件名称
'       [parent]控件容器
'返回值：(boolean)
Public Function DeleteControl(strControlName As String, Optional parent = Nothing) As Boolean
    On Error Resume Next
    Dim p_control As CommandBarControl
    Set p_control = Me.GetControl(strControlName, parent)
    If p_control Is Nothing Then
        DeleteControl = False
    Else
        p_control.Delete
        DeleteControl = True
    End If
End Function

'Excel表、簿相关======================================================================
'名称：ExistsSheet
'作用：检查是否存在指定名称的Sheet
'参数：
'       strSheetName(string)表名称
'       [book](object)默认为当前工作簿
'返回值：(boolean)返回True为有找到表
Public Function ExistsSheet(ByVal strSheetName As String, Optional book = Nothing) As Boolean
    On Error Resume Next
    Dim p_book As Workbook
    Dim p_sheet
    If book Is Nothing Then
        Set p_book = m_app.ActiveWorkbook
    Else
        Set p_book = book
    End If
    Set p_sheet = p_book.Sheets(strSheetName)
    If Err.Number = 0 Then
        ExistsSheet = True
    ElseIf Err.Number = 9 Then
        ExistsSheet = False
        Err.Clear
    Else
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".ExistsSheet]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'名称：ExistsWorkbook
'作用：是否工作簿打开
'参数：
'       strWorkbookName(string)工作簿名称 如"test.xlsx"
'       [app](object)应用对象
'返回值：(boolean)返回True为有找到工作簿，加载项也同样适用
Public Function ExistsWorkbook(ByVal strWorkbookName As String, Optional app = Nothing) As Boolean
    On Error Resume Next
    Dim p_app As Excel.Application
    Dim p_book
    
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    Set p_book = p_app.Workbooks(strWorkbookName)
    If Err.Number = 0 Then
        ExistsWorkbook = True
    ElseIf Err.Number = 9 Then
        Err.Clear
        ExistsWorkbook = False
    Else
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".ExistsWorkbook]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'格式======================================================================
'名称：SetBorderLine
'作用：设置边框线
'参数：
'       rng(Range) 单元格区域
'       [flag](xlBorderFlag)边框   若需清除原有边框线,添加 xlsNone
'       [style](XlLineStyle)边框样式
'       [weight](XlBorderWeight)边框线粗细
Public Sub SetBorderLine(rng As Range, Optional Flag As xlBorderFlag = xlsOutside Or xlsInside, Optional style As XlLineStyle = xlContinuous, Optional weight As XlBorderWeight = xlThin)
    Dim i  As Integer
    For i = 0 To 7
        If Flag And xlsNone = xlsNone Then
            rng.Borders(i + 5).LineStyle = xlNone
        End If
        If (Flag And 2 ^ i) > 0 Then
            With rng.Borders(i + 5)
                .LineStyle = style
                .weight = weight
            End With
        End If
    Next
End Sub

'加载项======================================================================
'名称：LoadAddin
'作用：装载加载项
'参数:
'       strPath(string)加载项路径
'       [Quiet](boolean)是否默认替换已经打开为文件
'       [app](object)应用对象
'返回值：(boolean)是否加载成功
Public Function LoadAddin(strPath As String, Optional Quiet As Boolean = False, Optional app = Nothing) As Boolean
    Dim p_app
    Dim p_File As New clsFile
    Dim p_strBook As String
    Dim p_bookAddin As Workbook
    
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    
    '检查文件存在
    If p_File.ExistsFile(strPath) = False Then
        '文件不存在
        LoadAddin = False
        Exit Function
    End If
    
    '检查是否同名文件打开
    p_strBook = p_File.GetFilename(strPath)
    If Me.ExistsWorkbook(p_strBook, p_app) = False Then
LOADADDIN_OPEN:
        Set p_bookAddin = p_app.Workbooks.Open(strPath, ReadOnly:=True)
        p_app.Run (p_strBook & "!modAutoload.Auto_Open")
        'p_bookAddin.RunAutoMacros xlAutoOpen
        LoadAddin = True
    Else
        If Quiet = True Then
            UnloadAddin p_strBook, p_app
            GoTo LOADADDIN_OPEN
        Else
            If MsgBox("文件名与已打开的冲突，是否关闭同名文件后继续打开？", vbYesNo, "提示") = vbYes Then
                UnloadAddin p_strBook, p_app
                GoTo LOADADDIN_OPEN
            End If
        End If
        '同名文件冲突
        LoadAddin = False
    End If
    
    Set p_File = Nothing
End Function

'名称：UnloadAddin
'作用：关闭加载项
'参数：
'       strName(string)加载项名称
'       [app](object)应用对象
'返回值：(boolean)执行结果
Public Function UnloadAddin(ByVal strName As String, Optional app = Nothing) As Boolean
    Dim p_app

    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    '检查文件是否加载
    If Me.ExistsWorkbook(strName, p_app) = True Then
        '执行卸载过程
        p_app.Run (strName & "!modAutoload.Auto_Close")
        'p_app.Workbooks(strName).RunAutoMacros xlAutoClose
        '关闭文件
        p_app.Workbooks(strName).Close False
        UnloadAddin = True
    Else
        UnloadAddin = False
    End If
End Function

'名称：LoadAddinWithValidate
'作用：带验证的加载
'参数：
'       strPath(string)加载项路径
'       strName(string)加载项名称
'       strVersion(string)加载项版本
'       [app](object)应用对象
'返回值：(boolean)执行结果
Public Function LoadAddinWithValidate(strPath As String, strName As String, strVersion As String, Optional app = Nothing) As Boolean
    Dim p_app
    Dim p_File As New clsFile
    Dim p_strBook As String
    Dim p_strAddinName As String, p_strAddinVersion As String
    
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    
    If LoadAddin(strPath, True, p_app) = True Then
        '验证版本
        p_strBook = p_File.GetFilename(strPath)
        p_strAddinName = Me.GetAddinName(p_strBook, p_app)
        p_strAddinVersion = Me.GetAddinVersion(p_strBook, p_app)
        If strName = p_strAddinName And strVersion = p_strAddinVersion Then
            LoadAddinWithValidate = True
        Else
            UnloadAddin p_strBook, p_app
            LoadAddinWithValidate = False
        End If
    Else
        '加载失败
        LoadAddinWithValidate = False
    End If
End Function

'名称：GetAddinVersion
'作用：获取加载项版本信息（版本信息通过调用 [modGlobal]中GetVersion函数获取）
'参数：
'       strAddinName(string)文件名称(包含后缀) 如"test.xla"
'       [app](object)应用对象
'返回值：(string)版本信息
Public Function GetAddinVersion(ByVal strAddinName As String, Optional app = Nothing) As String
    Dim p_app
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    If Me.ExistsWorkbook(strAddinName, p_app) = True Then
        GetAddinVersion = p_app.Run(strAddinName & "!modGlobal.GetVersion")
    Else
        GetAddinVersion = vbNullString
    End If
End Function

'名称：GetAddinName
'作用：获取加载项版本信息（版本信息通过调用 [modGlobal]中GetAppname函数获取）
'参数：
'       strAddinName(string)文件名称(包含后缀) 如"test.xla"
'       [app](object)应用对象
'返回值：(string)版本信息
Public Function GetAddinName(ByVal strAddinName As String, Optional app = Nothing) As String
    Dim p_app
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    If Me.ExistsWorkbook(strAddinName, p_app) = True Then
        GetAddinName = p_app.Run(strAddinName & "!modGlobal.GetAppname")
    Else
        GetAddinName = vbNullString
    End If
End Function

'名称：GetAddin
'作用：获取加载项对象
'参数：
'       strShortName(string)加载项名称 如 "test"
'       [app](object) 应用对象
'返回值：(Addin)加载项对象 如果没找到返回nothing
Public Function GetAddin(ByVal strShortName As String, Optional app = Nothing) As AddIn
    On Error Resume Next
    Dim p_app As Excel.Application
    Dim p_addin As AddIn
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    Set p_addin = p_app.AddIns(strShortName)
    If Err.Number = 0 Then
        Set GetAddin = p_addin
    ElseIf Err.Number = 9 Then
        Set GetAddin = Nothing
        Err.Clear
    Else
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".GetAddin]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'名称：UninstallAddin
'作用：卸载加载项
'参数：
'       strShortName(string)加载项名称 如"test"
'       [app](object)应用对象
'返回值：(boolean)结果
Public Function UninstallAddin(ByVal strShortName As String, Optional app = Nothing) As Boolean
    Dim p_addin As AddIn
    
    Set p_addin = GetAddin(strShortName, app)
    If p_addin Is Nothing Then
        '没有该加载项
        UninstallAddin = False
    Else
        p_addin.Installed = False
        Kill p_addin.FullName
        UninstallAddin = True
    End If
End Function

'名称：InstallAddin
'作用：安装加载项
'参数：
'       strPath(string)文件地址
'       [app](object)
'返回值：(boolean)安装结果
Public Function InstallAddin(strPath As String, Optional app = Nothing) As Boolean
    Dim p_File As New clsFile, p_addin As AddIn, p_app
    Dim strLocalPath As String, strLocalAddinPath As String
    
    strLocalPath = VBA.Environ$("appdata") & "\microsoft\addins\"
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    '检查文件是否存在
    If p_File.ExistsFile(strPath) = False Then
        InstallAddin = False
        Exit Function
    End If
    '卸载同名加载项
    UninstallAddin p_File.GetShortFilename(strPath)
    '复制文件到指定目录
    strLocalAddinPath = strLocalPath & p_File.GetFilename(strPath)
    FileCopy strPath, strLocalAddinPath
    Set p_addin = p_app.AddIns.Add(strLocalAddinPath)
    p_addin.Installed = True
    InstallAddin = True
End Function

Private Sub Class_Initialize()
    Set m_app = Application
End Sub

Private Sub Class_Terminate()
  Set m_app = Nothing
End Sub

