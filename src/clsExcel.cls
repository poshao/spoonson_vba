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
'desc: ʵ��Excel��صĻ�������

Option Explicit
'API����
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
Private Const WS_THICKFRAME As Long = &H40000 '(�ָ���С)
Private Const WS_MINIMIZEBOX As Long = &H20000 '(��С��)
Private Const WS_MAXIMIZEBOX As Long = &H10000 '(���)

Private Const CLASSNAME As String = "clsExcel"

'�ڲ�����
Dim m_app As Excel.Application 'app����
Dim m_DefaultParentControl As Object

'����
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
        Err.Raise 2001, Description:="Ĭ�Ϲ�����������Ч(Excel)"
    Else
        Set m_DefaultParentControl = v
    End If
End Property

'ͨ����������ȡ���
Public Function GetHwndByTitle(strTitle As String) As Long
    GetHwndByTitle = FindWindow("ThunderDFrame", strTitle)
End Function

'������С����ť
Public Sub AddMinBox(strTitle As String)
    Dim hwnd As Long, lStyle As Long
    hwnd = GetHwndByTitle(strTitle)
    If hwnd = 0 Then Debug.Print "AddMinBox: can't find hwnd": Exit Sub
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX
    SetWindowLong hwnd, GWL_STYLE, lStyle
End Sub

'������󻯰�ť
Public Sub AddMaxBox(strTitle As String)
    Dim hwnd As Long, lStyle As Long
    hwnd = GetHwndByTitle(strTitle)
    If hwnd = 0 Then Debug.Print "AddMaxBox: can't find hwnd": Exit Sub
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MAXIMIZEBOX
    SetWindowLong hwnd, GWL_STYLE, lStyle
End Sub

'��Ļˢ�¼�������ʾ
Public Sub SetScreenUpdate(Optional b As Boolean = True)
    m_app.ScreenUpdating = b
    m_app.DisplayAlerts = b
End Sub

'������======================================================================
'���ƣ�ExistsCommandBar
'���ã�����Ƿ���ڹ�����
'������strName(string)����������
'����ֵ��(CommandBar)����������򷵻�nothing ���򷵻�CommandBar����
Public Function ExistsCommandBar(ByVal strName As String) As CommandBar
    On Error Resume Next
    Dim p_cb As CommandBar
    Set p_cb = m_app.CommandBars(strName)
    If Err.Number = 0 Then
        '�ҵ���Ӧ����
        Set ExistsCommandBar = p_cb
    ElseIf Err.Number = 5 Then
        'δ�ҵ���Ӧ����
        Err.Clear
        Set ExistsCommandBar = Nothing
    Else
        '��������
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".ExistsCommandBar]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'���ƣ�GetCommandBarByName
'���ã��������ƻ�ȡCommandBar������ߴ���һ��
'������strName(string) ����������
'����ֵ��(CommandBar)���û���ҵ� ����nothing ���򷵻�CommandBar����
Public Function GetCommandBarByName(ByVal strName As String) As CommandBar
    Dim p_cb As CommandBar
    Set p_cb = Me.ExistsCommandBar(strName)
    If p_cb Is Nothing Then
        Set p_cb = m_app.CommandBars.Add(strName)
        p_cb.Visible = True
    End If
    Set GetCommandBarByName = p_cb
End Function

'���ƣ�DeleteCommandbar
'���ã�ɾ��������
'������strName(string) ����������
'����ֵ��(boolean)���
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

'���ƣ�GetMenuRoot
'���ã���ȡ�˵�����������
'����ֵ��(CommandBar)�˵�������
Public Function GetMenuRoot() As CommandBar
    Set GetMenuRoot = Me.ExistsCommandBar("Worksheet Menu Bar")
End Function

'���ƣ�AddButton
'���ã�����һ����ť
'������
'       strName(string) ��ʾ����
'       [strAction](string)�������� Ĭ��vbnullstring
'       [iFaceId](integer)��ʾͼ�� Ĭ�Ͽհ�
'       [parent](CommandBar)���������� Ĭ��ΪNothing
'����ֵ��(CommandBarButton)��ť����
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

'���ƣ�AddComboBox
'���ã�����һ����Ͽ�
'������
'       strName(string) ��Ͽ���ʾ����
'       [width](integer)��Ͽ����
'       [parent](CommandBar)���������� Ĭ��ΪNothing
'����ֵ��(CommandBarComboBox)��Ͽ����
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

'���ƣ�AddPopup
'���ã�����һ�������˵�
'������
'       strName(string) ��ʾ����
'       [strAction](string)�������� Ĭ��vbnullstring
'       [parent](CommandBar)���������� Ĭ��ΪNothing
'����ֵ��(CommandBarPopup)�˵�����
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
    Debug.Print "��������δʵ��:AddEdit"
End Function

Public Function AddDropdown() As CommandBarControl
    Debug.Print "��������δʵ��:AddDropdown"
End Function

'���ƣ�GetControl
'���ã���ȡ�ؼ�����
'������
'       strName(string)�ؼ�����
'       [parent]������
'����ֵ��(CommandBarControl)�ؼ����� ���û���ҵ�����Nothing
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
        'û���ҵ�����
        Err.Clear
        Set GetControl = Nothing
    Else
        On Error GoTo 0
        Err.Raise Err.Number, Err.Source, "[" & CLASSNAME & ".GetControl]" & Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

'���ƣ�DeleteControl
'���ã���������ɾ���ؼ�
'������
'       strControlName(string)�ؼ�����
'       [parent]�ؼ�����
'����ֵ��(boolean)
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

'Excel���������======================================================================
'���ƣ�ExistsSheet
'���ã�����Ƿ����ָ�����Ƶ�Sheet
'������
'       strSheetName(string)������
'       [book](object)Ĭ��Ϊ��ǰ������
'����ֵ��(boolean)����TrueΪ���ҵ���
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

'���ƣ�ExistsWorkbook
'���ã��Ƿ�������
'������
'       strWorkbookName(string)���������� ��"test.xlsx"
'       [app](object)Ӧ�ö���
'����ֵ��(boolean)����TrueΪ���ҵ���������������Ҳͬ������
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

'��ʽ======================================================================
'���ƣ�SetBorderLine
'���ã����ñ߿���
'������
'       rng(Range) ��Ԫ������
'       [flag](xlBorderFlag)�߿�   �������ԭ�б߿���,���� xlsNone
'       [style](XlLineStyle)�߿���ʽ
'       [weight](XlBorderWeight)�߿��ߴ�ϸ
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

'������======================================================================
'���ƣ�LoadAddin
'���ã�װ�ؼ�����
'����:
'       strPath(string)������·��
'       [Quiet](boolean)�Ƿ�Ĭ���滻�Ѿ���Ϊ�ļ�
'       [app](object)Ӧ�ö���
'����ֵ��(boolean)�Ƿ���سɹ�
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
    
    '����ļ�����
    If p_File.ExistsFile(strPath) = False Then
        '�ļ�������
        LoadAddin = False
        Exit Function
    End If
    
    '����Ƿ�ͬ���ļ���
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
            If MsgBox("�ļ������Ѵ򿪵ĳ�ͻ���Ƿ�ر�ͬ���ļ�������򿪣�", vbYesNo, "��ʾ") = vbYes Then
                UnloadAddin p_strBook, p_app
                GoTo LOADADDIN_OPEN
            End If
        End If
        'ͬ���ļ���ͻ
        LoadAddin = False
    End If
    
    Set p_File = Nothing
End Function

'���ƣ�UnloadAddin
'���ã��رռ�����
'������
'       strName(string)����������
'       [app](object)Ӧ�ö���
'����ֵ��(boolean)ִ�н��
Public Function UnloadAddin(ByVal strName As String, Optional app = Nothing) As Boolean
    Dim p_app

    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    '����ļ��Ƿ����
    If Me.ExistsWorkbook(strName, p_app) = True Then
        'ִ��ж�ع���
        p_app.Run (strName & "!modAutoload.Auto_Close")
        'p_app.Workbooks(strName).RunAutoMacros xlAutoClose
        '�ر��ļ�
        p_app.Workbooks(strName).Close False
        UnloadAddin = True
    Else
        UnloadAddin = False
    End If
End Function

'���ƣ�LoadAddinWithValidate
'���ã�����֤�ļ���
'������
'       strPath(string)������·��
'       strName(string)����������
'       strVersion(string)������汾
'       [app](object)Ӧ�ö���
'����ֵ��(boolean)ִ�н��
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
        '��֤�汾
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
        '����ʧ��
        LoadAddinWithValidate = False
    End If
End Function

'���ƣ�GetAddinVersion
'���ã���ȡ������汾��Ϣ���汾��Ϣͨ������ [modGlobal]��GetVersion������ȡ��
'������
'       strAddinName(string)�ļ�����(������׺) ��"test.xla"
'       [app](object)Ӧ�ö���
'����ֵ��(string)�汾��Ϣ
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

'���ƣ�GetAddinName
'���ã���ȡ������汾��Ϣ���汾��Ϣͨ������ [modGlobal]��GetAppname������ȡ��
'������
'       strAddinName(string)�ļ�����(������׺) ��"test.xla"
'       [app](object)Ӧ�ö���
'����ֵ��(string)�汾��Ϣ
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

'���ƣ�GetAddin
'���ã���ȡ���������
'������
'       strShortName(string)���������� �� "test"
'       [app](object) Ӧ�ö���
'����ֵ��(Addin)��������� ���û�ҵ�����nothing
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

'���ƣ�UninstallAddin
'���ã�ж�ؼ�����
'������
'       strShortName(string)���������� ��"test"
'       [app](object)Ӧ�ö���
'����ֵ��(boolean)���
Public Function UninstallAddin(ByVal strShortName As String, Optional app = Nothing) As Boolean
    Dim p_addin As AddIn
    
    Set p_addin = GetAddin(strShortName, app)
    If p_addin Is Nothing Then
        'û�иü�����
        UninstallAddin = False
    Else
        p_addin.Installed = False
        Kill p_addin.FullName
        UninstallAddin = True
    End If
End Function

'���ƣ�InstallAddin
'���ã���װ������
'������
'       strPath(string)�ļ���ַ
'       [app](object)
'����ֵ��(boolean)��װ���
Public Function InstallAddin(strPath As String, Optional app = Nothing) As Boolean
    Dim p_File As New clsFile, p_addin As AddIn, p_app
    Dim strLocalPath As String, strLocalAddinPath As String
    
    strLocalPath = VBA.Environ$("appdata") & "\microsoft\addins\"
    If app Is Nothing Then
        Set p_app = m_app
    Else
        Set p_app = app
    End If
    '����ļ��Ƿ����
    If p_File.ExistsFile(strPath) = False Then
        InstallAddin = False
        Exit Function
    End If
    'ж��ͬ��������
    UninstallAddin p_File.GetShortFilename(strPath)
    '�����ļ���ָ��Ŀ¼
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
