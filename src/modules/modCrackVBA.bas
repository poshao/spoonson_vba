Attribute VB_Name = "modCrackVBA"
Option Explicit
'�ƽ�VBA�༭��������
'sample
'=====================================
'sub Test_Crack()
'    CrackVBA()
'end sub
'sub Test_Recover()
'    RecoverVBA()
'end sub
'=====================================
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As Long, ByVal Length As Long)
    Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, ByVal pTemplateName As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As Long, ByVal Length As Long)
    Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, ByVal pTemplateName As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
#End If

Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long, Flag As Boolean

Public Sub RecoverVBA()
    '���Ѿ�hook,��ָ�ԭAPI��ͷ��6�ֽ�,Ҳ���ǻָ�ԭ�������Ĺ���
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function CrackVBA() As Boolean
    Dim TmpBytes(0 To 5) As Byte
    Dim p As Long
    Dim OriginProtect As Long
   
    CrackVBA = False
   
    'VBE6.dll����DialogBoxParamA��ʾVB6INTL.dll��Դ�еĵ�4070�ŶԻ���(������������Ĵ���)
    '��DialogBoxParamA����ֵ��0,��VBE����Ϊ������ȷ,��������Ҫhook DialogBoxParamA����
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
   
    '��׼api hook����֮һ: �޸��ڴ�����,ʹ���д
    If VirtualProtect(ByVal pFunc, 6, &H40, OriginProtect) <> 0 Then
        '��׼api hook����֮��: �ж��Ƿ��Ѿ�hook,����API�ĵ�һ���ֽ��Ƿ�Ϊ&H68,
        '������˵���Ѿ�Hook
        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
        If TmpBytes(0) <> &H68 Then
            '��׼api hook����֮��: ����ԭ������ͷ�ֽ�,������6���ֽ�,�Ա�����ָ�
            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6
            '��AddressOf��ȡMyDialogBoxParam�ĵ�ַ
            '��Ϊ�﷨������д��p = AddressOf MyDialogBoxParam,��������дһ������
            'GetPtr,���ý����Ƿ���AddressOf MyDialogBoxParam��ֵ,�Ӷ�ʵ�ֽ�
            'MyDialogBoxParam�ĵ�ַ����p��Ŀ��
            p = GetPtr(AddressOf MyDialogBoxParam)
            
            '��׼api hook����֮��: ��װAPI��ڵ��´���
            'HookBytes ������»��
            'push MyDialogBoxParam�ĵ�ַ
            'ret
            '��������ת��MyDialogBoxParam����
            HookBytes(0) = &H68
            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
            HookBytes(5) = &HC3
            
            '��׼api hook����֮��: ��HookBytes�����ݸ�дAPIǰ6���ֽ�
            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
            '����hook�ɹ���־
            Flag = True
            CrackVBA = True
        End If
    End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
        ByVal pTemplateName As Long, ByVal hWndParent As Long, _
        ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
    If pTemplateName = 4070 Then
        '�г������DialogBoxParamAװ��4070�ŶԻ���,��������ֱ�ӷ���1,��
        'VBE��Ϊ������ȷ��
        MyDialogBoxParam = 1
    Else
        '�г������DialogBoxParamA,��װ��Ĳ���4070�ŶԻ���,�������ǵ���
        'RecoverBytes�����ָ�ԭ�������Ĺ���,�ڽ���ԭ���ĺ���
        RecoverVBA
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                           hWndParent, lpDialogFunc, dwInitParam)
        'ԭ���ĺ���ִ�����,�ٴ�hook
        CrackVBA
    End If
End Function

Private Function GetPtr(ByVal Value As Long) As Long
    '��ú����ĵ�ַ
    GetPtr = Value
End Function

