VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'        :::::         ʹ�� Windows API ������ MD5          :::::
'  * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *
'           -=  �� �� �� ��  =-
'     GetMD5Text        ��ȡ�ϴμ����MD5���ı�
'     MD5Bytes          ����һ���ֽ������MD5��
'     MD5String         ����һ���ַ�����ANSI���룩��MD5��
'     MD5File           ����һ���ļ���MD5��
'  * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *
'
Option Explicit

' ==============================
' ��     �Զ��� ��������      ��
' ==============================
Private Type MD5_CTX
      dwNUMa      As Long
      dwNUMb      As Long
      Buffer(15)  As Byte
      cIN(63)     As Byte
      cDig(15)    As Byte
End Type

' ==============================
' ��     API �� �� �� ��      ��
' ==============================
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)
  Private Declare PtrSafe Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)
  Private Declare PtrSafe Sub MD5Update Lib "advapi32" (lpContext As MD5_CTX, ByRef lpBuffer As Any, ByVal BufSize As Long)
#Else
  Private Declare Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)
  Private Declare Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)
  Private Declare Sub MD5Update Lib "advapi32" (lpContext As MD5_CTX, ByRef lpBuffer As Any, ByVal BufSize As Long)
#End If
Private stcContext   As MD5_CTX

' ==============================
' ��     ͨ�� ���� & ����     ��
' ==============================
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' ��    �ܣ�����һ���ַ�����ANSI���룩��MD5��
' ��ڲ�����
'  strText     �ַ����ı�
' ���ز�����   MD5�� ��16�ֽڵ�Byte���飩
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MD5String(strText As String) As String
  Dim aBuffer()     As Byte

   Call MD5Init(stcContext)
   If (Len(strText) > 0) Then
      aBuffer = VBA.StrConv(strText, vbFromUnicode)
      Call MD5Update(stcContext, aBuffer(0), UBound(aBuffer) + 1)
   Else
      Call MD5Update(stcContext, 0, 0)
   End If
   Call MD5Final(stcContext)
   'MD5String = stcContext.cDig
   '����Ϊ�ַ���
   MD5String = GetMD5Text
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' ��    �ܣ�����һ���ֽ�����MD5��
' ��ڲ�����
'  Buffer      Byte����
'  size        ���ȣ���ѡ��Ĭ�ϼ����������ȣ�
' ���ز�����   MD5�� ��16�ֽڵ�Byte���飩
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MD5Bytes(Buffer() As Byte, _
                        Optional ByVal size As Long = -1) As Byte()
      Dim U As Long, pBase   As Long

   pBase = LBound(Buffer)
   U = UBound(Buffer) - pBase
   If (-1 = size) Then size = U + 1
   Call MD5Init(stcContext)
   If (-1 = U) Then
      Call MD5Update(stcContext, 0, 0)
   Else
      Call MD5Update(stcContext, Buffer(pBase), size)
   End If
   Call MD5Final(stcContext)
   MD5Bytes = stcContext.cDig
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' ��    �ܣ�����һ���ļ���MD5��
' ��ڲ�����
'  FileName    �����ļ���������·����
' ���ز�����   MD5�� ��16�ֽڵ�Byte���飩
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MD5File(ByVal Filename As String) As Byte()
    Const BUFFERSIZE  As Long = 1024& * 512      ' ������ 512KB
    Dim DataBuff() As Byte
    Dim lFileSize  As Long
    Dim iFn        As Long

    On Error GoTo E_Handle_MD5
    If (Len(Dir$(Filename)) = 0) Then Err.Raise 5      '�ļ�������
    ReDim DataBuff(BUFFERSIZE - 1)
    iFn = FreeFile()
    Open Filename For Binary As #iFn
    lFileSize = LOF(iFn)
    Call MD5Init(stcContext)
    If (lFileSize = 0) Then
        Call MD5Update(stcContext, 0, 0)
    Else
        Do While (lFileSize > 0)
            Get iFn, , DataBuff
            If (lFileSize > BUFFERSIZE) Then
                Call MD5Update(stcContext, DataBuff(0), BUFFERSIZE)
            Else
                Call MD5Update(stcContext, DataBuff(0), lFileSize)
            End If
            lFileSize = lFileSize - BUFFERSIZE
        Loop
    End If
    Close iFn
    Call MD5Final(stcContext)
E_Handle_MD5:
    MD5File = stcContext.cDig
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' ��    �ܣ���ȡ�ϴμ����MD5���ı�
' ��ڲ�����   < �� >
' ���ز�����   MD5���ı��ַ�����û��MD5���� ���ؿմ���
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetMD5Text() As String
    Dim sResult As String, i&
    If (stcContext.dwNUMa = 0) Then
        sResult = vbNullString
    Else
        sResult = VBA.Space$(32)
        For i = 0 To 15
            Mid$(sResult, i + i + 1) = VBA.Right$("0" & VBA.Hex$(stcContext.cDig(i)), 2)
        Next
    End If
    GetMD5Text = sResult       ' LCase$(sResult) '��ĸСд
End Function