Attribute VB_Name = "M_ErrorAPI"
Option Explicit

Private Declare PtrSafe Function GetLastError Lib "kernel32.dll" () As Long
Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
        ByVal dwFlags As FORMAT_MESSAGE_FLAGS, _
        ByRef lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        ByRef Arguments As LongPtr _
    ) As Long

Private Enum FORMAT_MESSAGE_FLAGS
    MAX_WIDTH_MASK = &HFF&

    ALLOCATE_BUFFER = &H100& 'FormatMessage ���ŕ�����̈�����蓖�ĂĂ��炤(���ʂ̎擾�ɂ͗v����������)�B
    IGNORE_INSERTS = &H200&
    FROM_STRING = &H400&
    FROM_HMODULE = &H800&
    FROM_SYSTEM = &H1000& '�V�X�e�����烁�b�Z�[�W���擾����(DLL�֐��̃G���[�擾���Ȃ�)
    ARGUMENT_ARRAY = &H2000&
End Enum

Public Function ShowErrorMessage() As String
    
    Dim er As Long
    er = GetLastError()

   ShowErrorMessage = GetDllErrorMessage(er)
    
End Function

'DLL �֐��̃G���[���b�Z�[�W���擾����B
'dwMessageId    :�G���[���b�Z�[�W�� Id�B�ȗ����� Err.LastDllError ���g�p�����B
Public Function GetDllErrorMessage( _
        Optional ByVal dwMessageId As Long = 0 _
    ) As String

    '�����ȗ��Ή��B
    If dwMessageId = 0 Then _
        dwMessageId = VBA.Information.Err().LastDllError

    'ALLOCATE_BUFFER ���w�肵�Ȃ����߁A���O�ŗ̈���m�ۂ���B
    Dim paddingSize As Long
    paddingSize = &HFF
    Const paddingChar = VBA.Constants.vbNullChar

    Dim apiResult As Long
    Do
        '���b�Z�[�W�p�̗̈�m�ہB
        Dim lpBuffer As String
        lpBuffer = VBA.Strings.String$(paddingSize, paddingChar)
        Dim nSize As Long
        nSize = VBA.Strings.Len(lpBuffer)

        apiResult = FormatMessage( _
            FROM_SYSTEM Or MAX_WIDTH_MASK, _
            0, _
            dwMessageId, _
            0, _
            lpBuffer, _
            nSize, _
            0)

        '���s��(���̈�s����)�� 0 �ɂȂ�B
        If apiResult <> 0 Then _
            Exit Do

        '�m�ۃT�C�Y��傫�����čăg���C�B
        paddingSize = paddingSize * 2
    Loop

    '�K�v�Ȕ͈͂����擾���ďo��(apiResult �̌��ʂ��̂܂܂͎g���ɂ���)�B
    Let GetDllErrorMessage = VBA.Strings.Left$(lpBuffer, VBA.Strings.InStr(1, lpBuffer, paddingChar) - 1)
End Function

