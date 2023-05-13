Attribute VB_Name = "M_ElementFromPoint"
Option Explicit

'GetCursorPos�֐�'
'�}�E�X�J�[�\���̍��W��POINTAPI�\���̂Ƃ��Ď擾����'
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'POINTAPI�\����'
Public Type POINTAPI
    x As Long
    y As Long
End Type


'DispCallFunc�֐�'
'�N���X�C���X�^���X�̊֐��|�C���^���g�p���邱�ƂŁA'
'�C�ӂ̃I�v�V�����i�Ăяo���K�񓙁j�ŃI�u�W�F�N�g�̃��\�b�h�����s����'
Public Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal offsetinVft As LongPtr, _
    ByVal CallConv As Long, _
    ByVal retTYP As Integer, _
    ByVal paCNT As Long, _
    ByRef paTypes As Integer, _
    ByRef paValues As LongPtr, _
    ByRef retVAR As Variant _
) As Long

'CUIAutomation::ElementFromPoint��DispCallFunc�o�R�ŌĂяo�����߂̒萔�Z�b�g'
Public Const S_OK = 0
Public Const CC_STDCALL As Long = 4

#If Win64 Then
    Public Const pElementFromPoint As Long = 56
    'Address of 8th Function in the virtual function table of the CUIAutomation Class : (8-1)th * 8 Byte'
    Public Const pCount = 2
    'Number of arguments in 64bit environment ( pt, Element* )'
#Else
    Public Const pElementFromPoint As Long = 28
    'Address of 8th Function in the virtual function table of the CUIAutomation Class : (8-1)th * 4 Byte'
    Public Const pCount = 3
    'Number of arguments in 32bit environment ( pt.x, pt.y, Element* )'
#End If

'�w�肳�ꂽ���W�̃G�������g���擾���郁�\�b�h'
Public Function ElementFromPoint(ByRef uia As CUIAutomation, ByRef pt As POINTAPI) As IUIAutomationElement

    Dim Element As IUIAutomationElement
    Dim vParams(pCount) As Variant     'ElementFromPoint�ɓn���e��������A�o���A���g�^�̔z��Ƃ��ď���'
    Dim vParamPtr(pCount) As LongPtr   'ElementFromPoint�Ɏg���e������̃|�C���^���A�z��Ƃ��ď���'
    Dim vParamType(pCount) As Integer  'ElementFromPoint�Ɏg���e������̌^�������l���AInteger�^�̔z��Ƃ��ď���'


    '�����{�̂̊i�[�����i32bit��64bit�ŏ���������j'
    #If Win64 Then
    
        '64bitExcel�ł�WOW64��ʂ������ڊ֐����Ăяo�����B'
        'stdcall�Ăяo���K�����������邽�߁A������CPU���ł̓X�^�b�N�̈�ł͂Ȃ�'
        '���W�X�^�ɂ��̂܂ܕ��荞�܂��B���̌��ʁACUIAutomation�̃C���X�^���X�̈��'
        '�s�K�؂ȃA�h���X�ɒl���n����ă������j�󂪋N����Excel���N���b�V�����郊�X�N������B'
        '�����h�����߁A���O��PINTAPI�̊e�����o���A�Ăяo����Ŋi�[�����tagPOINT��'
        '�݊���������P��̕ϐ��Ɏ蓮�Ŋi�[���Ă����K�v������B'
        Dim llpt As LongLong
        llpt = pt.y * (2 ^ 32) + pt.x
        'Shift y left by 32 bits and then add x to put the two parameters into one variable'
        '0000YYYY pt.y '
        'YYYY0000 pt.y * 2 ^ 32 '
        'YYYYXXXX (pt.y * 2 ^ 32) + pt.x '
        
        vParams(0) = llpt
        vParams(1) = VarPtr(Element)
    
    #Else
    
        '32bit���ł�WOW64�̒��ԏ����ō\���̂̊e�����o�ɓK�؂ɒl���n����邽�߁A�n����̎d�l��z������K�v������'
        vParams(0) = pt.x
        vParams(1) = pt.y
        vParams(2) = VarPtr(Element)
    
    #End If

    '�����̏��i�^�ƃA�h���X�j�̊i�[����'
    Dim pIndex As Long
    For pIndex = 0 To pCount
        vParamPtr(pIndex) = VarPtr(vParams(pIndex))   'ElementFromPoint�Ɏg���e������̃|�C���^'
        vParamType(pIndex) = VarType(vParams(pIndex)) 'ElementFromPoint�Ɏg���e������̌^�������l'
    Next

    Dim lRtn As Variant  'DispCallFunc�̐��۔�����i�[����ϐ�'
    Dim vRtn As Variant  'ElementFromPoint�̐��۔�����i�[����ϐ�'

    lRtn = DispCallFunc(ObjPtr(uia), pElementFromPoint, CC_STDCALL, vbLong, pCount, vParamType(0), vParamPtr(0), vRtn)
    '�@�@�@�@�@�@�@�@�@�@�@ByVal�@�@�@�@�@ByVal�@�@�@    �@ByVal�@ �@ByVal �@ByVal�@�@�@ByRef�@�@�@�@ByRef�@�@�@ ByRef'
                
                
    '��DispCallFunc���ǂ̂悤�ɂ���ElementFromPoint���Ăяo���Ă���̂��̉�ǁ�'
    
    '�i�������jCUIAutomation�N���X�́A'
    '�i�������j���z�֐��e�[�u����́A8�Ԗڂ̃|�C���^�������A�h���X�ɓW�J����Ă��鏈�������s���Ă��������B'
    
    '�i��O�����j�Ăяo���K��́uCC_STDCALL�v�ł��B��64bit�ł�Excel�ł́A���̒l�͖��������'

    '�i��l�����j���s���������̖߂�l�iElementFromPoint�̐��۔���j��Long�^�Ŏ󂯎��܂��B'

    '�i��܈����j������64bit�łł�2�A32bit�łł�3����܂��B�i�l�n���j'

    '�i��Z�����j�u�����̌^�̎�ނ������l�v���i�[�����z��͂����ɂ���܂��B�i�Q�Ɠn���j'

    '�i�掵�����j�u�����{�̂����݂��郁�����A�h���X�̒l�v���i�[�����z��͂����ɂ���܂��B�i�Q�Ɠn���j'

    '�i�攪�����j���s���������̖߂�l�iElementFromPoint�̐��۔���j�͂���(vRet)�Ɋi�[���Ă��������B�i�Q�Ɠn���j'
    
                
    If lRtn = S_OK Then
        If vRtn = S_OK Then
            If Not Element Is Nothing Then
                Set ElementFromPoint = Element
                Set Element = Nothing
                'Element��IUnknown::Release�ŉ�����悤�Ƃ���ƁA
                '�ďo�����Q�Ƃ��悤�Ƃ���C���^�[�t�F�[�X������Excel���N���b�V������悤�Ȃ̂Œ���'
            End If
        Else
            SetLastError vRtn
            Debug.Print ShowErrorMessage
        End If
    Else
        SetLastError lRtn
        Debug.Print ShowErrorMessage
    End If
    
End Function


'���݂̃J�[�\���ʒu�̃G�������g���擾���郁�\�b�h'
'�iGetCursorPos���֐����Ɉϑ����Ă���̂ŁA�J�[�\���ȊO�̍��W��n���K�v�������Ȃ炱������g�������y�j'
Public Function ElementFromCursor(ByRef uia As CUIAutomation) As IUIAutomationElement

    Dim pt As POINTAPI
    GetCursorPos pt
    
    Dim elem As IUIAutomationElement
    Set elem = ElementFromPoint(uia, pt)

    If Not elem Is Nothing Then
        Set ElementFromCursor = elem
    End If

End Function
