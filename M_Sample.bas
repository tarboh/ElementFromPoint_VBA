Attribute VB_Name = "M_Sample"
Option Explicit

Sub ElemntFromPoint�T���v��()
    
    '���s����ƁAEsc�L�[�⃊�Z�b�g�{�^���Ȃǂŏ������~�߂�܂ŃJ�[�\����̃G�������g���擾�������܂�
    
    Dim uia As New CUIAutomation
    Dim elem As IUIAutomationElement
    Dim pt As POINTAPI
    
    Do
        
        '�J�[�\�����W�擾
        GetCursorPos pt
        
        '���W��̃G�������g���擾
        Set elem = ElementFromPoint(uia, pt)
        
        If Not elem Is Nothing Then
            Debug.Print (elem.CurrentName)
            Set elem = Nothing
        End If
        
        DoEvents
        
    Loop
    
End Sub


Sub ElementFromCursor�T���v��()
    
    '���s����ƁAEsc�L�[�⃊�Z�b�g�{�^���Ȃǂŏ������~�߂�܂ŃJ�[�\����̃G�������g���擾�������܂�
    
    Dim uia As New CUIAutomation
    Dim elem As IUIAutomationElement
    
    Do

        '�J�[�\����̃G�������g���擾
        Set elem = ElementFromCursor(uia)
        
        If Not elem Is Nothing Then
            Debug.Print (elem.CurrentName)
            Set elem = Nothing
        End If
        
        DoEvents
        
    Loop
    
End Sub
