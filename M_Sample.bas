Attribute VB_Name = "M_Sample"
Option Explicit

Sub ElemntFromPointサンプル()
    
    '実行すると、Escキーやリセットボタンなどで処理を止めるまでカーソル上のエレメントを取得し続けます
    
    Dim uia As New CUIAutomation
    Dim elem As IUIAutomationElement
    Dim pt As POINTAPI
    
    Do
        
        'カーソル座標取得
        GetCursorPos pt
        
        '座標上のエレメントを取得
        Set elem = ElementFromPoint(uia, pt)
        
        If Not elem Is Nothing Then
            Debug.Print (elem.CurrentName)
            Set elem = Nothing
        End If
        
        DoEvents
        
    Loop
    
End Sub


Sub ElementFromCursorサンプル()
    
    '実行すると、Escキーやリセットボタンなどで処理を止めるまでカーソル上のエレメントを取得し続けます
    
    Dim uia As New CUIAutomation
    Dim elem As IUIAutomationElement
    
    Do

        'カーソル上のエレメントを取得
        Set elem = ElementFromCursor(uia)
        
        If Not elem Is Nothing Then
            Debug.Print (elem.CurrentName)
            Set elem = Nothing
        End If
        
        DoEvents
        
    Loop
    
End Sub
