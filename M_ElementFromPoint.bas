'Module UIA_ElementFromPoint

Option Explicit

Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal offsetinVft As Long, _
    ByVal CallConv As Long, _
    ByVal retTYP As Integer, _
    ByVal paCNT As Long, _
    ByRef paTypes As Integer, _
    ByRef paValues As LongPtr, _
    ByRef retVAR As Variant _
) As Long

Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As PointAPI) As Long

Private Type PointAPI: x As Long: y As Long: End Type

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Public Function ElementFromPoint(pt As PointAPI) As IUIAutomationElement

    'IUIAutomationのCOMインターフェースのアドレスを用意
    Dim uia As New CUIAutomation
    Dim CUIAutomationのインスタンスのアドレス As LongPtr
    CUIAutomationのインスタンスのアドレス = ObjPtr(uia)

    'ElementFromPontメソッドの、IUIAutomationの仮想関数テーブル内での定義番号と、
    'それを元にした関数ポインタ格納位置のオフセット値を用意
    Dim ElementFromPointの関数ID As Long
    ElementFromPointの関数ID = 7 '0スタートで7番目の関数。この値はoleView等のツールで調査が必要
    Dim ElementFromPoint関数ポインタ座標 As Long
    #If Win64 Then
        ElementFromPoint関数ポインタ座標 = ElementFromPointの関数ID * 8
    #Else
        ElementFromPoint関数ポインタ座標 = ElementFromPointの関数ID * 4
    #End If
    
    'ElementFromPointメソッドの呼び出し規約を用意
    Const STDCALL As Long = 4

    'ElementFromPointに渡す２つの引数の実体を用意 ※絶対にVariant型じゃないとダメ

    '引数1：座標情報
    '64bit整数(LongLong型)にキャストしないとDouble型として扱われ、値が壊れる。
    '同じ64bitサイズのCurrency型は固定小数点（整数部32bit＋小数部32bit）として扱われ、
    '64bit整数の全ビットを保持できないため、必ずLongLong型にキャストする必要がある。
    '32bit環境だとSTDCALL呼び出し規約は、引数をスタックに積み上げるので構造体渡しと同じ形になる
    #If Win64 Then
        Dim 座標値 As Variant
        座標値 = CLngLng(pt.y * (2 ^ 32) + pt.x)
    #Else
        Dim 座標値_上位ビット As Variant, 座標値_下位ビット As Variant
        座標値_下位ビット = pt.x
        座標値_上位ビット = pt.y
    #End If

    '引数2：最終的に受け取りたいIUIAutomation変数のポインタ
    'シグネチャがIUIAutomationElement**なので、変数のポインタを用意する必要がある。
    Dim ElemPtr As Variant, Element As IUIAutomationElement
    ElemPtr = VarPtr(Element)

    'この引数1,引数2はDispCallFuncからの直接的な干渉を受けるので、
    'C言語のVARIANTとバイナリレベルで互換性があるVariant型にしておく必要がある。
    'LongPtrなどのようなプリミティブ型はVBA内部ではC言語のネイティブな型とは違い、
    '[変数テーブル＋型情報]という構造体として保管されているため、
    'この構造体にC言語側の変数定義でアクセスされるとメモリ破壊が発生し、Excelが即クラッシュする。

    #If Win64 Then
        '用意した引数の実体の型情報を定義した配列を用意
        Dim 引数の型情報の配列(0 To 1) As Integer
        引数の型情報の配列(0) = VarType(座標値)
        引数の型情報の配列(1) = VarType(ElemPtr)
    
        '用意した引数の実体のアドレス情報を定義した配列を用意
        Dim 引数のアドレス情報の配列(0 To 1) As LongPtr
        引数のアドレス情報の配列(0) = VarPtr(座標値)
        引数のアドレス情報の配列(1) = VarPtr(ElemPtr)
    
        '用意した引数の実体の個数を定義した変数を用意
        Dim ElementFromPointメソッドに渡す引数の個数 As Long
        ElementFromPointメソッドに渡す引数の個数 = 2
    #Else
        Dim 引数の型情報の配列(0 To 2) As Integer
        引数の型情報の配列(0) = VarType(座標値_下位ビット)
        引数の型情報の配列(1) = VarType(座標値_上位ビット)
        引数の型情報の配列(2) = VarType(ElemPtr)
        
        Dim 引数のアドレス情報の配列(0 To 2) As LongPtr
        引数のアドレス情報の配列(0) = VarPtr(座標値_下位ビット)
        引数のアドレス情報の配列(1) = VarPtr(座標値_上位ビット)
        引数のアドレス情報の配列(2) = VarPtr(ElemPtr)
        
        Dim ElementFromPointメソッドに渡す引数の個数 As Long
        ElementFromPointメソッドに渡す引数の個数 = 3
    #End If

    'ElementFromPointメソッドの戻り値（成否判定）を何の型で受け取るか指定する変数を用意
    Dim ElementFromPointメソッドの戻り値の型 As Integer
    ElementFromPointメソッドの戻り値の型 = vbLong '定数値：3

    'ElementFromPointメソッドの戻り値（成否判定）を受け取る変数をVariant型で用意
    Dim ElementFromPointメソッドの成否判定 As Variant

    'DispCallFuncメソッド自体の成否判定を受け取る変数を用意
    Dim DispCallFunc成否判定 As Long

    'DispCallFuncメソッドにより、COMインターフェースを迂回して直接ElementFromPointのエントリーポイントをコール
    DispCallFunc成否判定 = DispCallFunc( _
        CUIAutomationのインスタンスのアドレス, _
        ElementFromPoint関数ポインタ座標, _
        STDCALL, _
        ElementFromPointメソッドの戻り値の型, _
        ElementFromPointメソッドに渡す引数の個数, _
        引数の型情報の配列(0), _
        引数のアドレス情報の配列(0), _
        ElementFromPointメソッドの成否判定 _
    )

    If DispCallFunc成否判定 = 0 Then '0:成功を示す値（S_OK定数）
        If ElementFromPointメソッドの成否判定 = 0 Then '0:成功を示す値（S_OK定数）
            Set ElementFromPoint = Element
        Else
            Debug.Print "ElementFromPointメソッドは異常終了しました。戻り値：", ElementFromPointメソッドの成否判定
        End If
    Else
        Debug.Print "DispCallFuncメソッドは異常終了しました。戻り値：", DispCallFunc成否判定
    End If

End Function

Public Function ElementFromCursor() As IUIAutomationElement

    Dim pt As PointAPI, res As Long
    res = GetCursorPos(pt)
    If res = 1 Then
        Dim elem As IUIAutomationElement
        Set elem = ElementFromPoint(pt)
        If Not elem Is Nothing Then
            Set ElementFromCursor = elem
        Else
            Debug.Print "ElementFromPointメソッド失敗"
        End If
    Else
        Debug.Print "GetCursorPos失敗"
    End If

End Function

Public Sub ElementFromPoint動作サンプル()

    If Selection.Value <> "" Then
        MsgBox "空白セルを選択した態状で開始してください。"
        Exit Sub
    Else
        MsgBox "自動ループでカーソル上の要素を取得し続けます。" & vbCrLf & _
        "終了する時はALTキーを押してください。"
    End If
    
    Application.EnableEvents = False

    Do
        Dim elem As IUIAutomationElement
        Set elem = ElementFromCursor
        If Not elem Is Nothing Then
            Debug.Print elem.CurrentName
            Selection.Value = elem.CurrentName
        End If
        If (GetAsyncKeyState(vbKeyMenu) And -32768) = -32768 Then Exit Do
    Sleep 1
    DoEvents
    Loop
    
    Application.EnableEvents = True

End Sub