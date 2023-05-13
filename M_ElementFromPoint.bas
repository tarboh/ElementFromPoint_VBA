Attribute VB_Name = "M_ElementFromPoint"
Option Explicit

'GetCursorPos関数'
'マウスカーソルの座標をPOINTAPI構造体として取得する'
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'POINTAPI構造体'
Public Type POINTAPI
    x As Long
    y As Long
End Type


'DispCallFunc関数'
'クラスインスタンスの関数ポインタを使用することで、'
'任意のオプション（呼び出し規約等）でオブジェクトのメソッドを実行する'
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

'CUIAutomation::ElementFromPointをDispCallFunc経由で呼び出すための定数セット'
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

'指定された座標のエレメントを取得するメソッド'
Public Function ElementFromPoint(ByRef uia As CUIAutomation, ByRef pt As POINTAPI) As IUIAutomationElement

    Dim Element As IUIAutomationElement
    Dim vParams(pCount) As Variant     'ElementFromPointに渡す各種引数を、バリアント型の配列として準備'
    Dim vParamPtr(pCount) As LongPtr   'ElementFromPointに使う各種引数のポインタを、配列として準備'
    Dim vParamType(pCount) As Integer  'ElementFromPointに使う各種引数の型を示す値を、Integer型の配列として準備'


    '引数本体の格納処理（32bitと64bitで処理が分岐）'
    #If Win64 Then
    
        '64bitExcelではWOW64を通さず直接関数が呼び出される。'
        'stdcall呼び出し規約も無視されるため、引数はCPU内ではスタック領域ではなく'
        'レジスタにそのまま放り込まれる。その結果、CUIAutomationのインスタンス領域の'
        '不適切なアドレスに値が渡されてメモリ破壊が起こりExcelがクラッシュするリスクがある。'
        'これを防ぐため、事前にPINTAPIの各メンバを、呼び出し先で格納されるtagPOINTと'
        '互換性がある単一の変数に手動で格納しておく必要がある。'
        Dim llpt As LongLong
        llpt = pt.y * (2 ^ 32) + pt.x
        'Shift y left by 32 bits and then add x to put the two parameters into one variable'
        '0000YYYY pt.y '
        'YYYY0000 pt.y * 2 ^ 32 '
        'YYYYXXXX (pt.y * 2 ^ 32) + pt.x '
        
        vParams(0) = llpt
        vParams(1) = VarPtr(Element)
    
    #Else
    
        '32bit環境ではWOW64の中間処理で構造体の各メンバに適切に値が渡されるため、渡し先の仕様を配慮する必要が無い'
        vParams(0) = pt.x
        vParams(1) = pt.y
        vParams(2) = VarPtr(Element)
    
    #End If

    '引数の情報（型とアドレス）の格納処理'
    Dim pIndex As Long
    For pIndex = 0 To pCount
        vParamPtr(pIndex) = VarPtr(vParams(pIndex))   'ElementFromPointに使う各種引数のポインタ'
        vParamType(pIndex) = VarType(vParams(pIndex)) 'ElementFromPointに使う各種引数の型を示す値'
    Next

    Dim lRtn As Variant  'DispCallFuncの成否判定を格納する変数'
    Dim vRtn As Variant  'ElementFromPointの成否判定を格納する変数'

    lRtn = DispCallFunc(ObjPtr(uia), pElementFromPoint, CC_STDCALL, vbLong, pCount, vParamType(0), vParamPtr(0), vRtn)
    '　　　　　　　　　　　ByVal　　　　　ByVal　　　    　ByVal　 　ByVal 　ByVal　　　ByRef　　　　ByRef　　　 ByRef'
                
                
    '＜DispCallFuncがどのようにしてElementFromPointを呼び出しているのかの解読＞'
    
    '（第一引数）CUIAutomationクラスの、'
    '（第二引数）仮想関数テーブル上の、8番目のポインタが示すアドレスに展開されている処理を実行してください。'
    
    '（第三引数）呼び出し規約は「CC_STDCALL」です。※64bit版のExcelでは、この値は無視される'

    '（第四引数）実行した処理の戻り値（ElementFromPointの成否判定）はLong型で受け取ります。'

    '（第五引数）引数は64bit版では2つ、32bit版では3つあります。（値渡し）'

    '（第六引数）「引数の型の種類を示す値」を格納した配列はここにあります。（参照渡し）'

    '（第七引数）「引数本体が存在するメモリアドレスの値」を格納した配列はここにあります。（参照渡し）'

    '（第八引数）実行した処理の戻り値（ElementFromPointの成否判定）はここ(vRet)に格納してください。（参照渡し）'
    
                
    If lRtn = S_OK Then
        If vRtn = S_OK Then
            If Not Element Is Nothing Then
                Set ElementFromPoint = Element
                Set Element = Nothing
                'ElementをIUnknown::Releaseで解放しようとすると、
                '呼出元が参照しようとするインターフェースも壊れてExcelがクラッシュするようなので注意'
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


'現在のカーソル位置のエレメントを取得するメソッド'
'（GetCursorPosも関数側に委託しているので、カーソル以外の座標を渡す必要が無いならこちらを使う方が楽）'
Public Function ElementFromCursor(ByRef uia As CUIAutomation) As IUIAutomationElement

    Dim pt As POINTAPI
    GetCursorPos pt
    
    Dim elem As IUIAutomationElement
    Set elem = ElementFromPoint(uia, pt)

    If Not elem Is Nothing Then
        Set ElementFromCursor = elem
    End If

End Function
