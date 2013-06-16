VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lib_stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------
'概要：Excel VBA スタックオブジェクト
'使い方：
'   -オブジェクト生成
'        Dim stack as lib_stack
'        Set stack = New lib_stackオブジェクト生成
'   -データ投入
'        stack.push( VARIANT_DATA )
'   -データ取り出し
'        VARIANT_DATA = stack.pop
'構文：
'   Public Sub push(ByVal value As Variant)
'    リターン値
'        なし
'   Public Function pop() As Variant
'    リターン値
'        STACKにデータがある場合 -> データを渡す
'        STACKにデータがない場合 -> Emptyを渡す
'---------------------------------------


Private Const chunk_size As Long = 50

Private data() As Variant 'スタックデータ格納用配列
Private element As Long   'スタック配列要素数長
Private depth As Long     'スタック現在の深さ


'---------------------------------------
'オブジェクト生成
'---------------------------------------
Private Sub Class_Initialize()

    '初期値設定
    element = 0
    depth = 0
    
    'スタック用配列初期確保
    Call extend_chunk

End Sub


'---------------------------------------
'オブジェクト破棄
'---------------------------------------
Private Sub class_terminate()

End Sub


'---------------------------------------
'pushメソッド(データをスタックに追加)
'---------------------------------------
Public Sub push(ByVal value As Variant)
    
    'データの格納
    data(depth) = value
    
    '現在のスタックの深さをインクリメントする
    depth = depth + 1
    
    'スタックデータ配列をすべて利用している場合は配列を拡張する
    If depth > element Then
        Call extend_chunk
    End If

End Sub


'---------------------------------------
'popメソッド(スタックからデータを取り出し)
'*スタックにデータがある場合 => 該当のデータをリターン
'*スタックにデータがない場合 => Empty(Variant)をリターン
'---------------------------------------
Public Function pop() As Variant

    'データがない場合は終了する
    If depth <= 0 Then GoTo Failure

    '配列の最後のデータをリターン値に設定
    pop = data(depth - 1)
    
    '現在のスタックの深さをデクリメント
    depth = depth - 1
    
    '配列を縮小
    '(データ配列が2チャンク以上あいている場合)
    If depth < element - 2 * chunk_size Then
        Call reduce_chunk
    End If


Success:
    Exit Function
    
Failure:
    Exit Function

End Function

'---------------------------------------
'スタック配列の拡張
'---------------------------------------
Private Sub extend_chunk()

    '配列の拡張
    ReDim Preserve data(element + chunk_size)
    element = UBound(data)
    
End Sub


'---------------------------------------
'スタック配列の縮小
'---------------------------------------
Private Sub reduce_chunk()

    Dim i As Long
    i = element - chunk_size
    
    '状態確認
    '(最後のチャンクまたは、利用中のelementがある場合は縮小を中断する)
    If i < chunk_size Or i < depth Then
        GoTo Failuer
    End If

    '配列の縮小
    ReDim Preserve data(i)
    element = UBound(data)

Success:
    Exit Sub
    
Failuer:
    Exit Sub

End Sub
