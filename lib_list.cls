VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lib_list"
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
'
'構文：
'   Public Sub push(ByVal value As Variant)
'    リターン値
'        なし
'   Public Function pop() As Variant
'    リターン値
'        STACKにデータがある場合 -> データを渡す
'        STACKにデータがない場合 -> Emptyを渡す
'
'実行環境前提：
'   lib_stackクラスが存在すること
'---------------------------------------


'定数
Private Const chunk_size As Long = 50
Private Const an_start As Long = -1
Private Const an_end As Long = -2

'双方向リスト用のコンテナ構造体設定
Private Type struct_Container
    value As Variant
    prev As Long
    next As Long
End Type

'コンテナ定義
Private Container() As struct_Container
Private element As Long

'双方向リストの末端のアンカー
Private anchor_start As Long
Private anchor_end As Long

'コンテナのフリーリスト
Private free As lib_stack


'---------------------------------------
'オブジェクト生成
'---------------------------------------
Private Sub Class_Initialize()

    'freeリストの生成
    Set free = New lib_stack
    
    '初期値設定
    element = 0
    anchor_start = an_end
    anchor_end = an_start
    
    'リスト用配列初期確保
    Call extend_Container


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
    
    Dim index As Long
    Dim buf As Variant
    
    'データ格納のため空コンテナを確保
    Do
        buf = free.pop
        If IsEmpty(buf) Then
            '空きコンテナがない場合、コンテナを拡張子リトライする
            Call extend_Container
        End If
    Loop While IsEmpty(buf)
    index = CLng(buf)
    
    '空コンテナへデータ入力
    With Container(index)
        .value = value
        .next = an_end
        .prev = anchor_end
    End With
    
    '一つ前のコンテナと開始アンカーのポインタ変更
    If anchor_end < 0 Then
        'リスト上にデータがない場合
        anchor_start = index
    Else
        'リスト上にデータが一つ以上ある場合
        Container(anchor_end).next = index
    End If
    
    'ENDアンカーのポインタ変更
    anchor_end = index
    
End Sub


'---------------------------------------
'pullメソッド(リストからデータを取り出し)
'*リストにデータがある場合 => 該当のデータをリターン
'*リストにデータがない場合 => Empty(Variant)をリターン
'---------------------------------------
Public Function pull() As Variant

    Dim first As Long
    Dim second As Long

    'データがない場合は終了する
    If anchor_start < 0 Then GoTo Failure

    '先頭コンテナからデータと2番目のコンテナポインタを抜き出し
    first = anchor_start
    With Container(first)
        second = .next
        pull = .value
    End With
    
    '先頭コンテナの開放
    If second < 0 Then
        '最後のコンテナ開放(anchor_end -> anchor_start)
        anchor_end = an_start
    Else
        '他のコンテナが存在する場合
        Container(second).prev = an_start
    End If
    anchor_start = Container(first).next
    free.push (first)

Success:
    Exit Function
    
Failure:
    Exit Function

End Function


'---------------------------------------
'リスト配列の拡張
'---------------------------------------
Private Sub extend_Container()

    Dim old_element As Long
    Dim i As Long
    
    '現在の配列インデックスの保存
    old_element = element
    If old_element <= 0 Then old_element = -1

    '配列の拡張
    ReDim Preserve Container(element + chunk_size)
    element = UBound(Container)
    
    '拡張したコンテナをfreeリストへ登録
    For i = old_element + 1 To element
        free.push (i)
    Next
    
End Sub

