VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lib_hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------
'概要：Excel VBA ハッシュオブジェクト
'使い方：
'   -オブジェクト生成
'        Dim hash as lib_hash
'        Set hash = New lib_hashオブジェクト生成
'   -データ投入
'        hash.add( HASH_KEY, VARIANT_DATA )
'   -データ参照
'        VARIANT_DATA = hash.show(HASH_KEY)
'   -データ削除
'        hash.del(HASH_KEY)
'
'構文：
'   Public Function add(ByVal key As String, ByVal value As Variant) As Boolean
'    リターン値
'        ハッシュへの追加成功 -> true
'        既にkeyが存在する    -> false
'   Public Function show(ByVal key As String) As Variant
'    リターン値
'        keyが存在する場合   -> データを渡す
'        keyが存在しない場合 -> Emptyを渡す
'   Public Function del(ByVal key As String) As Boolean
'    リターン値
'        ハッシュの削除成功 -> true
'        keyが存在しない    -> false
'
'実行環境前提：
'   lib_stackクラスが存在すること
'---------------------------------------


'定数
Private Const chunk_size As Long = 50

'ハッシュデータ格納用コンテナ構造体設定
Private Type struct_Container
    value As Variant
    key As String
End Type

'コンテナ定義
Private Container() As struct_Container
Private element As Long

'キーとindex情報の一時保存用バッファ
Private Type struct_map
    key As String
    index As Long
End Type
Private map As struct_map


'コンテナのフリーリスト
Private free As lib_stack

'ハッシュオブジェクト
Private dic As Object



'---------------------------------------
'オブジェクト生成
'---------------------------------------
Private Sub Class_Initialize()

    'VBAハッシュ用初期化
    Set dic = CreateObject("Scripting.Dictionary")

    'freeリストの生成
    Set free = New lib_stack
    
    '初期値設定
    element = 0
    
    'スタック用配列初期確保
    Call extend_Container


End Sub


'---------------------------------------
'オブジェクト破棄
'---------------------------------------
Private Sub class_terminate()

End Sub


'---------------------------------------
'addメソッド(データを追加)
'  add成功　         -> true
'  keyが既に存在する -> false
'---------------------------------------
Public Function add(ByVal key As String, ByVal value As Variant) As Boolean
    
    Dim index As Long
    Dim buf As Variant
    
    '初期化
    index = -1

    'キー確認
    If dic.exists(key) Then
        GoTo Failure
    End If
    
    'データ格納のため空コンテナを確保
    Do
        buf = free.pop
        If IsEmpty(buf) Then
            '空きコンテナがない場合、コンテナを拡張子リトライする
            Call extend_Container
        End If
    Loop While IsEmpty(buf)
    index = CLng(buf)
    
    'Dictonaryにkeyとコンテナのindex番号を登録
    dic.add key, index
    
    '空コンテナへデータ入力
    With Container(index)
        .value = value
        .key = key
    End With
    
Success:
    add = True
    Exit Function
    
Failure:
    'コンテナ取得済みの場合は開放
    If index >= 0 Then
        free.push (index)
    End If
    
    'エラー応答
    add = False
    
End Function


'---------------------------------------
'showメソッド(データを追加)
'  keyと値が存在する場合  -> keyに対応する値を応答
'  keyが存在しない場合    -> Emptyを応答
'---------------------------------------
Public Function show(ByVal key As String) As Variant

    Dim index As Long
    Dim buf As Variant
    
    '初期化
    index = -1

    If map.key = key Then
        index = map.index
    Else
        'キーからindex取得
        buf = dic.Item(key)
        If IsEmpty(buf) Then
            GoTo Failure
        End If
        index = CLng(buf)
    
        With map
            .key = key
            .index = index
        End With
    End If
    
Success:
    show = Container(index).value
    Exit Function
    
Failure:
    show = Empty

End Function


'---------------------------------------
'delメソッド(データを追加)
'  keyが存在し削除成功　-> true
'  keyが存在しない場合  -> false
'---------------------------------------
Public Function del(ByVal key As String) As Boolean

    Dim index As Long
    Dim buf As Variant
    
    '初期化
    index = -1

    'キーからindex取得
    buf = dic.Item(key)
    If IsEmpty(buf) Then
        GoTo Failure
    End If
    index = CLng(buf)
    
    'キーの削除
    dic.Remove (key)
    
    'コンテナの開放
    If index >= 0 Then
        free.push (index)
    Else
        MsgBox "lib_hash:Invalid index in del function. index=" & index
        End
    End If
    
Success:
    del = True
    Exit Function
    
Failure:
    del = False

End Function


'---------------------------------------
'全ハッシュキーのArrayを返す
'---------------------------------------
Public Function GetAllKey() As Variant

    GetAllKey = dic.keys

End Function




'---------------------------------------
'ハッシュ配列の拡張
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


