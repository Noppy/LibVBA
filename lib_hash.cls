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
'�T�v�FExcel VBA �n�b�V���I�u�W�F�N�g
'�g�����F
'   -�I�u�W�F�N�g����
'        Dim hash as lib_hash
'        Set hash = New lib_hash�I�u�W�F�N�g����
'   -�f�[�^����
'        hash.add( HASH_KEY, VARIANT_DATA )
'   -�f�[�^�Q��
'        VARIANT_DATA = hash.show(HASH_KEY)
'   -�f�[�^�폜
'        hash.del(HASH_KEY)
'
'�\���F
'   Public Function add(ByVal key As String, ByVal value As Variant) As Boolean
'    ���^�[���l
'        �n�b�V���ւ̒ǉ����� -> true
'        ����key�����݂���    -> false
'   Public Function show(ByVal key As String) As Variant
'    ���^�[���l
'        key�����݂���ꍇ   -> �f�[�^��n��
'        key�����݂��Ȃ��ꍇ -> Empty��n��
'   Public Function del(ByVal key As String) As Boolean
'    ���^�[���l
'        �n�b�V���̍폜���� -> true
'        key�����݂��Ȃ�    -> false
'
'���s���O��F
'   lib_stack�N���X�����݂��邱��
'---------------------------------------


'�萔
Private Const chunk_size As Long = 50

'�n�b�V���f�[�^�i�[�p�R���e�i�\���̐ݒ�
Private Type struct_Container
    value As Variant
    key As String
End Type

'�R���e�i��`
Private Container() As struct_Container
Private element As Long

'�L�[��index���̈ꎞ�ۑ��p�o�b�t�@
Private Type struct_map
    key As String
    index As Long
End Type
Private map As struct_map


'�R���e�i�̃t���[���X�g
Private free As lib_stack

'�n�b�V���I�u�W�F�N�g
Private dic As Object



'---------------------------------------
'�I�u�W�F�N�g����
'---------------------------------------
Private Sub Class_Initialize()

    'VBA�n�b�V���p������
    Set dic = CreateObject("Scripting.Dictionary")

    'free���X�g�̐���
    Set free = New lib_stack
    
    '�����l�ݒ�
    element = 0
    
    '�X�^�b�N�p�z�񏉊��m��
    Call extend_Container


End Sub


'---------------------------------------
'�I�u�W�F�N�g�j��
'---------------------------------------
Private Sub class_terminate()

End Sub


'---------------------------------------
'add���\�b�h(�f�[�^��ǉ�)
'  add�����@         -> true
'  key�����ɑ��݂��� -> false
'---------------------------------------
Public Function add(ByVal key As String, ByVal value As Variant) As Boolean
    
    Dim index As Long
    Dim buf As Variant
    
    '������
    index = -1

    '�L�[�m�F
    If dic.exists(key) Then
        GoTo Failure
    End If
    
    '�f�[�^�i�[�̂��ߋ�R���e�i���m��
    Do
        buf = free.pop
        If IsEmpty(buf) Then
            '�󂫃R���e�i���Ȃ��ꍇ�A�R���e�i���g���q���g���C����
            Call extend_Container
        End If
    Loop While IsEmpty(buf)
    index = CLng(buf)
    
    'Dictonary��key�ƃR���e�i��index�ԍ���o�^
    dic.add key, index
    
    '��R���e�i�փf�[�^����
    With Container(index)
        .value = value
        .key = key
    End With
    
Success:
    add = True
    Exit Function
    
Failure:
    '�R���e�i�擾�ς݂̏ꍇ�͊J��
    If index >= 0 Then
        free.push (index)
    End If
    
    '�G���[����
    add = False
    
End Function


'---------------------------------------
'show���\�b�h(�f�[�^��ǉ�)
'  key�ƒl�����݂���ꍇ  -> key�ɑΉ�����l������
'  key�����݂��Ȃ��ꍇ    -> Empty������
'---------------------------------------
Public Function show(ByVal key As String) As Variant

    Dim index As Long
    Dim buf As Variant
    
    '������
    index = -1

    If map.key = key Then
        index = map.index
    Else
        '�L�[����index�擾
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
'del���\�b�h(�f�[�^��ǉ�)
'  key�����݂��폜�����@-> true
'  key�����݂��Ȃ��ꍇ  -> false
'---------------------------------------
Public Function del(ByVal key As String) As Boolean

    Dim index As Long
    Dim buf As Variant
    
    '������
    index = -1

    '�L�[����index�擾
    buf = dic.Item(key)
    If IsEmpty(buf) Then
        GoTo Failure
    End If
    index = CLng(buf)
    
    '�L�[�̍폜
    dic.Remove (key)
    
    '�R���e�i�̊J��
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
'�S�n�b�V���L�[��Array��Ԃ�
'---------------------------------------
Public Function GetAllKey() As Variant

    GetAllKey = dic.keys

End Function




'---------------------------------------
'�n�b�V���z��̊g��
'---------------------------------------
Private Sub extend_Container()

    Dim old_element As Long
    Dim i As Long
    
    '���݂̔z��C���f�b�N�X�̕ۑ�
    old_element = element
    If old_element <= 0 Then old_element = -1

    '�z��̊g��
    ReDim Preserve Container(element + chunk_size)
    element = UBound(Container)
    
    '�g�������R���e�i��free���X�g�֓o�^
    For i = old_element + 1 To element
        free.push (i)
    Next
    
End Sub

