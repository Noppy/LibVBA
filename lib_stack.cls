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
'�T�v�FExcel VBA �X�^�b�N�I�u�W�F�N�g
'�g�����F
'   -�I�u�W�F�N�g����
'        Dim stack as lib_stack
'        Set stack = New lib_stack�I�u�W�F�N�g����
'   -�f�[�^����
'        stack.push( VARIANT_DATA )
'   -�f�[�^���o��
'        VARIANT_DATA = stack.pop
'�\���F
'   Public Sub push(ByVal value As Variant)
'    ���^�[���l
'        �Ȃ�
'   Public Function pop() As Variant
'    ���^�[���l
'        STACK�Ƀf�[�^������ꍇ -> �f�[�^��n��
'        STACK�Ƀf�[�^���Ȃ��ꍇ -> Empty��n��
'---------------------------------------


Private Const chunk_size As Long = 50

Private data() As Variant '�X�^�b�N�f�[�^�i�[�p�z��
Private element As Long   '�X�^�b�N�z��v�f����
Private depth As Long     '�X�^�b�N���݂̐[��


'---------------------------------------
'�I�u�W�F�N�g����
'---------------------------------------
Private Sub Class_Initialize()

    '�����l�ݒ�
    element = 0
    depth = 0
    
    '�X�^�b�N�p�z�񏉊��m��
    Call extend_chunk

End Sub


'---------------------------------------
'�I�u�W�F�N�g�j��
'---------------------------------------
Private Sub class_terminate()

End Sub


'---------------------------------------
'push���\�b�h(�f�[�^���X�^�b�N�ɒǉ�)
'---------------------------------------
Public Sub push(ByVal value As Variant)
    
    '�f�[�^�̊i�[
    data(depth) = value
    
    '���݂̃X�^�b�N�̐[�����C���N�������g����
    depth = depth + 1
    
    '�X�^�b�N�f�[�^�z������ׂė��p���Ă���ꍇ�͔z����g������
    If depth > element Then
        Call extend_chunk
    End If

End Sub


'---------------------------------------
'pop���\�b�h(�X�^�b�N����f�[�^�����o��)
'*�X�^�b�N�Ƀf�[�^������ꍇ => �Y���̃f�[�^�����^�[��
'*�X�^�b�N�Ƀf�[�^���Ȃ��ꍇ => Empty(Variant)�����^�[��
'---------------------------------------
Public Function pop() As Variant

    '�f�[�^���Ȃ��ꍇ�͏I������
    If depth <= 0 Then GoTo Failure

    '�z��̍Ō�̃f�[�^�����^�[���l�ɐݒ�
    pop = data(depth - 1)
    
    '���݂̃X�^�b�N�̐[�����f�N�������g
    depth = depth - 1
    
    '�z����k��
    '(�f�[�^�z��2�`�����N�ȏ゠���Ă���ꍇ)
    If depth < element - 2 * chunk_size Then
        Call reduce_chunk
    End If


Success:
    Exit Function
    
Failure:
    Exit Function

End Function

'---------------------------------------
'�X�^�b�N�z��̊g��
'---------------------------------------
Private Sub extend_chunk()

    '�z��̊g��
    ReDim Preserve data(element + chunk_size)
    element = UBound(data)
    
End Sub


'---------------------------------------
'�X�^�b�N�z��̏k��
'---------------------------------------
Private Sub reduce_chunk()

    Dim i As Long
    i = element - chunk_size
    
    '��Ԋm�F
    '(�Ō�̃`�����N�܂��́A���p����element������ꍇ�͏k���𒆒f����)
    If i < chunk_size Or i < depth Then
        GoTo Failuer
    End If

    '�z��̏k��
    ReDim Preserve data(i)
    element = UBound(data)

Success:
    Exit Sub
    
Failuer:
    Exit Sub

End Sub