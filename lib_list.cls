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
'�T�v�FExcel VBA �X�^�b�N�I�u�W�F�N�g
'�g�����F
'   -�I�u�W�F�N�g����
'        Dim stack as lib_stack
'        Set stack = New lib_stack�I�u�W�F�N�g����
'   -�f�[�^����
'        stack.push( VARIANT_DATA )
'   -�f�[�^���o��
'        VARIANT_DATA = stack.pop
'
'�\���F
'   Public Sub push(ByVal value As Variant)
'    ���^�[���l
'        �Ȃ�
'   Public Function pop() As Variant
'    ���^�[���l
'        STACK�Ƀf�[�^������ꍇ -> �f�[�^��n��
'        STACK�Ƀf�[�^���Ȃ��ꍇ -> Empty��n��
'
'���s���O��F
'   lib_stack�N���X�����݂��邱��
'---------------------------------------


'�萔
Private Const chunk_size As Long = 50
Private Const an_start As Long = -1
Private Const an_end As Long = -2

'�o�������X�g�p�̃R���e�i�\���̐ݒ�
Private Type struct_Container
    value As Variant
    prev As Long
    next As Long
End Type

'�R���e�i��`
Private Container() As struct_Container
Private element As Long

'�o�������X�g�̖��[�̃A���J�[
Private anchor_start As Long
Private anchor_end As Long

'�R���e�i�̃t���[���X�g
Private free As lib_stack


'---------------------------------------
'�I�u�W�F�N�g����
'---------------------------------------
Private Sub Class_Initialize()

    'free���X�g�̐���
    Set free = New lib_stack
    
    '�����l�ݒ�
    element = 0
    anchor_start = an_end
    anchor_end = an_start
    
    '���X�g�p�z�񏉊��m��
    Call extend_Container


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
    
    Dim index As Long
    Dim buf As Variant
    
    '�f�[�^�i�[�̂��ߋ�R���e�i���m��
    Do
        buf = free.pop
        If IsEmpty(buf) Then
            '�󂫃R���e�i���Ȃ��ꍇ�A�R���e�i���g���q���g���C����
            Call extend_Container
        End If
    Loop While IsEmpty(buf)
    index = CLng(buf)
    
    '��R���e�i�փf�[�^����
    With Container(index)
        .value = value
        .next = an_end
        .prev = anchor_end
    End With
    
    '��O�̃R���e�i�ƊJ�n�A���J�[�̃|�C���^�ύX
    If anchor_end < 0 Then
        '���X�g��Ƀf�[�^���Ȃ��ꍇ
        anchor_start = index
    Else
        '���X�g��Ƀf�[�^����ȏ゠��ꍇ
        Container(anchor_end).next = index
    End If
    
    'END�A���J�[�̃|�C���^�ύX
    anchor_end = index
    
End Sub


'---------------------------------------
'pull���\�b�h(���X�g����f�[�^�����o��)
'*���X�g�Ƀf�[�^������ꍇ => �Y���̃f�[�^�����^�[��
'*���X�g�Ƀf�[�^���Ȃ��ꍇ => Empty(Variant)�����^�[��
'---------------------------------------
Public Function pull() As Variant

    Dim first As Long
    Dim second As Long

    '�f�[�^���Ȃ��ꍇ�͏I������
    If anchor_start < 0 Then GoTo Failure

    '�擪�R���e�i����f�[�^��2�Ԗڂ̃R���e�i�|�C���^�𔲂��o��
    first = anchor_start
    With Container(first)
        second = .next
        pull = .value
    End With
    
    '�擪�R���e�i�̊J��
    If second < 0 Then
        '�Ō�̃R���e�i�J��(anchor_end -> anchor_start)
        anchor_end = an_start
    Else
        '���̃R���e�i�����݂���ꍇ
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
'���X�g�z��̊g��
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
