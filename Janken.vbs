Dim te(2)

te(0) = "�O�["
te(1) = "�`���L"
te(2) = "�p�["

kachi = 0

Randomize

MsgBox "����񂯂�Q�[��"

For i = 1 To 5
    user = CInt(InputBox("0:�O�[, 1:�`���L ,2:�p�["))

    computer = CInt(Rnd * 2)

    s = "���[�U�[:" & te(user) & ",�R���s���[�^�[:" & te(computer)

    If user = computer Then
        MsgBox s & "�E�E�E�������ł�"
    ElseIf computer = (user + 1) Mod 3 Then
        MsgBox s & "�E�E�E���[�U�[�̏����ł�"
        kachi = kachi + 1
    Else
        MsgBox s & "�E�E�E�R���s���[�^�[�̏����ł�"
    End If
Next

MsgBox "���[�U�[�̏�����:" & kachi
    