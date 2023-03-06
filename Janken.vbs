Dim te(2)

te(0) = "グー"
te(1) = "チョキ"
te(2) = "パー"

kachi = 0

Randomize

MsgBox "じゃんけんゲーむ"

For i = 1 To 5
    user = CInt(InputBox("0:グー, 1:チョキ ,2:パー"))

    computer = CInt(Rnd * 2)

    s = "ユーザー:" & te(user) & ",コンピューター:" & te(computer)

    If user = computer Then
        MsgBox s & "・・・あいこです"
    ElseIf computer = (user + 1) Mod 3 Then
        MsgBox s & "・・・ユーザーの勝ちです"
        kachi = kachi + 1
    Else
        MsgBox s & "・・・コンピューターの勝ちです"
    End If
Next

MsgBox "ユーザーの勝ち数:" & kachi
    