Public Class Form1
    Dim Card(7, 5) As PictureBox '牌組的位子(插槽)
    Dim CardNumber(7, 5), PictureCk(52), PlayPoint(7), Win, Lose, Tie As Integer 'CardNumber(卡片的數字(k=13) 方便注入圖片用) PictureCk(判斷卡片出現過沒有)
    Dim N As String '玩家數量
    Dim Picture(52) As Image 'Picture讀取所有撲克牌的圖片
    Dim Play(7) As GroupBox '玩家
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            For i = 0 To 52 : Picture(i) = Image.FromFile(i & ".gif") : Next '讀取所有圖片檔案
            For i = 1 To 7 : Play(i) = Me.Controls("GroupBox" & i) : Next 'GroupBox
            For i As Integer = 1 To N : Play(i).Visible = True : Next '幾個玩家就顯示幾個介面
            For i As Integer = 1 To 7
                For x As Integer = 1 To 5
                    Card(i, x) = Me.Play(i).Controls("PictureBox" & (i - 1) * 5 + x)
                    Card(i, x).Image = Picture(0) '蓋牌
                Next
            Next
            Re()
        Catch ex As Exception

        End Try
    End Sub

    Sub CardPoint(ByVal Now As Integer) '看牌組的總點數多少
        Dim cnt, sum, P As Integer
        Do Until CardNumber(Now, cnt + 1) = 0
            If cnt + 1 < 5 Then cnt += 1
        Loop
        For i As Integer = 1 To cnt
            P = CardNumber(Now, i)
            If P > 13 Then
                If P Mod 13 > 10 Then '如果是 J Q K都算10點
                    sum += 10
                ElseIf P = 1 Then '如果是ㄟ系(A) 看當分數+10有沒有超過21了 有的話就算1 沒有的話則算10
                    If sum + 10 > 21 Then sum += 1 Else sum += 10
                Else : sum += P Mod 13
                End If
            Else
                If P > 10 Then '如果是 J Q K都算10點
                    sum += 10
                ElseIf P = 1 Then '如果是ㄟ系(A) 看當分數+10有沒有超過21了 有的話就算1 沒有的話則算10
                    If sum + 10 > 21 Then sum += 1 Else sum += 10
                Else : sum += P Mod 13
                End If
            End If
        Next
        PlayPoint(Now) = sum : Out()
    End Sub

    Sub Out() '出局
        For i As Integer = 1 To N
            If PlayPoint(i) = 21 Then
                Play(i).Text = "Win"
                Me.Play(i).Controls("Button" & i).Enabled = False
            ElseIf PlayPoint(i) > 21 Then
                Play(i).Text = "Lose"
                Me.Play(i).Controls("Button" & i).Enabled = False
            End If
        Next
    End Sub

    Sub GiveCard(ByVal Now As Integer) '發牌
        Try
            Randomize()
            Dim cnt, Point As Integer
            Do
                cnt += 1
            Loop Until CardNumber(Now, cnt) = 0 '現在有幾張牌了 檢測
            Do
                Point = Int((Rnd() * 52) + 1)
            Loop Until PictureCk(Point) = 0 '隨機發牌 不能重複出現 出現過的PictureCK=1
            PictureCk(Point) = 1
            CardNumber(Now, cnt) = Point
            Card(Now, cnt).Image = Picture(Point)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click '玩家1補牌
        GiveCard(2) : CardPoint(2)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click '玩家2補牌
        GiveCard(3) : CardPoint(3)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click '玩家3補牌
        GiveCard(4) : CardPoint(4)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click '玩家4補牌
        GiveCard(5) : CardPoint(5)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click '玩家5補牌
        GiveCard(6) : CardPoint(6)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click '玩家6補牌
        GiveCard(7) : CardPoint(7)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click '莊家開牌
        If MsgBox("是否要開牌，開牌後其他玩家都無法再進行補牌了" & vbNewLine & "Ps:點數未滿17點將會自動補牌", 36) = vbYes Then
            Card(1, 2).Image = Picture(CardNumber(1, 2))
            Button1.Enabled = True : Button8.Enabled = False
            CardPoint(1) : WhoWin()
            Do Until PlayPoint(1) >= 17
                GiveCard(1) : CardPoint(1) : WhoWin()
            Loop
            WinLoseTieFr() '當莊家點數大於17就關閉補牌功能
            For i As Integer = 2 To 7
                Me.Play(i).Controls("Button" & i).Enabled = False
            Next
        End If
    End Sub

    Sub WhoWin() '誰贏了 跟莊家比
        If Play(1).Text = "Lose" Then
            For i As Integer = 2 To 7
                Play(i).Text = "Win"
            Next
        Else
            For i As Integer = 2 To 7
                If Play(i).Text <> "Lose" Then
                    If PlayPoint(1) = PlayPoint(i) Then
                        Play(i).Text = "平手"
                    ElseIf PlayPoint(1) > PlayPoint(i) Then
                        Play(i).Text = "Lose"
                    ElseIf PlayPoint(1) < PlayPoint(i) Then
                        Play(i).Text = "Win"
                    End If
                End If
            Next
        End If
    End Sub

    Sub WinLoseTieFr() '莊家勝利 戰敗 平手 次數
        For i As Integer = 2 To N
            If Play(i).Text = "Lose" Then
                Win += 1
            ElseIf Play(i).Text = "Win" Then
                Lose += 1
            Else : Tie += 1
            End If
        Next
    End Sub

    Private Sub 重新開始ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 重新開始ToolStripMenuItem.Click
        Re()
    End Sub

    Sub Re() '全部初始化 重新開始
        Try
            Do
                N = InputBox("請輸入有幾個玩家(1~6個)", , 3) '有幾個玩家
                If N < 1 Or N > 6 Or Not IsNumeric(N) Then MsgBox("數字請介於1~6") : Close() Else N += 1 : Exit Do
            Loop
            ReDim PlayPoint(7) : ReDim CardNumber(7, 5) : ReDim PictureCk(52)
            For i As Integer = 1 To 7
                For x As Integer = 1 To 5
                    Card(i, x).Image = Picture(0) '蓋牌
                Next
            Next
            For i As Integer = 3 To 7
                Play(i).Visible = False
            Next
            For i As Integer = 1 To N  '每個玩家都要發牌
                For F As Integer = 1 To 2 '發兩張
                    GiveCard(i) : CardPoint(i)
                    If i = 1 And F = 2 Then Card(1, 2).Image = Picture(0) '莊家一開局 覆蓋一張卡
                Next
                Play(i).Visible = True
                If i <> 1 Then Me.Play(i).Controls("Button" & i).Enabled = True : Play(i).Text = "玩家" & i - 1 Else Button8.Enabled = True : Play(i).Text = "莊家"
            Next
        Catch ex As Exception
            Close()
        End Try
    End Sub

    Private Sub 說明ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 說明ToolStripMenuItem.Click
        MsgBox("玩家點數高於莊家時獲勝，或者是莊家爆牌，其點數必須等於或低於21點；超過21點的玩家稱為爆牌。2點至10點的牌以牌面的點數計算，J、Q、K 每張為10點。A可記為1點或為11點，若玩家會因A而爆牌則A可算為1 點。")
    End Sub

    Private Sub 結束ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 結束ToolStripMenuItem.Click
        If Win <> 0 Or Lose <> 0 Or Tie <> 0 Then MsgBox("莊家勝利次數" & Win & " 戰敗次數" & Lose & " 平手次數" & Tie)
        Close()
    End Sub
End Class
