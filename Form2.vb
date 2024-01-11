Public Class Form2
    Public Const UppLim = 1000000
    Public GroupIndex As Long
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Long
        For i = 1 To Form1.MoveGroup
            ListBox1.Items.Add("第" + Form1.NumLength(i) + "组")
        Next i
        ListBox2.Items.Add("尚未选择任一重复分组")
        ListBox3.Items.Add("各组选择状态")
        CheckedListBox1.Items.Clear()
        For i = 0 To UppLim - 1
            If Len(Form1.Folder_Data(i)) <> 0 Then
                CheckedListBox1.Items.Add(Form1.Folder_Data(i))
            End If
        Next i
        ComboBox1.SelectedIndex = 2
        ComboBox2.SelectedIndex = 2
        ComboBox3.SelectedIndex = 2
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim i As Long, j As Long
        If ListBox1.SelectedIndex > -1 Then
            ListBox3.SelectedIndex = ListBox1.SelectedIndex + 1
        Else
            ListBox3.SelectedIndex = -1
        End If
        GroupIndex = ListBox1.SelectedIndex + 1
        If GroupIndex > 0 Then
            ListBox2.Items.Clear()
            For i = 0 To UppLim - 1
                Form1.Del_Temp(i) = ""
            Next i
            j = 0
            For i = 0 To Form1.MoveCount - 1
                If Form1.FlagTemp(i) = GroupIndex And StatusDetection() = 0 _
                    Or Form1.FlagTemp(i) = GroupIndex And StatusDetection() = 1 And FeFe_Flag(Form1.HandleData(i)) = True _
                    Or Form1.FlagTemp(i) = GroupIndex And StatusDetection() = 2 And FNFe_Flag(Form1.HandleData(i)) = True _
                    Or Form1.FlagTemp(i) = GroupIndex And StatusDetection() = 3 And FeFe_Flag(Form1.HandleData(i)) = True And FNFe_Flag(Form1.HandleData(i)) = True Then
                    Form1.Del_Temp(j) = Form1.HandleData(i)
                    j += 1
                End If
            Next i
            For i = 0 To j - 1
                ListBox2.Items.Add(Form1.Del_Temp(i))
            Next i
        End If
        If ListBox2.Items.Count = 0 Then ListBox2.Items.Add("本组筛选后无重复文件")
    End Sub
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim i As Long
        If ListBox2.SelectedIndex > -1 Then
            ListBox3.Items.Clear()
            ListBox3.Items.Add("当前各组保留文件选择状态")
            Form1.Del_Data(GroupIndex - 1) = ListBox2.SelectedItem
            For i = 0 To Form1.MoveGroup - 1
                If Form1.Del_Data(i) <> "" Then
                    ListBox3.Items.Add("第" + Form1.NumLength(i + 1) + "组已选择文件:" + Form1.Del_Data(i))
                Else
                    ListBox3.Items.Add("第" + Form1.NumLength(i + 1) + "组未选择文件")
                End If
            Next i
            ListBox3.SelectedIndex = ListBox1.SelectedIndex + 1
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Call StatusReset()
        If CheckBox1.Checked = True Then
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
        Else
            CheckBox2.Checked = False
            CheckBox2.Enabled = False
            CheckBox3.Checked = False
            CheckBox3.Enabled = False
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Call StatusReset()
        If CheckBox2.Checked = True Then
            Button1.Enabled = True
            Button2.Enabled = True
            CheckedListBox1.Enabled = True
        Else
            Button1.Enabled = False
            Button2.Enabled = False
            CheckedListBox1.Enabled = False
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Call StatusReset()
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            If ComboBox1.SelectedIndex <> 2 Then TextBox1.Enabled = True
            If ComboBox2.SelectedIndex <> 2 Then TextBox2.Enabled = True
            If ComboBox3.SelectedIndex <> 2 Then TextBox3.Enabled = True
        Else
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Long
        For i = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, Not CheckedListBox1.GetItemChecked(i))
        Next i
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim i As Long
        For i = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, True)
        Next i
    End Sub
    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Call StatusReset()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call StatusReset()
        If ComboBox1.SelectedIndex = 2 Then
            TextBox1.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Call StatusReset()
        If ComboBox2.SelectedIndex = 2 Then
            TextBox2.Enabled = False
        Else
            TextBox2.Enabled = True
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Call StatusReset()
        If ComboBox3.SelectedIndex = 2 Then
            TextBox3.Enabled = False
        Else
            TextBox3.Enabled = True
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim i As Integer, count As Integer
        Dim numindex(260) As Integer
        Dim temp As String
        Call StatusReset()
        If TextBox1.Text.Contains("\"c) Or TextBox1.Text.Contains("/"c) Or TextBox1.Text.Contains(":"c) Or TextBox1.Text.Contains("*"c) Or
            TextBox1.Text.Contains("?"c) Or TextBox1.Text.Contains(""""c) Or TextBox1.Text.Contains("<"c) Or TextBox1.Text.Contains(">"c) Or TextBox1.Text.Contains("|"c) Then
            count = 0
            For i = 1 To Len(TextBox1.Text)
                If Mid(TextBox1.Text, i, 1) = "\" Or Mid(TextBox1.Text, i, 1) = "/" Or Mid(TextBox1.Text, i, 1) = ":" Or Mid(TextBox1.Text, i, 1) = "*" Or
            Mid(TextBox1.Text, i, 1) = "?" Or Mid(TextBox1.Text, i, 1) = "<" Or Mid(TextBox1.Text, i, 1) = ">" Or Mid(TextBox1.Text, i, 1) = """" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = TextBox1.Text
            For i = 0 To count - 1
                If numindex(i) <> 1 And numindex(i) <> Len(TextBox1.Text) Then
                    temp = Mid(TextBox1.Text, 1, numindex(i) - 1) + Mid(TextBox1.Text, numindex(i) + 1)
                ElseIf numindex(i) = 1 Then
                    temp = Mid(TextBox1.Text, numindex(i) + 1)
                ElseIf numindex(i) = Len(TextBox1.Text) Then
                    temp = Mid(TextBox1.Text, 1, numindex(i) - 1)
                End If
            Next i
            TextBox1.Text = temp
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim i As Integer, count As Integer
        Dim numindex(260) As Integer
        Dim temp As String
        Call StatusReset()
        If TextBox2.Text.Contains("\"c) Or TextBox2.Text.Contains("/"c) Or TextBox2.Text.Contains(":"c) Or TextBox2.Text.Contains("*"c) Or
            TextBox2.Text.Contains("?"c) Or TextBox2.Text.Contains(""""c) Or TextBox2.Text.Contains("<"c) Or TextBox2.Text.Contains(">"c) Or TextBox2.Text.Contains("|"c) Then
            count = 0
            For i = 1 To Len(TextBox2.Text)
                If Mid(TextBox2.Text, i, 1) = "\" Or Mid(TextBox2.Text, i, 1) = "/" Or Mid(TextBox2.Text, i, 1) = ":" Or Mid(TextBox2.Text, i, 1) = "*" Or
            Mid(TextBox2.Text, i, 1) = "?" Or Mid(TextBox2.Text, i, 1) = "<" Or Mid(TextBox2.Text, i, 1) = ">" Or Mid(TextBox2.Text, i, 1) = """" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = TextBox2.Text
            For i = 0 To count - 1
                If numindex(i) <> 1 And numindex(i) <> Len(TextBox2.Text) Then
                    temp = Mid(TextBox2.Text, 1, numindex(i) - 1) + Mid(TextBox2.Text, numindex(i) + 1)
                ElseIf numindex(i) = 1 Then
                    temp = Mid(TextBox2.Text, numindex(i) + 1)
                ElseIf numindex(i) = Len(TextBox2.Text) Then
                    temp = Mid(TextBox2.Text, 1, numindex(i) - 1)
                End If
            Next i
            TextBox2.Text = temp
        End If
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim i As Integer, count As Integer
        Dim numindex(260) As Integer
        Dim temp As String
        Call StatusReset()
        If TextBox3.Text.Contains("\"c) Or TextBox3.Text.Contains("/"c) Or TextBox3.Text.Contains(":"c) Or TextBox3.Text.Contains("*"c) Or
            TextBox3.Text.Contains("?"c) Or TextBox3.Text.Contains(""""c) Or TextBox3.Text.Contains("<"c) Or TextBox3.Text.Contains(">"c) Or TextBox3.Text.Contains("|"c) Then
            count = 0
            For i = 1 To Len(TextBox3.Text)
                If Mid(TextBox3.Text, i, 1) = "\" Or Mid(TextBox3.Text, i, 1) = "/" Or Mid(TextBox3.Text, i, 1) = ":" Or Mid(TextBox3.Text, i, 1) = "*" Or
            Mid(TextBox3.Text, i, 1) = "?" Or Mid(TextBox3.Text, i, 1) = "<" Or Mid(TextBox3.Text, i, 1) = ">" Or Mid(TextBox3.Text, i, 1) = """" Then
                    numindex(count) = i
                    count += 1
                End If
            Next i
            temp = TextBox3.Text
            For i = 0 To count - 1
                If numindex(i) <> 1 And numindex(i) <> Len(TextBox3.Text) Then
                    temp = Mid(TextBox3.Text, 1, numindex(i) - 1) + Mid(TextBox3.Text, numindex(i) + 1)
                ElseIf numindex(i) = 1 Then
                    temp = Mid(TextBox3.Text, numindex(i) + 1)
                ElseIf numindex(i) = Len(TextBox3.Text) Then
                    temp = Mid(TextBox3.Text, 1, numindex(i) - 1)
                End If
            Next i
            TextBox3.Text = temp
        End If
    End Sub
    Private Function FeFe_Flag(Str_HandleData As String) As Boolean
        Dim i As Long
        For i = 0 To CheckedListBox1.CheckedItems.Count - 1
            If StrComp(Mid(Form1.FolderPath + Str_HandleData, 1, Len(CheckedListBox1.CheckedItems(i).ToString()) + 1), CheckedListBox1.CheckedItems(i).ToString() + "\") = 0 Then
                Exit For
            End If
        Next i
        If i >= CheckedListBox1.CheckedItems.Count Then
            FeFe_Flag = False
        Else
            FeFe_Flag = True
        End If
    End Function
    Private Function FNFe_Flag(Str_HandleData As String) As Boolean
        Dim i As Long
        Dim FN As String = ""
        Dim TextBox1_Flag As Boolean
        Dim TextBox2_Flag As Boolean
        Dim TextBox3_Flag As Boolean
        For i = Len(Str_HandleData) To 1 Step -1
            If StrComp(Mid(Str_HandleData, i, 1), "\") = 0 Then
                FN = Mid(Str_HandleData, i + 1)
                Exit For
            End If
        Next i
        If i < 1 Then FN = Str_HandleData
        If ComboBox1.SelectedIndex = 0 Then
            TextBox1_Flag = False
            For i = 1 To Len(FN) - Len(TextBox1.Text) + 1
                If StrComp(TextBox1.Text, Mid(FN, i, Len(TextBox1.Text))) = 0 Then
                    TextBox1_Flag = True
                    Exit For
                End If
            Next i
        ElseIf ComboBox1.SelectedIndex = 1 Then
            TextBox1_Flag = True
            For i = 1 To Len(FN) - Len(TextBox1.Text) + 1
                If StrComp(TextBox1.Text, Mid(FN, i, Len(TextBox1.Text))) = 0 Then
                    TextBox1_Flag = False
                    Exit For
                End If
            Next i
        Else
            TextBox1_Flag = True
        End If
        If ComboBox2.SelectedIndex = 0 Then
            TextBox2_Flag = False
            For i = 1 To Len(FN) - Len(TextBox2.Text) + 1
                If StrComp(TextBox2.Text, Mid(FN, i, Len(TextBox2.Text))) = 0 Then
                    TextBox2_Flag = True
                    Exit For
                End If
            Next i
        ElseIf ComboBox2.SelectedIndex = 1 Then
            TextBox2_Flag = True
            For i = 1 To Len(FN) - Len(TextBox2.Text) + 1
                If StrComp(TextBox2.Text, Mid(FN, i, Len(TextBox2.Text))) = 0 Then
                    TextBox2_Flag = False
                    Exit For
                End If
            Next i
        Else
            TextBox2_Flag = True
        End If
        If ComboBox3.SelectedIndex = 0 Then
            TextBox3_Flag = False
            For i = 1 To Len(FN) - Len(TextBox3.Text) + 1
                If StrComp(TextBox3.Text, Mid(FN, i, Len(TextBox3.Text))) = 0 Then
                    TextBox3_Flag = True
                    Exit For
                End If
            Next i
        ElseIf ComboBox3.SelectedIndex = 1 Then
            TextBox3_Flag = True
            For i = 1 To Len(FN) - Len(TextBox3.Text) + 1
                If StrComp(TextBox3.Text, Mid(FN, i, Len(TextBox3.Text))) = 0 Then
                    TextBox3_Flag = False
                    Exit For
                End If
            Next i
        Else
            TextBox3_Flag = True
        End If
        If TextBox1_Flag = True And TextBox2_Flag = True And TextBox3_Flag = True Then
            FNFe_Flag = True
        Else
            FNFe_Flag = False
        End If
    End Function
    Private Function StatusDetection() As Integer
        If CheckBox1.Checked = False Or CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = False Then
            StatusDetection = 0
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = False Then
            StatusDetection = 1
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = True Then
            StatusDetection = 2
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True Then
            StatusDetection = 3
        Else
            StatusDetection = 4
        End If
    End Function
    Private Sub StatusReset()
        Dim i As Long
        ListBox1.SelectedIndex = -1
        For i = 0 To Form1.MoveGroup - 1
            Form1.Del_Data(i) = ""
        Next i
        ListBox2.Items.Clear()
        ListBox2.Items.Add("尚未选择任一重复分组")
        ListBox3.Items.Clear()
        ListBox3.Items.Add("各组选择状态")
        For i = 0 To Form1.MoveGroup - 1
            If Form1.Del_Data(i) <> "" Then
                ListBox3.Items.Add("第" + Form1.NumLength(i + 1) + "组已选择文件:" + Form1.Del_Data(i))
            Else
                ListBox3.Items.Add("第" + Form1.NumLength(i + 1) + "组未选择文件")
            End If
        Next i
    End Sub
    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim i As Long, j As Long
        Dim Result As MsgBoxResult
        For i = 0 To Form1.MoveGroup - 1
            If Form1.Del_Data(i) = "" Or StrComp(Form1.Del_Data(i), "本组筛选后无重复文件") = 0 Then
                For j = 0 To Form1.MoveCount - 1
                    If Form1.FlagTemp(j) = i + 1 And StatusDetection() = 0 _
                    Or Form1.FlagTemp(j) = i + 1 And StatusDetection() = 1 And FeFe_Flag(Form1.HandleData(j)) = True _
                    Or Form1.FlagTemp(j) = i + 1 And StatusDetection() = 2 And FNFe_Flag(Form1.HandleData(j)) = True _
                    Or Form1.FlagTemp(j) = i + 1 And StatusDetection() = 3 And FeFe_Flag(Form1.HandleData(j)) = True And FNFe_Flag(Form1.HandleData(j)) = True Then
                        Form1.Del_Data(i) = Form1.HandleData(j)
                        Exit For
                    Else
                        Form1.Del_Data(i) = "本组筛选后无重复文件"
                    End If
                Next j
            End If
        Next i
        If StatusDetection() = 1 Or StatusDetection() = 3 Then
            Form1.ResFolderCount = CheckedListBox1.CheckedItems.Count
            For i = 0 To UppLim - 1
                Form1.ResFolderList(i) = ""
            Next i
            For i = 0 To Form1.ResFolderCount - 1
                Form1.ResFolderList(i) = CheckedListBox1.CheckedItems(i).ToString()
            Next i
            Form1.FeFeEn_Flag = True
        Else
            Form1.FeFeEn_Flag = False
        End If
        ListBox3.Items.Clear()
        ListBox3.Items.Add(Now + "选择保留文件保存结果如下")
        For i = 0 To Form1.MoveGroup - 1
            ListBox3.Items.Add("第" + Form1.NumLength(i + 1) + "组已选择文件:" + Form1.Del_Data(i))
        Next i
        Form1.ListBox4.Items.Add(Now + "选择保留文件结果保存成功")
        Form1.ListBox4.SelectedItem = Form1.ListBox4.Items(Form1.ListBox4.Items.Count - 1)
        Result = MsgBox("即将退出选择保留文件界面,请在确认保留结果无误后点击<是(Y)>" + vbCrLf + "注意:每次进入选择保留文件界面时会清除上一次的保留结果", vbYesNo + vbExclamation, "警告")
        If Result = vbYes Then
            Form1.Button1.Enabled = True
            Form1.Button2.Enabled = True
            Form1.Button3.Enabled = True
            Form1.Button4.Enabled = True
            Form1.Button5.Enabled = True
            Form1.Button6.Enabled = True
            e.Cancel = False
        Else
            e.Cancel = True
        End If
    End Sub
End Class