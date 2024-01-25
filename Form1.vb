Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Public Class Form1
    <DllImport("Everything64.dll", EntryPoint:="Everything_SetSearchW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_SetSearchW(search As String) As UInt32
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_SetRequestFlags", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_SetRequestFlags(dwRequestFlags As UInt32) As UInt32
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_QueryW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_QueryW(bWait As Integer) As Integer
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_GetNumResults", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_GetNumResults() As UInt32
    End Function
    <DllImport("Everything64.dll", EntryPoint:="Everything_GetResultFullPathNameW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function Everything_GetResultFullPathNameW(index As UInt32, buf As StringBuilder, size As UInt32) As UInt32
    End Function
    Public Const UppLim = 1000000
    Public Const EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME = &H4
    Public FolderPath As String
    Public RawData() As String = New String(UppLim - 1) {}
    Public RawData_Backup() As String = New String(UppLim - 1) {}
    Public HandleData() As String = New String(UppLim - 1) {}
    Public Hash_Data() As String = New String(UppLim - 1) {}
    Public Del_Data() As String = New String(UppLim - 1) {}
    Public Folder_Data() As String = New String(UppLim - 1) {}
    Public Del_Temp() As String = New String(UppLim - 1) {}
    Public ResFolderList() As String = New String(UppLim - 1) {}
    Public FlagTemp() As Long = New Long(UppLim - 1) {}
    Public RawCount As Long, MoveGroup As Long, MoveCount As Long, ResFolderCount As Long
    Public FeFeEn_Flag As Boolean
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FolderPath = ""
        RawCount = 0
        ComboBox1.SelectedIndex = 0
        ListBox4.Items.Add(Now + "请选择查重文件夹路径")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            FolderPath = FolderBrowserDialog1.SelectedPath
            ListBox4.Items.Add(Now + "已选择查重文件夹路径:" + FolderPath)
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            Button2.Enabled = True
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            Button7.Enabled = False
            ComboBox1.Enabled = False
        Else
            ListBox4.Items.Add(Now + "已取消查重文件夹路径选择")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim AppPath As String, FolderPathTemp As String = "", Comp_1 As String, Comp_2 As String
        Dim Filename As New StringBuilder(260)
        Dim FolderCheck As Boolean = True
        Dim NumResults As Long
        Dim i As Long, j As Long
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        ComboBox1.Enabled = False
        Form2.Close()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Add(Now + "正在确认文件夹路径是否有误")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If Boo_DirExist(FolderPath) = 0 Then
            ListBox4.Items.Add(Now + "文件夹路径错误,程序已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:未找到该文件夹路径,请重新指定", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
            Exit Sub
        End If
        AppPath = Directory.GetCurrentDirectory()
        If Mid(FolderPath, Len(FolderPath), 1) <> "\" Then
            FolderPath += "\"
        End If
        If Mid(AppPath, Len(AppPath), 1) <> "\" Then
            AppPath += "\"
        End If
        Comp_1 = ""
        Comp_2 = ""
        For i = 1 To Len(FolderPath)
            If "a" <= Mid(FolderPath, i, 1) And Mid(FolderPath, i, 1) <= "z" Then
                Comp_1 += Chr(Asc(Mid(FolderPath, i, 1)) - 32)
            Else
                Comp_1 += Mid(FolderPath, i, 1)
            End If
        Next i
        For i = 1 To Len(AppPath)
            If "a" <= Mid(AppPath, i, 1) And Mid(AppPath, i, 1) <= "z" Then
                Comp_2 += Chr(Asc(Mid(AppPath, i, 1)) - 32)
            Else
                Comp_2 += Mid(AppPath, i, 1)
            End If
        Next i
        If Len(Comp_1) <= Len(Comp_2) Then
            If StrComp(Comp_1, Mid(Comp_2, 1, Len(Comp_1))) = 0 Then FolderCheck = False
        End If
        If FolderCheck = False Then
            ListBox4.Items.Add(Now + "文件夹路径错误,程序已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:文件夹路径不能嵌套程序所在文件夹或与程序文件夹相同", 0 + vbCritical + vbSystemModal, "文件夹路径错误")
            Exit Sub
        End If
        For i = 1 To Len(Comp_1)
            If StrComp(Mid(Comp_1, i, 1), " ") = 0 Then
                FolderCheck = False
                Exit For
            End If
        Next i
        If FolderCheck = False Then
            For j = i - 1 To 1 Step -1
                If StrComp(Mid(FolderPath, j, 1), "\") = 0 Then
                    FolderPathTemp = Mid(FolderPath, 1, j)
                    Exit For
                End If
            Next j
        Else
            FolderPathTemp = FolderPath
        End If
        ListBox4.Items.Add(Now + "文件夹路径无误,开始枚举文件夹下的全部文件")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Everything_SetSearchW(FolderPathTemp)
        Everything_SetRequestFlags(EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME)
        Everything_QueryW(1)
        NumResults = Everything_GetNumResults()
        RawCount = 0
        For i = 0 To NumResults - 1
            Everything_GetResultFullPathNameW(i, Filename, Filename.Capacity)
            If File.Exists(Filename.ToString) = True Then
                RawData(RawCount) = Filename.ToString
                If FolderCheck = True Or FolderCheck = False And StrComp(FolderPath, Mid(RawData(RawCount), 1, Len(FolderPath))) = 0 Then
                    RawCount += 1
                End If
                If RawCount = UppLim Then
                    ListBox4.Items.Add(Now + "文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    MsgBox("文件导入个数大于一百万,程序已停止导入,后续操作仅面向已列表的文件", 0 + vbExclamation, "警告")
                    Exit For
                End If
            End If
        Next i
        For i = 0 To RawCount - 1
            RawData(i) = Mid(RawData(i), Len(FolderPath) + 1)
        Next i
        ListBox4.Items.Add(Now + "文件夹枚举完成,开始输出文件列表")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        ListBox1.Items.Add(Now + "|文件列表如下:")
        For i = 0 To RawCount - 1
            ListBox1.Items.Add(NumLength(Trim(Str(i + 1))) + "|" + RawData(i))
        Next i
        ListBox4.Items.Add(Now + "文件列表输出完成,导入文件总个数为:" + CStr(RawCount))
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If RawCount = 0 Then
            ListBox1.Items.Add("该文件夹下没有文件")
            MsgBox("该文件夹下没有文件", 0 + vbInformation + vbSystemModal, "文件查重")
        Else
            Button3.Enabled = True
            ComboBox1.Enabled = True
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Long, j As Long, num_1 As Long, num_2 As Long, temp As Long, tmp As Long, HandleCount As Long
        Dim shatmp As String
        Dim Result As MsgBoxResult
        Dim RepCheck As Boolean, IgnFlag As Boolean
        Dim Flag() As Long = New Long(UppLim - 1) {}
        Dim SearchData() As Long = New Long(UppLim - 1) {}
        Dim SearchDataFin() As Long = New Long(UppLim - 1) {}
        Dim DupCheck_1() As Long = New Long(UppLim - 1) {}
        Dim DupCheck_2() As Long = New Long(UppLim - 1) {}
        Dim Output() As String = New String(UppLim - 1) {}
        Dim HashValueCalculationArray() As HashValueCalculation = New HashValueCalculation(UppLim - 1) {}
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        ComboBox1.Enabled = False
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Add(Now + "开始获取全部列表文件大小")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        ListBox2.Items.Add(Now + "|第一轮查重对比数据(文件大小)列表如下:")
        IgnFlag = False
        For i = 0 To RawCount - 1
            Dim SourceFileInfo As New FileInfo(FolderPath + RawData(i))
            If File.Exists(FolderPath + RawData(i)) = False Then
                Output(i) = "Error"
                ListBox4.Items.Add(Now + "错误:文件" + RawData(i) + "文件路径获取失败,文件不存在")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                If IgnFlag = False Then
                    MsgBox("错误:文件" + RawData(i) + vbCrLf + "文件路径获取失败,文件不存在", 0 + vbCritical + vbSystemModal, "文件路径获取失败")
                    ListBox4.Items.Add(Now + "已发生获取失败,开始询问是否继续执行查重操作")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    Result = MsgBox("已发生获取失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来所有的获取失败警告", 0 + vbExclamation + vbYesNo, "警告")
                    If Result = vbYes Then
                        IgnFlag = True
                        ListBox4.Items.Add(Now + "已选择是,将继续执行并忽略获取失败警告")
                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    Else
                        ListBox4.Items.Add(Now + "已选择否,获取全部列表文件大小操作已终止")
                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                        Exit Sub
                    End If
                End If
            Else
                Try
                    DupCheck_1(i) = SourceFileInfo.Length
                    Output(i) = CStr(DupCheck_1(i))
                Catch ex As Exception
                    Output(i) = "Error"
                    ListBox4.Items.Add(Now + "错误:文件" + RawData(i) + "程序没有读取权限,无法获取文件大小")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                End Try
            End If
            ListBox2.Items.Add(NumLength(i + 1) + "|" + Output(i))
        Next i
        ListBox4.Items.Add(Now + "全部列表文件大小获取完成,开始进行第一轮查重(文件大小)")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        For i = 0 To UppLim - 1
            Flag(i) = 0
            FlagTemp(i) = 0
            DupCheck_2(i) = 0
            SearchData(i) = 0
            SearchDataFin(i) = 0
            Output(i) = ""
        Next i
        num_2 = 0
        temp = 0
        tmp = 0
        HandleCount = 0
        shatmp = ""
        num_1 = 1
        RepCheck = False
        For i = 0 To RawCount - 2
            If Flag(i) = 0 Then
                For j = i + 1 To RawCount - 1
                    If DupCheck_1(i) = DupCheck_1(j) Then
                        If RepCheck = False Then
                            Flag(i) = num_1
                            RepCheck = True
                        End If
                        Flag(j) = num_1
                    End If
                Next j
                If RepCheck = True Then
                    num_1 += 1
                    RepCheck = False
                End If
            End If
        Next i
        j = 0
        For i = 0 To RawCount - 1
            If Flag(i) <> 0 Then
                HandleData(j) = RawData(i)
                DupCheck_2(j) = DupCheck_1(i)
                FlagTemp(j) = Flag(i)
                SearchData(j) = i
                j += 1
            End If
        Next i
        RepCheck = False
        num_2 = j - 1
        ListBox2.Items.Add("-----------------------------------------")
        If num_2 > 0 Then
            ListBox4.Items.Add(Now + "第一轮查重(文件大小)已完成,开始输出重复列表")
            ListBox4.Items.Add(Now + "重复文件组数为:" + CStr(num_1 - 1) + "|重复文件总数为:" + CStr(num_2 + 1))
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            For i = 1 To num_1 - 1
                For j = 0 To num_2
                    If i = FlagTemp(j) Then
                        If RepCheck = False Then
                            temp = 1
                            tmp = DupCheck_2(j)
                            ListBox2.Items.Add("第" + CStr(i) + "组重复文件文件名为:")
                            RepCheck = True
                        End If
                        ListBox2.Items.Add(NumLength(temp) + "|" + NumLength(SearchData(j) + 1) + "|" + HandleData(j))
                        temp += 1
                    End If
                Next j
                ListBox2.Items.Add("以上重复项的文件大小为:" + CStr(tmp) + "字节")
                If i <> num_1 - 1 Then ListBox2.Items.Add("")
                RepCheck = False
            Next i
            ListBox4.Items.Add(Now + "第一轮查重(文件大小)重复列表已输出完成,开始计算第一轮重复文件的SHA1校验值")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            HandleCount = num_2
            For i = 0 To UppLim - 1
                Hash_Data(i) = ""
                DupCheck_1(i) = 0
                DupCheck_2(i) = 0
                Flag(i) = 0
                FlagTemp(i) = 0
            Next i
            num_2 = 0
            temp = 0
            tmp = 0
            shatmp = 0
            If ComboBox1.SelectedIndex = 0 Then
                For i = 0 To HandleCount
                    If File.Exists(FolderPath + HandleData(i)) = False Then
                        Hash_Data(i) = "Error_" + CStr(i + 1)
                        ListBox4.Items.Add(Now + "错误:文件" + HandleData(i) + "文件路径获取失败,文件不存在")
                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    Else
                        Try
                            Dim HashValueCalculationThread As New HashValueCalculation With {
                                .sn = i,
                                .sn_filestream = File.Open(FolderPath + HandleData(i), FileMode.Open, FileAccess.Read, FileShare.Read)
                            }
                            HashValueCalculationArray(i) = HashValueCalculationThread
                            Call HashValueCalculationThread.ThreadWork()
                        Catch ex As Exception
                            Hash_Data(i) = "Error_" + CStr(i + 1)
                            ListBox4.Items.Add(Now + "错误:文件" + HandleData(i) + "程序没有读取权限,无法计算SHA1校验值")
                            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                        End Try
                    End If
                Next i
            ElseIf ComboBox1.SelectedIndex = 1 Then
                Dim ThreadState() As Task = New Task(HandleCount) {}
                For i = 0 To HandleCount
                    If File.Exists(FolderPath + HandleData(i)) = False Then
                        Hash_Data(i) = "Error_" + CStr(i + 1)
                        ListBox4.Items.Add(Now + "错误:文件" + HandleData(i) + "文件路径获取失败,文件不存在")
                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    Else
                        Try
                            Dim HashValueCalculationThread As New HashValueCalculation With {
                                .sn = i,
                                .sn_filestream = File.Open(FolderPath + HandleData(i), FileMode.Open, FileAccess.Read, FileShare.Read)
                            }
                            HashValueCalculationArray(i) = HashValueCalculationThread
                            ThreadState(i) = Task.Run(AddressOf HashValueCalculationThread.ThreadWork)
                        Catch ex As Exception
                            Hash_Data(i) = "Error_" + CStr(i + 1)
                            ListBox4.Items.Add(Now + "错误:文件" + HandleData(i) + "程序没有读取权限,无法计算SHA1校验值")
                            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                        End Try
                    End If
                Next i
                Try
                    Task.WaitAll(ThreadState)
                Catch ex As Exception
                    ListBox4.Items.Add(Now + "错误:等待全部线程完成失败,程序中止")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    MsgBox("错误:等待全部线程完成失败,程序中止", 0 + vbCritical + vbSystemModal, "等待全部线程完成失败")
                    Exit Sub
                End Try
            Else
                ListBox4.Items.Add(Now + "错误:选择SHA1校验值计算模式错误,程序中止")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MsgBox("错误:选择SHA1校验值计算模式错误,程序中止", 0 + vbCritical + vbSystemModal, "选择SHA1校验值计算模式错误")
                Exit Sub
            End If
            ListBox3.Items.Add(Now + "|第二轮查重对比数据(SHA1校验值)列表如下:")
            For i = 0 To HandleCount
                Hash_Data(i) = HashValueCalculationArray(i).sn_hashresult
                ListBox3.Items.Add(NumLength(i + 1) + "|" + NumLength(SearchData(i) + 1) + "|" + Hash_Data(i))
            Next i
            For i = 0 To UppLim - 1
                RawData_Backup(i) = RawData(i)
                RawData(i) = ""
            Next i
            For i = 0 To UppLim - 1
                RawData(i) = HandleData(i)
                HandleData(i) = ""
            Next i
            ListBox4.Items.Add(Now + "第一轮重复文件的SHA1校验值计算完成,开始第二轮查重(SHA1校验值)")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            num_1 = 1
            RepCheck = False
            For i = 0 To HandleCount - 1
                If Flag(i) = 0 Then
                    For j = i + 1 To HandleCount
                        If StrComp(Hash_Data(i), Hash_Data(j)) = 0 Then
                            If RepCheck = False Then
                                Flag(i) = num_1
                                RepCheck = True
                            End If
                            Flag(j) = num_1
                        End If
                    Next j
                    If RepCheck = True Then
                        num_1 += 1
                        RepCheck = False
                    End If
                End If
            Next i
            j = 0
            For i = 0 To HandleCount
                If Flag(i) <> 0 Then
                    HandleData(j) = RawData(i)
                    Output(j) = Hash_Data(i)
                    FlagTemp(j) = Flag(i)
                    SearchDataFin(j) = SearchData(i)
                    Flag(i) = 0
                    j += 1
                End If
            Next i
            RepCheck = False
            num_2 = j - 1
            j = 0
            ListBox3.Items.Add("-----------------------------------------")
            If num_2 > 0 Then
                ListBox4.Items.Add(Now + "第二轮查重(SHA1校验值)已完成,开始输出重复列表")
                ListBox4.Items.Add(Now + "重复文件组数为:" + CStr(num_1 - 1) + "|重复文件总数为:" + CStr(num_2 + 1))
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                For i = 1 To num_1 - 1
                    For j = 0 To num_2
                        If i = FlagTemp(j) Then
                            If RepCheck = False Then
                                temp = 1
                                shatmp = Output(j)
                                ListBox3.Items.Add("第" + CStr(i) + "组重复文件文件名为:")
                                RepCheck = True
                            End If
                            ListBox3.Items.Add(NumLength(temp) + "|" + NumLength(SearchDataFin(j) + 1) + "|" + HandleData(j))
                            temp += 1
                        End If
                    Next j
                    ListBox3.Items.Add("以上重复项的SHA1校验值为:" + shatmp)
                    If i <> num_1 - 1 Then ListBox3.Items.Add("")
                    RepCheck = False
                Next i
                For i = 0 To UppLim - 1
                    Folder_Data(i) = ""
                Next i
                For i = 0 To RawCount - 1
                    FolderList(FolderPath + HandleData(i))
                Next i
                ListBox4.Items.Add(Now + "第二轮查重(SHA1校验值)重复列表已输出完成,请查验查重结果")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                Button4.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = False
                Button7.Enabled = True
                ComboBox1.Enabled = True
            Else
                ListBox3.Items.Add("未找到重复文件")
                ListBox4.Items.Add(Now + "未找到重复文件")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MsgBox("未找到重复文件", 0 + vbInformation + vbSystemModal, "查重")
                Button3.Enabled = False
            End If
        Else
            ListBox2.Items.Add("未找到重复文件")
            ListBox4.Items.Add(Now + "未找到重复文件")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("未找到重复文件", 0 + vbInformation + vbSystemModal, "查重")
            Button3.Enabled = False
        End If
        For i = 0 To UppLim - 1
            Hash_Data(i) = ""
            Output(i) = ""
            Del_Data(i) = ""
            RawData(i) = RawData_Backup(i)
            SearchData(i) = 0
            SearchDataFin(i) = 0
            Flag(i) = 0
        Next i
        MoveGroup = num_1 - 1
        MoveCount = num_2 + 1
        num_1 = 0
        num_2 = 0
        temp = 0
        tmp = 0
        HandleCount = 0
        shatmp = ""
        RepCheck = False
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Result As MsgBoxResult
        Dim i As Long, j As Long, k As Long, MoveFail As Long, DupCount As Long
        Dim MovePath As String, TimeNow As String, NameTemp As String, NameTmp As String
        Dim IgnFlag As Boolean
        ListBox4.Items.Add(Now + "请选择移动文件夹路径")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            MovePath = FolderBrowserDialog1.SelectedPath
            ListBox4.Items.Add(Now + "已选择移动文件夹路径:" + MovePath)
            ListBox4.Items.Add(Now + "正在确认文件夹路径是否有误")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            If Mid(FolderPath, Len(FolderPath), 1) <> "\" Then
                FolderPath += "\"
            End If
            If Mid(MovePath, Len(MovePath), 1) <> "\" Then
                MovePath += "\"
            End If
            ListBox4.Items.Add(Now + "文件夹路径无误,开始询问是否移动全部重复文件")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            Result = MsgBox("文件夹路径无误,开始移动全部重复文件至指定文件夹下" + vbCrLf + "请在确认无误后点击是,存疑点否", 0 + vbExclamation + vbYesNo, "警告")
            If Result = vbYes Then
                Button3.Enabled = False
                Button4.Enabled = False
                Button5.Enabled = False
                Button6.Enabled = False
                ComboBox1.Enabled = False
                ListBox4.Items.Add(Now + "已选择是,开始进行重复文件分组移动")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                If Boo_DirExist(MovePath) = 0 Then
                    ListBox4.Items.Add(Now + "移动文件夹不存在,程序将自动创建移动文件夹")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    Try
                        Dim OutFolder As New DirectoryInfo(MovePath)
                        OutFolder.Create()
                    Catch ex As Exception
                        ListBox4.Items.Add(Now + "创建移动文件夹失败,程序已终止")
                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                        MsgBox("错误:创建移动文件夹失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "移动文件夹创建失败")
                        Exit Sub
                    End Try
                End If
                ListBox4.Items.Add(Now + "移动文件夹路径处理完成,开始在文件夹路径内创建日期文件夹")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                TimeNow = Format(Now, "yyyyMMddHHmmss")
                Try
                    Dim OutFolder As New DirectoryInfo(MovePath + TimeNow)
                    OutFolder.Create()
                Catch ex As Exception
                    ListBox4.Items.Add(Now + "创建日期文件夹失败,程序已终止")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    MsgBox("错误:创建日期文件夹失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "日期文件夹创建失败")
                    Exit Sub
                End Try
                ListBox4.Items.Add(Now + "日期文件夹:" + TimeNow + ",创建成功,开始创建分组文件夹并移动重复文件")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MoveFail = 0
                NameTemp = ""
                NameTmp = ""
                IgnFlag = False
                For i = 1 To MoveGroup
                    Try
                        Dim OutFolder As New DirectoryInfo(MovePath + TimeNow + "\" + CStr(i))
                        OutFolder.Create()
                    Catch ex As Exception
                        ListBox4.Items.Add(Now + "分组文件夹:" + CStr(i) + "创建失败,程序已终止")
                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                        MsgBox("错误:创建分组文件夹失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "分组文件夹创建失败")
                        Exit Sub
                    End Try
                    ListBox4.Items.Add(Now + "分组文件夹:" + CStr(i) + ",创建成功,开始移动当前组重复文件")
                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                    For j = 0 To MoveCount - 1
                        If FlagTemp(j) = i Then
                            If File.Exists(FolderPath + HandleData(j)) = False Then
                                MoveFail += 1
                                ListBox4.Items.Add(Now + "第" + CStr(i) + "组重复文件:" + HandleData(j) + "移动失败")
                                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                If IgnFlag = False Then
                                    MsgBox("错误:" + HandleData(j) + vbCrLf + "文件移动失败,文件不存在", 0 + vbCritical + vbSystemModal, "移动重复文件失败")
                                    ListBox4.Items.Add(Now + "已发生移动失败,开始询问是否继续执行移动操作")
                                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    Result = MsgBox("已发生移动失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来所有的移动失败警告", 0 + vbExclamation + vbYesNo, "警告")
                                    If Result = vbYes Then
                                        IgnFlag = True
                                        ListBox4.Items.Add(Now + "已选择是,将继续执行并忽略移动失败警告")
                                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    Else
                                        ListBox4.Items.Add(Now + "已选择否,重复文件移动操作已终止")
                                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                For k = Len(HandleData(j)) To 1 Step -1
                                    If StrComp(Mid(HandleData(j), k, 1), "\") = 0 Then
                                        NameTemp = Mid(HandleData(j), k + 1)
                                        NameTmp = Mid(HandleData(j), k + 1)
                                        Exit For
                                    End If
                                Next k
                                If k = 0 Then
                                    NameTemp = HandleData(j)
                                    NameTmp = HandleData(j)
                                End If
                                If File.Exists(MovePath + TimeNow + "\" + CStr(i) + "\" + NameTemp) = True Then
                                    For k = Len(NameTemp) To 1 Step -1
                                        If StrComp(Mid(NameTemp, k, 1), ".") = 0 Then
                                            Exit For
                                        End If
                                    Next k
                                    DupCount = 0
                                    Do While (File.Exists(MovePath + TimeNow + "\" + CStr(i) + "\" + NameTemp) = True)
                                        DupCount += 1
                                        If k = 0 Then
                                            NameTemp = NameTmp + "_" + CStr(DupCount)
                                        Else
                                            NameTemp = Mid(NameTmp, 1, k - 1) + "_" + CStr(DupCount) + Mid(NameTmp, k)
                                        End If
                                    Loop
                                    If Len(MovePath + TimeNow + "\" + CStr(i) + "\" + NameTemp) > 252 Then
                                        NameTemp = Mid(NameTemp, Len(MovePath + TimeNow + "\" + CStr(i) + "\" + NameTemp) - 251)
                                        For k = Len(NameTemp) To 1 Step -1
                                            If StrComp(Mid(NameTemp, k, 1), ".") = 0 Then
                                                Exit For
                                            End If
                                        Next k
                                        DupCount = 0
                                        Do While (File.Exists(MovePath + TimeNow + "\" + CStr(i) + "\" + NameTemp) = True)
                                            DupCount += 1
                                            If k = 0 Then
                                                NameTemp = NameTmp + "_" + CStr(DupCount)
                                            Else
                                                NameTemp = Mid(NameTmp, 1, k - 1) + "_" + CStr(DupCount) + Mid(NameTmp, k)
                                            End If
                                        Loop
                                        ListBox4.Items.Add(Now + "出现文件名长度超出Windows系统限制的文件,已进行重命名处理")
                                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    End If
                                    ListBox4.Items.Add(Now + "出现文件名重复,已进行重命名处理")
                                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                End If
                                Try
                                    File.Move(FolderPath + HandleData(j), MovePath + TimeNow + "\" + CStr(i) + "\" + NameTemp)
                                Catch ex As Exception
                                    MoveFail += 1
                                    ListBox4.Items.Add(Now + "第" + CStr(i) + "组重复文件:" + HandleData(j) + "移动失败")
                                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    If IgnFlag = False Then
                                        MsgBox("错误:" + HandleData(j) + vbCrLf + "文件移动失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "移动重复文件失败")
                                        ListBox4.Items.Add(Now + "已发生移动失败,开始询问是否继续执行移动操作")
                                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                        Result = MsgBox("已发生移动失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来所有的移动失败警告", 0 + vbExclamation + vbYesNo, "警告")
                                        If Result = vbYes Then
                                            IgnFlag = True
                                            ListBox4.Items.Add(Now + "已选择是,将继续执行并忽略移动失败警告")
                                            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                        Else
                                            ListBox4.Items.Add(Now + "已选择否,重复文件移动操作已终止")
                                            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                            Exit Sub
                                        End If
                                    End If
                                End Try
                                ListBox4.Items.Add(Now + "第" + CStr(i) + "组重复文件:" + HandleData(j) + "移动完成")
                                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                            End If
                        End If
                    Next j
                Next i
                ListBox4.Items.Add(Now + "分组文件夹创建完成,重复文件移动完成,计划移动:" + CStr(MoveCount) + ",成功:" + CStr(MoveCount - MoveFail) + ",失败:" + CStr(MoveFail))
                ListBox4.Items.Add(Now + "请打开:" + MovePath + TimeNow + ",移动文件夹路径确认移动情况")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MsgBox("分组文件夹创建完成,重复文件移动完成,计划移动:" + CStr(MoveCount) + ",成功:" + CStr(MoveCount - MoveFail) + ",失败:" + CStr(MoveFail) +
                       vbCrLf + "请打开:" + MovePath + TimeNow + vbCrLf + "移动文件夹路径确认移动情况", 0 + vbInformation + vbSystemModal, "重复文件移动完成")
            Else
                ListBox4.Items.Add(Now + "已选择否,重复文件分组移动操作已终止")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            End If
        Else
            ListBox4.Items.Add(Now + "已取消移动文件夹路径选择")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        ComboBox1.Enabled = False
        ListBox4.Items.Add(Now + "开始选择保留文件")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Form2.Show()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim Result As MsgBoxResult
        Dim i As Long, j As Long, DelCount As Long = 0, DelFail As Long = 0
        Dim IgnFlag As Boolean
        ListBox4.Items.Add(Now + "开始询问是否批量删除重复文件")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Result = MsgBox("开始根据第二轮查重结果及用户选择保留情况批量删除重复文件" + vbCrLf + "请在确认无误后点击是,存疑点否", 0 + vbExclamation + vbYesNo, "警告")
        If Result = vbYes Then
            Button3.Enabled = False
            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
            ComboBox1.Enabled = False
            IgnFlag = False
            ListBox4.Items.Add(Now + "已选择是,开始批量删除重复文件")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            DelFail = 0
            For i = 1 To MoveGroup
                For j = 0 To MoveCount - 1
                    If StrComp(Del_Data(i - 1), HandleData(j)) <> 0 And FlagTemp(j) = i _
                        And StrComp(Del_Data(i - 1), "本组筛选后无重复文件") <> 0 _
                        And (ResFolderFlag(HandleData(j)) Xor FeFeEn_Flag) = False Then
                        If File.Exists(FolderPath + HandleData(j)) = False Then
                            DelFail += 1
                            ListBox4.Items.Add(Now + "第" + CStr(i) + "组重复文件:" + HandleData(j) + "删除失败")
                            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                            If IgnFlag = False Then
                                MsgBox("错误:" + HandleData(j) + vbCrLf + "文件删除失败,文件不存在", 0 + vbCritical + vbSystemModal, "批量删除重复文件失败")
                                ListBox4.Items.Add(Now + "已发生删除失败,开始询问是否继续执行删除操作")
                                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                Result = MsgBox("已发生删除失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来所有的删除失败警告", 0 + vbExclamation + vbYesNo, "警告")
                                If Result = vbYes Then
                                    IgnFlag = True
                                    ListBox4.Items.Add(Now + "已选择是,将继续执行并忽略删除失败警告")
                                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                Else
                                    ListBox4.Items.Add(Now + "已选择否,重复文件删除操作已终止")
                                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    Exit Sub
                                End If
                            End If
                        Else
                            DelCount += 1
                            Try
                                My.Computer.FileSystem.DeleteFile(FolderPath + HandleData(j), FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                            Catch ex As Exception
                                DelFail += 1
                                ListBox4.Items.Add(Now + "第" + CStr(i) + "组重复文件:" + HandleData(j) + "删除失败")
                                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                If IgnFlag = False Then
                                    MsgBox("错误:" + HandleData(j) + vbCrLf + "文件删除失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "批量删除重复文件失败")
                                    ListBox4.Items.Add(Now + "已发生删除失败,开始询问是否继续执行删除操作")
                                    ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    Result = MsgBox("已发生删除失败,是否继续执行?" + vbCrLf + "继续执行将忽略未来所有的删除失败警告", 0 + vbExclamation + vbYesNo, "警告")
                                    If Result = vbYes Then
                                        IgnFlag = True
                                        ListBox4.Items.Add(Now + "已选择是,将继续执行并忽略删除失败警告")
                                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                    Else
                                        ListBox4.Items.Add(Now + "已选择否,重复文件删除操作已终止")
                                        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                                        Exit Sub
                                    End If
                                End If
                            End Try
                            ListBox4.Items.Add(Now + "第" + CStr(i) + "组重复文件:" + HandleData(j) + "删除完成")
                            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                        End If
                    End If
                Next j
            Next i
            If DelCount > 0 Then
                ListBox4.Items.Add(Now + "重复文件删除完成,计划删除:" + CStr(DelCount) + ",成功:" + CStr(DelCount - DelFail) + ",失败:" + CStr(DelFail))
                ListBox4.Items.Add(Now + "删除的文件位于回收站中,请检查查重文件夹和回收站删除情况")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MsgBox("重复文件删除完成,计划删除:" + CStr(DelCount) + ",成功:" + CStr(DelCount - DelFail) + ",失败:" + CStr(DelFail) +
                       vbCrLf + "删除的文件位于回收站中,请检查查重文件夹和回收站删除情况", 0 + vbInformation + vbSystemModal, "批量删除重复文件完成")
            Else
                ListBox4.Items.Add(Now + "没有被删除的重复文件,可再次进入选择保留文件界面调整选择状态")
                ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
                MsgBox("没有被删除的重复文件,可再次进入选择保留文件界面调整选择状态", 0 + vbInformation + vbSystemModal, "批量删除重复文件完成")
                Button3.Enabled = True
                Button4.Enabled = True
                Button5.Enabled = True
                ComboBox1.Enabled = True
            End If
        Else
            ListBox4.Items.Add(Now + "已选择否,重复文件删除操作已终止")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        End If
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim OutputPath As String, FolderTime As String
        Dim i As Long
        Button7.Enabled = False
        ListBox4.Items.Add(Now + "请选择导出查重结果文件夹路径")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            OutputPath = FolderBrowserDialog1.SelectedPath
            ListBox4.Items.Add(Now + "已选择导出查重结果文件夹路径:" + OutputPath)
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Else
            ListBox4.Items.Add(Now + "已取消导出查重结果文件夹路径选择")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            Button7.Enabled = True
            Exit Sub
        End If
        If Mid(OutputPath, Len(OutputPath), 1) <> "\" Then OutputPath += "\"
        FolderTime = Format(Now, "yyyyMMddHHmmss") + ".txt"
        ListBox4.Items.Add(Now + "开始导出查重结果")
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Try
            My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, "此文本文档包含了:文件列表、第一轮查重对比数据(文件大小)列表、第二轮查重对比数据(SHA1校验值)列表", True)
            My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf, True)
            My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, "当导出数据较多时,可查找上述关键词以快速跳转至需要位置", True)
            My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf + vbCrLf + vbCrLf, True)
            For i = 0 To ListBox1.Items.Count - 1
                My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, ListBox1.Items(i), True)
                My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf, True)
            Next i
            My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf + vbCrLf, True)
            For i = 0 To ListBox2.Items.Count - 1
                My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, ListBox2.Items(i), True)
                My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf, True)
            Next i
            My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf + vbCrLf, True)
            For i = 0 To ListBox3.Items.Count - 1
                My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, ListBox3.Items(i), True)
                My.Computer.FileSystem.WriteAllText(OutputPath + FolderTime, vbCrLf, True)
            Next i
        Catch ex As Exception
            ListBox4.Items.Add(Now + "导出数据写入失败,请检查路径写入权限")
            ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
            MsgBox("错误:导出数据写入失败,请检查路径写入权限", 0 + vbCritical + vbSystemModal, "导出数据写入失败")
            Button7.Enabled = True
            Exit Sub
        End Try
        ListBox4.Items.Add(Now + "查重结果导出成功")
        ListBox4.Items.Add(Now + "导出文本文档的完整路径与文件名为:" + OutputPath + FolderTime)
        ListBox4.SelectedItem = ListBox4.Items(ListBox4.Items.Count - 1)
        Button7.Enabled = True
    End Sub
    Private Shared Function Boo_DirExist(StrPath As String) As Boolean
        Boo_DirExist = Directory.Exists(StrPath)
    End Function
    Private Function FolderList(FolderPath_Data As String) As Boolean
        Dim i As Long, FolderCount As Long
        For i = Len(FolderPath_Data) - 1 To 1 Step -1
            If StrComp(Mid(FolderPath_Data, i, 1), "\") = 0 Then
                FolderPath_Data = Mid(FolderPath_Data, 1, i - 1)
                Exit For
            End If
        Next i
        For i = 0 To UppLim - 1
            If Len(Folder_Data(i)) = 0 Then
                FolderCount = i
                Exit For
            End If
        Next i
        For i = 0 To FolderCount - 1
            If StrComp(Folder_Data(i), FolderPath_Data) = 0 Then Exit For
        Next i
        If i >= FolderCount Then
            Folder_Data(FolderCount) = FolderPath_Data
            FolderList = True
        Else
            FolderList = False
        End If
    End Function
    Private Function ResFolderFlag(ResFolderName As String) As Boolean
        Dim i As Long
        For i = 0 To ResFolderCount - 1
            If StrComp(ResFolderList(i), Mid(FolderPath + ResFolderName, 1, Len(ResFolderList(i)))) = 0 Then Exit For
        Next i
        If i >= ResFolderCount Then
            ResFolderFlag = False
        Else
            ResFolderFlag = True
        End If
    End Function
    Public Shared Function NumLength(Num As String) As String
        If Len(Num) = 1 Then
            NumLength = "     " + Num
        ElseIf Len(Num) = 2 Then
            NumLength = "    " + Num
        ElseIf Len(Num) = 3 Then
            NumLength = "   " + Num
        ElseIf Len(Num) = 4 Then
            NumLength = "  " + Num
        ElseIf Len(Num) = 5 Then
            NumLength = " " + Num
        Else
            NumLength = Num
        End If
    End Function
End Class
Class HashValueCalculation
    Public sn As Long
    Public sn_filestream As FileStream
    Public sn_hashresult As String
    Public Sub ThreadWork()
        Using mySHA1 As SHA1 = SHA1.Create()
            sn_filestream.Position = 0
            Dim hashvalue() As Byte = mySHA1.ComputeHash(sn_filestream)
            sn_filestream.Close()
            sn_hashresult = Byte2Str(hashvalue)
        End Using
    End Sub
    Private Shared Function Byte2Str(HashByte() As Byte) As String
        Dim i As Integer
        Dim sb As New StringBuilder()
        For i = 0 To HashByte.Length - 1
            sb.Append($"{HashByte(i):X2}")
        Next i
        Return sb.ToString
    End Function
End Class