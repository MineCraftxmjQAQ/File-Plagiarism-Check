<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
    Inherits System.Windows.Forms.Form
    'Form 重写 Dispose,以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer
    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        CheckBox1 = New CheckBox()
        ListBox3 = New ListBox()
        ListBox2 = New ListBox()
        ListBox1 = New ListBox()
        CheckBox2 = New CheckBox()
        CheckBox3 = New CheckBox()
        Button1 = New Button()
        CheckedListBox1 = New CheckedListBox()
        Button2 = New Button()
        TextBox1 = New TextBox()
        TextBox2 = New TextBox()
        TextBox3 = New TextBox()
        ComboBox1 = New ComboBox()
        ComboBox2 = New ComboBox()
        ComboBox3 = New ComboBox()
        SuspendLayout()
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.Location = New Point(1044, 12)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(99, 21)
        CheckBox1.TabIndex = 16
        CheckBox1.Text = "启用高级筛选"
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' ListBox3
        ' 
        ListBox3.FormattingEnabled = True
        ListBox3.HorizontalScrollbar = True
        ListBox3.ItemHeight = 17
        ListBox3.Location = New Point(573, 12)
        ListBox3.Name = "ListBox3"
        ListBox3.Size = New Size(465, 684)
        ListBox3.TabIndex = 15
        ' 
        ' ListBox2
        ' 
        ListBox2.FormattingEnabled = True
        ListBox2.HorizontalScrollbar = True
        ListBox2.ItemHeight = 17
        ListBox2.Location = New Point(102, 12)
        ListBox2.Name = "ListBox2"
        ListBox2.Size = New Size(465, 684)
        ListBox2.TabIndex = 14
        ' 
        ' ListBox1
        ' 
        ListBox1.FormattingEnabled = True
        ListBox1.HorizontalScrollbar = True
        ListBox1.ItemHeight = 17
        ListBox1.Location = New Point(12, 12)
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(84, 684)
        ListBox1.TabIndex = 13
        ' 
        ' CheckBox2
        ' 
        CheckBox2.AutoSize = True
        CheckBox2.Enabled = False
        CheckBox2.Location = New Point(1044, 39)
        CheckBox2.Name = "CheckBox2"
        CheckBox2.Size = New Size(111, 21)
        CheckBox2.TabIndex = 17
        CheckBox2.Text = "启用文件夹筛选"
        CheckBox2.UseVisualStyleBackColor = True
        ' 
        ' CheckBox3
        ' 
        CheckBox3.AutoSize = True
        CheckBox3.Enabled = False
        CheckBox3.Location = New Point(1044, 582)
        CheckBox3.Name = "CheckBox3"
        CheckBox3.Size = New Size(111, 21)
        CheckBox3.TabIndex = 18
        CheckBox3.Text = "启用文件名筛选"
        CheckBox3.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.Enabled = False
        Button1.Location = New Point(1044, 66)
        Button1.Name = "Button1"
        Button1.Size = New Size(90, 24)
        Button1.TabIndex = 19
        Button1.Text = "反选"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' CheckedListBox1
        ' 
        CheckedListBox1.Enabled = False
        CheckedListBox1.FormattingEnabled = True
        CheckedListBox1.HorizontalScrollbar = True
        CheckedListBox1.Location = New Point(1044, 96)
        CheckedListBox1.Name = "CheckedListBox1"
        CheckedListBox1.Size = New Size(375, 472)
        CheckedListBox1.TabIndex = 20
        ' 
        ' Button2
        ' 
        Button2.Enabled = False
        Button2.Location = New Point(1140, 66)
        Button2.Name = "Button2"
        Button2.Size = New Size(90, 24)
        Button2.TabIndex = 21
        Button2.Text = "全选"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' TextBox1
        ' 
        TextBox1.Enabled = False
        TextBox1.Location = New Point(1171, 611)
        TextBox1.Name = "TextBox1"
        TextBox1.PlaceholderText = "在此键入第一个关键词"
        TextBox1.Size = New Size(248, 23)
        TextBox1.TabIndex = 22
        ' 
        ' TextBox2
        ' 
        TextBox2.Enabled = False
        TextBox2.Location = New Point(1171, 642)
        TextBox2.Name = "TextBox2"
        TextBox2.PlaceholderText = "在此键入第二个关键词"
        TextBox2.Size = New Size(248, 23)
        TextBox2.TabIndex = 23
        ' 
        ' TextBox3
        ' 
        TextBox3.Enabled = False
        TextBox3.Location = New Point(1171, 673)
        TextBox3.Name = "TextBox3"
        TextBox3.PlaceholderText = "在此键入第三个关键词"
        TextBox3.Size = New Size(248, 23)
        TextBox3.TabIndex = 24
        ' 
        ' ComboBox1
        ' 
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox1.Enabled = False
        ComboBox1.FormattingEnabled = True
        ComboBox1.Items.AddRange(New Object() {"包含关键词", "排除关键词", "不启用此项"})
        ComboBox1.Location = New Point(1044, 609)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(121, 25)
        ComboBox1.TabIndex = 25
        ' 
        ' ComboBox2
        ' 
        ComboBox2.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox2.Enabled = False
        ComboBox2.FormattingEnabled = True
        ComboBox2.Items.AddRange(New Object() {"包含关键词", "排除关键词", "不启用此项"})
        ComboBox2.Location = New Point(1044, 640)
        ComboBox2.Name = "ComboBox2"
        ComboBox2.Size = New Size(121, 25)
        ComboBox2.TabIndex = 26
        ' 
        ' ComboBox3
        ' 
        ComboBox3.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox3.Enabled = False
        ComboBox3.FormattingEnabled = True
        ComboBox3.Items.AddRange(New Object() {"包含关键词", "排除关键词", "不启用此项"})
        ComboBox3.Location = New Point(1044, 671)
        ComboBox3.Name = "ComboBox3"
        ComboBox3.Size = New Size(121, 25)
        ComboBox3.TabIndex = 27
        ' 
        ' Form2
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1431, 714)
        Controls.Add(ComboBox3)
        Controls.Add(ComboBox2)
        Controls.Add(ComboBox1)
        Controls.Add(TextBox3)
        Controls.Add(TextBox2)
        Controls.Add(TextBox1)
        Controls.Add(Button2)
        Controls.Add(CheckedListBox1)
        Controls.Add(Button1)
        Controls.Add(CheckBox3)
        Controls.Add(CheckBox2)
        Controls.Add(CheckBox1)
        Controls.Add(ListBox3)
        Controls.Add(ListBox2)
        Controls.Add(ListBox1)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        MinimizeBox = False
        Name = "Form2"
        ShowInTaskbar = False
        StartPosition = FormStartPosition.CenterScreen
        Text = "选择保留文件"
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents ListBox3 As ListBox
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents Button1 As Button
    Friend WithEvents CheckedListBox1 As CheckedListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox3 As ComboBox
End Class
