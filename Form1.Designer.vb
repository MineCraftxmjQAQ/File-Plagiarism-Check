<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form
    'Form overrides dispose to clean up the component list.
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
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Button9 = New Button()
        Button8 = New Button()
        Button7 = New Button()
        Button6 = New Button()
        Button5 = New Button()
        Button4 = New Button()
        ListBox4 = New ListBox()
        Button3 = New Button()
        ListBox3 = New ListBox()
        ListBox2 = New ListBox()
        ListBox1 = New ListBox()
        Button2 = New Button()
        Button1 = New Button()
        FolderBrowserDialog1 = New FolderBrowserDialog()
        Button10 = New Button()
        SuspendLayout()
        ' 
        ' Button9
        ' 
        Button9.Location = New Point(1233, 12)
        Button9.Name = "Button9"
        Button9.Size = New Size(90, 24)
        Button9.TabIndex = 28
        Button9.Text = "关于文件查重"
        Button9.UseVisualStyleBackColor = True
        ' 
        ' Button8
        ' 
        Button8.Location = New Point(1137, 12)
        Button8.Name = "Button8"
        Button8.Size = New Size(90, 24)
        Button8.TabIndex = 27
        Button8.Text = "程序使用说明"
        Button8.UseVisualStyleBackColor = True
        ' 
        ' Button7
        ' 
        Button7.Enabled = False
        Button7.Location = New Point(588, 12)
        Button7.Name = "Button7"
        Button7.Size = New Size(90, 24)
        Button7.TabIndex = 26
        Button7.Text = "导出查重结果"
        Button7.UseVisualStyleBackColor = True
        ' 
        ' Button6
        ' 
        Button6.Enabled = False
        Button6.Location = New Point(492, 12)
        Button6.Name = "Button6"
        Button6.Size = New Size(90, 24)
        Button6.TabIndex = 25
        Button6.Text = "批量删除重复"
        Button6.UseVisualStyleBackColor = True
        ' 
        ' Button5
        ' 
        Button5.Enabled = False
        Button5.Location = New Point(396, 12)
        Button5.Name = "Button5"
        Button5.Size = New Size(90, 24)
        Button5.TabIndex = 24
        Button5.Text = "选择保留文件"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' Button4
        ' 
        Button4.Enabled = False
        Button4.Location = New Point(300, 12)
        Button4.Name = "Button4"
        Button4.Size = New Size(90, 24)
        Button4.TabIndex = 23
        Button4.Text = "移动重复文件"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' ListBox4
        ' 
        ListBox4.FormattingEnabled = True
        ListBox4.HorizontalScrollbar = True
        ListBox4.ItemHeight = 17
        ListBox4.Location = New Point(12, 596)
        ListBox4.Name = "ListBox4"
        ListBox4.Size = New Size(1407, 106)
        ListBox4.TabIndex = 22
        ' 
        ' Button3
        ' 
        Button3.Enabled = False
        Button3.Location = New Point(204, 12)
        Button3.Name = "Button3"
        Button3.Size = New Size(90, 24)
        Button3.TabIndex = 21
        Button3.Text = "文件查重"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' ListBox3
        ' 
        ListBox3.FormattingEnabled = True
        ListBox3.HorizontalScrollbar = True
        ListBox3.ItemHeight = 17
        ListBox3.Location = New Point(954, 42)
        ListBox3.Name = "ListBox3"
        ListBox3.Size = New Size(465, 548)
        ListBox3.TabIndex = 20
        ' 
        ' ListBox2
        ' 
        ListBox2.FormattingEnabled = True
        ListBox2.HorizontalScrollbar = True
        ListBox2.ItemHeight = 17
        ListBox2.Location = New Point(483, 42)
        ListBox2.Name = "ListBox2"
        ListBox2.Size = New Size(465, 548)
        ListBox2.TabIndex = 19
        ' 
        ' ListBox1
        ' 
        ListBox1.FormattingEnabled = True
        ListBox1.HorizontalScrollbar = True
        ListBox1.ItemHeight = 17
        ListBox1.Location = New Point(12, 42)
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(465, 548)
        ListBox1.TabIndex = 18
        ' 
        ' Button2
        ' 
        Button2.Enabled = False
        Button2.Location = New Point(108, 12)
        Button2.Name = "Button2"
        Button2.Size = New Size(90, 24)
        Button2.TabIndex = 17
        Button2.Text = "获取文件列表"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(12, 12)
        Button1.Name = "Button1"
        Button1.Size = New Size(90, 24)
        Button1.TabIndex = 16
        Button1.Text = "打开文件夹"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button10
        ' 
        Button10.Location = New Point(1329, 12)
        Button10.Name = "Button10"
        Button10.Size = New Size(90, 24)
        Button10.TabIndex = 29
        Button10.Text = "程序更新日志"
        Button10.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1431, 714)
        Controls.Add(Button9)
        Controls.Add(Button8)
        Controls.Add(Button7)
        Controls.Add(Button6)
        Controls.Add(Button5)
        Controls.Add(Button4)
        Controls.Add(ListBox4)
        Controls.Add(Button3)
        Controls.Add(ListBox3)
        Controls.Add(ListBox2)
        Controls.Add(ListBox1)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(Button10)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "文件查重V2.0.0"
        ResumeLayout(False)
    End Sub
    Friend WithEvents Button9 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents ListBox4 As ListBox
    Friend WithEvents Button3 As Button
    Friend WithEvents ListBox3 As ListBox
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents Button10 As Button
End Class
