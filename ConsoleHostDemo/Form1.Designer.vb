<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ConsoleHostV21 = New ConsoleHostDemo.ConsoleHostV2()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 444)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Send Text:"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(68, 440)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(539, 20)
        Me.TextBox1.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(614, 437)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(70, 26)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Send"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(690, 437)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(80, 26)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "TestClolor"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ConsoleHostV21
        '
        Me.ConsoleHostV21.BackColor = System.Drawing.Color.FromArgb(CType(CType(12, Byte), Integer), CType(CType(12, Byte), Integer), CType(CType(12, Byte), Integer))
        Me.ConsoleHostV21.Dock = System.Windows.Forms.DockStyle.Top
        Me.ConsoleHostV21.ForeColor = System.Drawing.Color.White
        Me.ConsoleHostV21.Location = New System.Drawing.Point(0, 0)
        Me.ConsoleHostV21.Name = "ConsoleHostV21"
        Me.ConsoleHostV21.Size = New System.Drawing.Size(849, 428)
        Me.ConsoleHostV21.TabIndex = 5
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(776, 437)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(70, 26)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Matrix"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(849, 466)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.ConsoleHostV21)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents ConsoleHostV21 As ConsoleHostV2
    Friend WithEvents Button3 As Button
End Class
