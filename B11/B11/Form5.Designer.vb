<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
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
        Me.date1 = New System.Windows.Forms.DateTimePicker()
        Me.date2 = New System.Windows.Forms.DateTimePicker()
        Me.Date3 = New System.Windows.Forms.DateTimePicker()
        Me.DTadd1 = New System.Windows.Forms.DateTimePicker()
        Me.DTadd2 = New System.Windows.Forms.DateTimePicker()
        Me.DTgraf2 = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'date1
        '
        Me.date1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.date1.Location = New System.Drawing.Point(61, 37)
        Me.date1.Name = "date1"
        Me.date1.Size = New System.Drawing.Size(200, 20)
        Me.date1.TabIndex = 0
        '
        'date2
        '
        Me.date2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.date2.Location = New System.Drawing.Point(61, 63)
        Me.date2.Name = "date2"
        Me.date2.Size = New System.Drawing.Size(200, 20)
        Me.date2.TabIndex = 1
        '
        'Date3
        '
        Me.Date3.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.Date3.Location = New System.Drawing.Point(62, 89)
        Me.Date3.Name = "Date3"
        Me.Date3.Size = New System.Drawing.Size(200, 20)
        Me.Date3.TabIndex = 2
        '
        'DTadd1
        '
        Me.DTadd1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTadd1.Location = New System.Drawing.Point(61, 155)
        Me.DTadd1.Name = "DTadd1"
        Me.DTadd1.Size = New System.Drawing.Size(200, 20)
        Me.DTadd1.TabIndex = 3
        '
        'DTadd2
        '
        Me.DTadd2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTadd2.Location = New System.Drawing.Point(61, 181)
        Me.DTadd2.Name = "DTadd2"
        Me.DTadd2.Size = New System.Drawing.Size(200, 20)
        Me.DTadd2.TabIndex = 4
        '
        'DTgraf2
        '
        Me.DTgraf2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTgraf2.Location = New System.Drawing.Point(62, 259)
        Me.DTgraf2.Name = "DTgraf2"
        Me.DTgraf2.Size = New System.Drawing.Size(200, 20)
        Me.DTgraf2.TabIndex = 6
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(35, 18)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(248, 105)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "date untuk trend"
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(472, 389)
        Me.ControlBox = False
        Me.Controls.Add(Me.DTgraf2)
        Me.Controls.Add(Me.DTadd2)
        Me.Controls.Add(Me.DTadd1)
        Me.Controls.Add(Me.Date3)
        Me.Controls.Add(Me.date2)
        Me.Controls.Add(Me.date1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form5"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form5"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents date1 As DateTimePicker
    Friend WithEvents date2 As DateTimePicker
    Friend WithEvents Date3 As DateTimePicker
    Friend WithEvents DTadd1 As DateTimePicker
    Friend WithEvents DTadd2 As DateTimePicker
    Friend WithEvents DTgraf2 As DateTimePicker
    Friend WithEvents GroupBox1 As GroupBox
End Class
