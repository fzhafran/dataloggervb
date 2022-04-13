<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CommFADV
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
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.jmlhsalah2 = New System.Windows.Forms.TextBox()
        Me.n2 = New System.Windows.Forms.TextBox()
        Me.address2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.dei_request2 = New System.Windows.Forms.TextBox()
        Me.dei_list2 = New System.Windows.Forms.ListBox()
        Me.dei_crc_lo = New System.Windows.Forms.TextBox()
        Me.slaveAdd2 = New System.Windows.Forms.TextBox()
        Me.dei_crc_hi = New System.Windows.Forms.TextBox()
        Me.dei_inBuffer2 = New System.Windows.Forms.TextBox()
        Me.dei_response2 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Coral
        Me.Label7.Location = New System.Drawing.Point(54, 165)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(68, 13)
        Me.Label7.TabIndex = 137
        Me.Label7.Text = "Jumlah salah"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Coral
        Me.Label6.Location = New System.Drawing.Point(241, 165)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(13, 13)
        Me.Label6.TabIndex = 136
        Me.Label6.Text = "n"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Coral
        Me.Label5.Location = New System.Drawing.Point(202, 132)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 135
        Me.Label5.Text = "address(n)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Coral
        Me.Label4.Location = New System.Drawing.Point(189, 208)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(166, 13)
        Me.Label4.TabIndex = 134
        Me.Label4.Text = "*mengganti nilainya di timer waktu"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Coral
        Me.Label3.Location = New System.Drawing.Point(59, 187)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(196, 13)
        Me.Label3.TabIndex = 133
        Me.Label3.Text = "Nilai untuk mengaktifkan timer database"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Coral
        Me.Label1.Location = New System.Drawing.Point(179, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 20)
        Me.Label1.TabIndex = 131
        Me.Label1.Text = "Comm FADV"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(60, 237)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 130
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'jmlhsalah2
        '
        Me.jmlhsalah2.Location = New System.Drawing.Point(124, 160)
        Me.jmlhsalah2.Multiline = True
        Me.jmlhsalah2.Name = "jmlhsalah2"
        Me.jmlhsalah2.Size = New System.Drawing.Size(55, 20)
        Me.jmlhsalah2.TabIndex = 122
        '
        'n2
        '
        Me.n2.Location = New System.Drawing.Point(264, 158)
        Me.n2.Multiline = True
        Me.n2.Name = "n2"
        Me.n2.Size = New System.Drawing.Size(37, 26)
        Me.n2.TabIndex = 121
        '
        'address2
        '
        Me.address2.Location = New System.Drawing.Point(264, 125)
        Me.address2.Multiline = True
        Me.address2.Name = "address2"
        Me.address2.Size = New System.Drawing.Size(37, 26)
        Me.address2.TabIndex = 120
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(60, 205)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(126, 26)
        Me.TextBox1.TabIndex = 119
        '
        'dei_request2
        '
        Me.dei_request2.Location = New System.Drawing.Point(60, 68)
        Me.dei_request2.Multiline = True
        Me.dei_request2.Name = "dei_request2"
        Me.dei_request2.Size = New System.Drawing.Size(117, 52)
        Me.dei_request2.TabIndex = 112
        Me.dei_request2.Text = "req"
        '
        'dei_list2
        '
        Me.dei_list2.FormattingEnabled = True
        Me.dei_list2.Location = New System.Drawing.Point(313, 42)
        Me.dei_list2.Name = "dei_list2"
        Me.dei_list2.Size = New System.Drawing.Size(126, 160)
        Me.dei_list2.TabIndex = 118
        '
        'dei_crc_lo
        '
        Me.dei_crc_lo.Location = New System.Drawing.Point(124, 132)
        Me.dei_crc_lo.Name = "dei_crc_lo"
        Me.dei_crc_lo.Size = New System.Drawing.Size(55, 20)
        Me.dei_crc_lo.TabIndex = 117
        Me.dei_crc_lo.Text = "crc_lo"
        '
        'slaveAdd2
        '
        Me.slaveAdd2.Location = New System.Drawing.Point(60, 42)
        Me.slaveAdd2.Name = "slaveAdd2"
        Me.slaveAdd2.Size = New System.Drawing.Size(55, 20)
        Me.slaveAdd2.TabIndex = 114
        Me.slaveAdd2.Text = "slaveAdd"
        '
        'dei_crc_hi
        '
        Me.dei_crc_hi.Location = New System.Drawing.Point(62, 132)
        Me.dei_crc_hi.Name = "dei_crc_hi"
        Me.dei_crc_hi.Size = New System.Drawing.Size(55, 20)
        Me.dei_crc_hi.TabIndex = 116
        Me.dei_crc_hi.Text = "crc_hi"
        '
        'dei_inBuffer2
        '
        Me.dei_inBuffer2.Location = New System.Drawing.Point(244, 42)
        Me.dei_inBuffer2.Name = "dei_inBuffer2"
        Me.dei_inBuffer2.Size = New System.Drawing.Size(55, 20)
        Me.dei_inBuffer2.TabIndex = 115
        Me.dei_inBuffer2.Text = "inBuffer"
        '
        'dei_response2
        '
        Me.dei_response2.Location = New System.Drawing.Point(183, 68)
        Me.dei_response2.Multiline = True
        Me.dei_response2.Name = "dei_response2"
        Me.dei_response2.Size = New System.Drawing.Size(115, 51)
        Me.dei_response2.TabIndex = 113
        Me.dei_response2.Text = "respon"
        '
        'CommFADV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(484, 281)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.jmlhsalah2)
        Me.Controls.Add(Me.n2)
        Me.Controls.Add(Me.address2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.dei_request2)
        Me.Controls.Add(Me.dei_list2)
        Me.Controls.Add(Me.dei_crc_lo)
        Me.Controls.Add(Me.slaveAdd2)
        Me.Controls.Add(Me.dei_crc_hi)
        Me.Controls.Add(Me.dei_inBuffer2)
        Me.Controls.Add(Me.dei_response2)
        Me.Name = "CommFADV"
        Me.Text = "CommFADV"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents jmlhsalah2 As TextBox
    Friend WithEvents n2 As TextBox
    Friend WithEvents address2 As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents dei_request2 As TextBox
    Friend WithEvents dei_list2 As ListBox
    Friend WithEvents dei_crc_lo As TextBox
    Friend WithEvents slaveAdd2 As TextBox
    Friend WithEvents dei_crc_hi As TextBox
    Friend WithEvents dei_inBuffer2 As TextBox
    Friend WithEvents dei_response2 As TextBox
End Class
