<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TabelTZ
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
        Me.ReportTabel = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.CrystalReport21 = New B11.CrystalReport2()
        Me.SuspendLayout()
        '
        'ReportTabel
        '
        Me.ReportTabel.ActiveViewIndex = 0
        Me.ReportTabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ReportTabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.ReportTabel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ReportTabel.Location = New System.Drawing.Point(0, 0)
        Me.ReportTabel.Name = "ReportTabel"
        Me.ReportTabel.ReportSource = Me.CrystalReport21
        Me.ReportTabel.Size = New System.Drawing.Size(701, 452)
        Me.ReportTabel.TabIndex = 0
        Me.ReportTabel.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'TabelTZ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(701, 452)
        Me.Controls.Add(Me.ReportTabel)
        Me.Name = "TabelTZ"
        Me.Text = "TabelTZ"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ReportTabel As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents CrystalReport21 As CrystalReport2
End Class
