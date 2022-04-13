Public Class password
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "teknisi" Then
            If Val(Label2.Text) = 2 Then baru.Show()
            If Val(Label2.Text) = 1 Then commqualityTZ.Show()
            If Val(Label2.Text) = 4 Then CommFADV.Show()
            If Val(Label2.Text) = 3 Then
                Form1.Label77.Visible = True
                Form1.Label88.Visible = True
                Form1.Label89.Visible = True
                Form1.Label86.Visible = True
                Form1.Label85.Visible = True
                Form1.TextBox2.Visible = True
                Form1.setval.Visible = True
                Form1.ComboBox2.Visible = True
                Form1.waktuon.Visible = True
                Form1.waktuoff.Visible = True
                Form1.Button15.Visible = True
                Form1.Button17.Visible = True
                Form1.Button18.Visible = True
                Form1.Button22.Visible = True
            End If
            Form1.Label83.Text = "Teknisi"
        Else
            MsgBox("Password Salah")
        End If
        Me.Close()
    End Sub

    Private Sub password_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class