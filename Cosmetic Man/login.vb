Public Class login

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "admin" And TextBox2.Text = "password" Then
            mmenu.Show()
            Me.Hide()
        ElseIf TextBox1.Text <> "admin" Then
            MsgBox("Incorrect Username")
        ElseIf TextBox2.Text <> "password" Then
            MsgBox("Incorrect Password")
        Else
            MsgBox("Incorrect Username & Password")
        End If
    End Sub

    Private Sub login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.UseSystemPasswordChar = False
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
