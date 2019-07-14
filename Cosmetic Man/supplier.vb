Imports System.Data.OleDb
Public Class supplier
    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ob As Object
    Dim Dadapter As OleDbDataAdapter
    Dim table As DataTableCollection
    Dim DSet As DataSet
    Dim count As Integer
    Private Sub supplier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Developing\bak\Cosmetic Man\Cosmetic.accdb")
        con.Open()

        Dadapter = New OleDbDataAdapter("select * from supplier", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "supplier")
        Me.DataGridView1.DataSource = DSet.Tables("supplier")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cmd = New OleDbCommand("update supplier set look_name=' " & TextBox2.Text & " ', con_num=' " & TextBox3.Text & " ', prop=' " & TextBox4.Text & " ' where sup_id=' " & TextBox1.Text & " ' ", con)
        cmd.ExecuteNonQuery()

        Dadapter = New OleDbDataAdapter("select * from supplier", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "supplier")
        DataGridView1.DataSource = DSet.Tables("supplier")

        MessageBox.Show("Record updated successfully", "Update")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd = New OleDbCommand("insert into supplier values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')", con)
        cmd.ExecuteNonQuery()
        Dadapter = New OleDbDataAdapter("select * from supplier", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "supplier")
        DataGridView1.DataSource = DSet.Tables("supplier")
        MessageBox.Show("Record Inserted successfully", "Save")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        cmd = New OleDbCommand("delete from supplier where sup_id=' " & TextBox1.Text & " ' ", con)
        cmd.ExecuteNonQuery()
        Dadapter = New OleDbDataAdapter("select * from supplier", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "supplier")
        Me.DataGridView1.DataSource = DSet.Tables("supplier")
        MessageBox.Show("Record deleted successfully", "Delete")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cmd = New OleDbCommand("select * from supplier where sup_id=' " & TextBox1.Text & " '", con)
        dr = cmd.ExecuteReader()

        While dr.Read
            TextBox2.Text = dr("look_name").ToString
            TextBox3.Text = dr("con_num").ToString
            TextBox4.Text = dr("prop").ToString
        End While

        dr.Close()
        Dadapter = New OleDbDataAdapter("select * from supplier", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "supplier")
        Me.DataGridView1.DataSource = DSet.Tables("supplier")
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        cmd = New OleDbCommand("select count(*) from supplier ", con)
        dr = cmd.ExecuteReader()
        dr.Read()
        count = dr(0).ToString
        Dim extra As String
        extra = count + 1


        If extra <= 9 Then
            TextBox1.Text = "sup00" + extra
        ElseIf extra >= 99 Then
            TextBox1.Text = "sup" + extra
        Else
            TextBox1.Text = "sup0" + extra
        End If

        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        mmenu.Show()
        Me.Hide()

    End Sub
End Class