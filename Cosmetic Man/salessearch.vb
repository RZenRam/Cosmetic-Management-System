Imports System.Data.OleDb

Public Class salessearch
    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ob As Object
    Dim dt As New DataTable
    Dim Dadapter As OleDbDataAdapter
    Dim table As DataTableCollection
    Dim DSet As DataSet
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text Like "9*********" Then
            Try
                dt = New DataTable
                With cmd
                    .Connection = con
                    .CommandText = "Select * from sales where m_num like '" & TextBox1.Text & "%'"
                End With
                Dadapter.SelectCommand = cmd
                Dadapter.Fill(dt)

                DataGridView1.DataSource = dt

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Dadapter.Dispose()
        Else
            Try
                dt = New DataTable

                With cmd
                    .Connection = con
                    .CommandText = "Select * from sales where c_name like '" & TextBox1.Text & "%'"
                End With
                Dadapter.SelectCommand = cmd
                Dadapter.Fill(dt)

                DataGridView1.DataSource = dt

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Dadapter.Dispose()
        End If
    End Sub

    Private Sub salessearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Developing\bak\Cosmetic Man\Cosmetic.accdb")
        con.Open()
        Dadapter = New OleDbDataAdapter("select * from sales", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "sales")
            DataGridView1.DataSource = DSet.Tables("sales")
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Me.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
        Label2.Text = DataGridView1.CurrentRow.Cells("sales_id").Value.ToString
        Label5.Text = DataGridView1.CurrentRow.Cells("c_name").Value.ToString
        Label6.Text = DataGridView1.CurrentRow.Cells("m_num").Value.ToString
        Label9.Text = DataGridView1.CurrentRow.Cells("p_date").Value.ToString
        Label13.Text = DataGridView1.CurrentRow.Cells("p_name").Value.ToString
        Label14.Text = DataGridView1.CurrentRow.Cells("quantity").Value.ToString
        Label15.Text = DataGridView1.CurrentRow.Cells("total").Value.ToString
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd = New OleDbCommand("Select * from sales where m_num like '%" & TextBox1.Text & "%' ", con)
        dr = cmd.ExecuteReader()
        While dr.Read
            Label2.Text = dr("sales_id").ToString
            Label5.Text = dr("c_name").ToString
            Label6.Text = dr("m_num").ToString
            Label9.Text = dr("p_date").ToString
            Label13.Text = dr("p_name").ToString
            Label14.Text = dr("quantity").ToString
            Label15.Text = dr("total").ToString
        End While
        dr.Close()
        Dadapter = New OleDbDataAdapter("select * from sales", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "sales")
        Me.DataGridView1.DataSource = DSet.Tables("sales")
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        mmenu.Show()
        Me.Hide()

    End Sub
End Class