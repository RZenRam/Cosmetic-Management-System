Imports System.Data.OleDb

Public Class bill

    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ob As Object
    Dim Dadapter As OleDbDataAdapter
    Dim table As DataTableCollection
    Dim DSet As DataSet
    Public salescnt As Integer
    Dim cnt As Integer
    Public total As Integer

    Private Sub bill_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Developing\bak\Cosmetic Man\Cosmetic.accdb")
        con.Open()

        Dadapter = New OleDbDataAdapter("select * from tempbill", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "tempbill")
        DataGridView1.DataSource = DSet.Tables("tempbill")


        cmd = New OleDbCommand("select count(*) from bill", con)
        dr = cmd.ExecuteReader()
        dr.Read()
        cnt = dr(0).ToString
        Label2.Text = cnt + 1
        Label4.Text = sales.salesid
        Label10.Text = sales.total
        Label8.Text = sales.customername
        Label6.Text = sales.mobileno

        cmd = New OleDbCommand("insert into bill values('" & Label2.Text & "','" & Label4.Text & "','" & Label6.Text & "','" & Label8.Text & "','" & Label10.Text & "')", con)
        cmd.ExecuteNonQuery()


        Label12.Text = Date.Today
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        printin.Show()
        Me.Hide()

    End Sub
End Class