Imports System.Data.OleDb
Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RefreshGrid()

    End Sub

    Private Sub RefreshGrid()
        'Dim c As New OleDbCommand
        'c.Connection = conn
        'c.CommandText = "select * from Planets"

        'Dim ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        'da.Fill(ds, "Planets")
        'DataGridView1.DataSource = ds
        'DataGridView1.DataMember = "Planets"

        FillGridDA(DataGridView1, "select * from Planets", "Planets", DA1)
        DataGridView1.Columns("Код").Visible = False
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _conn = New OleDbConnection
        _conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Планеты\Planets.accdb;Persist Security Info=False"
        _conn.Open()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim s1, s2, s3 As String
        Dim r As DialogResult

        Form3.ShowDialog()

        Try
            s1 = Integer.Parse(Form3.TextBox1.Text)
            s2 = Double.Parse(Form3.TextBox2.Text)
            s3 = Double.Parse(Form3.TextBox3.Text)

        Catch ex As Exception
            MsgBox("Проверьте значения")
        End Try

        s1 = Form3.TextBox1.Text
        s2 = Form3.TextBox2.Text
        s3 = Form3.TextBox3.Text


        r = Form3.DialogResult

        Form3.Close()

        If r <> DialogResult.OK Then
            Exit Sub
        End If

        Dim c As New OleDbCommand
        c.Connection = _conn
        c.CommandText = "insert into Planets(PlanetName, M, R) values ('" & s1 & "','" & s2 & "','" & s3 & "')"
        c.ExecuteNonQuery()

        RefreshGrid()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim k As Integer
        Dim c As New OleDbCommand
        c.Connection = _conn
        k = DataGridView1.CurrentRow.Cells("Код").Value
        c.CommandText = "delete from Planets where Код = " & k
        c.ExecuteNonQuery()


        RefreshGrid()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim s1, s2, s3 As String
        Dim r As DialogResult
        Dim k As Integer
        k = DataGridView1.CurrentRow.Cells("Код").Value
        Form4.TextBox1.Text = DataGridView1.CurrentRow.Cells("PlanetName").Value
        Form4.TextBox2.Text = DataGridView1.CurrentRow.Cells("M").Value
        Form4.TextBox3.Text = DataGridView1.CurrentRow.Cells("R").Value

        Form4.ShowDialog()

        s1 = Form4.TextBox1.Text
        s2 = Form4.TextBox2.Text
        s3 = Form4.TextBox3.Text
        r = Form4.DialogResult

        Form4.Close()

        If r <> DialogResult.OK Then
            Exit Sub
        End If

        Dim _c As New OleDbCommand
        _c.Connection = _conn
        _c.CommandText = "update [Planets] set [PlanetName]='" & s1 & "', [M]='" & s2 & "', [R]= '" & s3 & "' where [Код]=" & k
        _c.ExecuteNonQuery()

        RefreshGrid()
    End Sub
End Class