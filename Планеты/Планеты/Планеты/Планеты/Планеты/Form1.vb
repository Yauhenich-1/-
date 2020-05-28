Imports System.Math
Public Class Form1

    Dim _id As Int16
    Dim _m As Double
    Dim _r As Double
    Dim _h As Double
    Dim _v1 As Double
    Dim _v2 As Double

    Const G As Double = 0.00000000006672


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fm2 As New Form2
        fm2.Show()

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "PlanetsDataSet.Planets". При необходимости она может быть перемещена или удалена.
        Me.PlanetsTableAdapter.Fill(Me.PlanetsDataSet.Planets)

    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged


    End Sub
    Private Sub GetPlanetInfo()

        Try
            _id = Int16.Parse(ComboBox1.SelectedValue)
            _h = Double.Parse(TextBox3.Text)

        Catch ex As Exception
            MsgBox("Обновлено")
        End Try
        Dim row = PlanetsTableAdapter.GetData().Item(_id - 1) 'так как нумерация строк с 0, а нумерация в таблице с 1
        _m = row.M
        _r = row.R
    End Sub

    Private Sub CalsVelocity()
        _v1 = Sqrt((G * _m) / (_r + _h))
        _v2 = _v1 * Sqrt(2)
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        GetPlanetInfo()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CalsVelocity()
        TextBox1.Text = _v1.ToString()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CalsVelocity()
        TextBox2.Text = _v2.ToString()
    End Sub
End Class
