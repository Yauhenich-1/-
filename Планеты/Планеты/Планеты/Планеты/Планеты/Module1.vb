Imports System.Data.OleDb
Module Module1
    Public _conn As OleDbConnection
    Public DA1 As OleDbDataAdapter


    Public Sub FillGrid(DataGridView1 As DataGridView, _cmd As String, TableName As String)
        Dim _c As New OleDbCommand
        _c.Connection = _conn
        _c.CommandText = _cmd

        Dim _ds As New DataSet
        Dim _da As New OleDbDataAdapter(_c)
        _da.Fill(_ds, TableName)
        DataGridView1.DataSource = _ds
        DataGridView1.DataMember = TableName


    End Sub

    Public Sub FillGridDA(DataGridView1 As DataGridView, _cmd As String, TableName As String, ByRef _da As OleDbDataAdapter)
        Dim _c As New OleDbCommand
        _c.Connection = _conn
        _c.CommandText = _cmd

        Dim _ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        _da = New OleDbDataAdapter(_c)
        _da.Fill(_ds, TableName)
        DataGridView1.DataSource = _ds
        DataGridView1.DataMember = TableName


    End Sub

End Module
