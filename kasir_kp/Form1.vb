Imports System.Data.Odbc
Imports System.Windows

Public Class Form1
    Dim Cmd As OdbcCommand
    Dim Conn As OdbcConnection
    Dim Ds As DataSet
    Dim Da As OdbcDataAdapter
    Dim Rd As OdbcDataReader
    Dim MyDB As String

    Sub koneksi()
        MyDB = "Driver={Mysql ODBC 3.51 Driver};Database=db_kasir;server=localhost;uid=root"
        Conn = New OdbcConnection(MyDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call koneksi()
        Cmd = New OdbcCommand("Select * from tb_admin where username = '" & TextBox1.Text & "' and password ='" & TextBox2.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            Form2.Show()
            Me.Hide()
        Else
            MsgBox("USERNAME atau PASSWORD Salah!!")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
    End Sub
End Class
