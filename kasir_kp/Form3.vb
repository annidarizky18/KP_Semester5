Imports System.Data.Odbc
Public Class Form3
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
    Sub munculgrid()
        Call koneksi()
        Da = New OdbcDataAdapter("Select * from tb_kasir", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "tb_kasir")
        DG1.DataSource = Ds.Tables("tb_kasir")
        'responsive dg
        DG1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub
    Sub kondisiawal()
        ComboBox1.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        txtjr.Text = ""
        TextBox5.Text = ""
    End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call munculgrid()
        ComboBox1.Items.Add("RT-01")
        ComboBox1.Items.Add("RC-02")
        ComboBox1.Items.Add("RCKR-03")
        ComboBox1.Items.Add("RB-04")
        ComboBox1.Items.Add("RA-05")
        ComboBox1.Items.Add("RSI-06")
        ComboBox1.Items.Add("RW-07")
        ComboBox1.Items.Add("RZ-08")
        ComboBox1.Items.Add("RCA-09")

        'ComboBox1.Items.Add("Rotan Tebu")
        'ComboBox1.Items.Add("Rotan Cincin")
        'ComboBox1.Items.Add("Rotan Cakre")
        'ComboBox1.Items.Add("Rotan Boga")
        'ComboBox1.Items.Add("Rotan Aurense")
        'ComboBox1.Items.Add("Rotan Somi 1")
        'ComboBox1.Items.Add("Rotan Warbugii")
        'ComboBox1.Items.Add("Rotan Zipeli")
        'ComboBox1.Items.Add("Rotan Calamus Sp")
    End Sub

    Private Sub txtjr_TextChanged(sender As Object, e As EventArgs) Handles txtjr.TextChanged
        If txtjr.Text = "Rotan Tebu" Then
            TextBox1.Text = 20000
        ElseIf txtjr.Text = "Rotan Cincin" Then
            TextBox1.Text = 20000
        ElseIf txtjr.Text = "Rotan Cakre" Then
            TextBox1.Text = 25000
        ElseIf txtjr.Text = "Rotan Boga" Then
            TextBox1.Text = 20000
        ElseIf txtjr.Text = "Rotan Aurense" Then
            TextBox1.Text = 25000
        ElseIf txtjr.Text = "Rotan Somi 1" Then
            TextBox1.Text = 25000
        ElseIf txtjr.Text = "Rotan Warbugii" Then
            TextBox1.Text = 25000
        ElseIf txtjr.Text = "Rotan Zipeli" Then
            TextBox1.Text = 20000
        ElseIf txtjr.Text = "Rotan Warbugii" Then
            TextBox1.Text = 20000
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'ComboBox1.Items.Add("Rotan Tebu")
        'ComboBox1.Items.Add("Rotan Cincin")
        'ComboBox1.Items.Add("Rotan Cakre")
        'ComboBox1.Items.Add("Rotan Boga")
        'ComboBox1.Items.Add("Rotan Aurense")
        'ComboBox1.Items.Add("Rotan Somi 1")
        'ComboBox1.Items.Add("Rotan Warbugii")
        'ComboBox1.Items.Add("Rotan Zipeli")
        'ComboBox1.Items.Add("Rotan Calamus Sp")

        If ComboBox1.Text = "RT-01" Then
            txtjr.Text = "Rotan Tebu"
        ElseIf ComboBox1.Text = "RC-02" Then
            txtjr.Text = "Rotan Cincin"
        ElseIf ComboBox1.Text = "RCKR-03" Then
            txtjr.Text = "Rotan Cakre"
        ElseIf ComboBox1.Text = "RB-04" Then
            txtjr.Text = "Rotan Boga"
        ElseIf ComboBox1.Text = "RA-05" Then
            txtjr.Text = "Rotan Aurense"
        ElseIf ComboBox1.Text = "RSI-06" Then
            txtjr.Text = "Rotan Somi 1"
        ElseIf ComboBox1.Text = "RW-07" Then
            txtjr.Text = "Rotan Warbugii"
        ElseIf ComboBox1.Text = "RZ-08" Then
            txtjr.Text = "Rotan Zipeli"
        ElseIf ComboBox1.Text = "RCA-09" Then
            txtjr.Text = "Rotan Calamus Sp"
        End If

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox3.Text = Val(TextBox1.Text) * Val(TextBox2.Text)

            'TextBox5.Text = Val(TextBox4.Text) - Val(TextBox3.Text)
            'Call koneksi()
            'Cmd = New OdbcCommand("Select * from tb_register where id='" & TextBox1.Text & "'", Conn)
            'Rd = Cmd.ExecuteReader
            'Rd.Read()
            'If Not Rd.HasRows Then
            '    MsgBox("Id Admin tidak ada")
            'Else
            '    TextBox1.Text = Rd.Item("id")
            '    TextBox2.Text = Rd.Item("Nama")
            '    TextBox3.Text = Rd.Item("Password")
            'End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ComboBox1.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        txtjr.Text = ""
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = Chr(13) Then
            'TextBox3.Text = Val(TextBox1.Text) * Val(TextBox2.Text)

            TextBox5.Text = Val(TextBox4.Text) - Val(TextBox3.Text)
            'Call koneksi()
            'Cmd = New OdbcCommand("Select * from tb_register where id='" & TextBox1.Text & "'", Conn)
            'Rd = Cmd.ExecuteReader
            'Rd.Read()
            'If Not Rd.HasRows Then
            '    MsgBox("Id Admin tidak ada")
            'Else
            '    TextBox1.Text = Rd.Item("id")
            '    TextBox2.Text = Rd.Item("Nama")
            '    TextBox3.Text = Rd.Item("Password")
            'End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call munculgrid()
        Call kondisiawal()
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If txtjr.Text = "" Or ComboBox1.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Pastikan semua terisi!")
        Else
            Call koneksi()
            Dim InputData As String = "insert into tb_kasir values ('" & ComboBox1.Text & "','" & txtjr.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "')"
            Cmd = New OdbcCommand(InputData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Disimpan")
            Call kondisiawal()
            Call munculgrid()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class