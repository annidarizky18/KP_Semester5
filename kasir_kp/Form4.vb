Imports System.Data.Odbc
Imports System.Drawing.Printing
Imports System.Windows.Forms.ComponentModel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Public Class Form4
    'Dim WithEvents PrintDoc As PrintDocument
    'Dim PrintPreviewDialog As PrintPreviewDialog

    'Private mRow As Integer = 0
    'Private newpage As Boolean = True

    ''Sub print()
    'Sub print()
    '    PrintDoc = New PrintDocument
    '    PrintPreviewDialog = New PrintPreviewDialog
    '    PrintDoc.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("PaperA4", 840, 1000)
    '    PrintPreviewDialog.Document = PrintDoc
    '    PrintDoc.DefaultPageSettings.Landscape = True
    '    PrintPreviewDialog.ShowDialog()
    'End Sub

    ''sub printdoc
    'Private Sub PrintDoc_BeginPrint(sender As Object, e As PrintEventArgs) Handles PrintDoc.BeginPrint
    '    mRow = 0
    '    newpage = True
    '    PrintPreviewDialog.PrintPreviewControl.StartPage = 0
    '    PrintPreviewDialog.PrintPreviewControl.Zoom = 1.0
    'End Sub

    ''function print
    'Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDoc.PrintPage
    '    Dim YAxis As Integer = 0

    '    ' sets it to show '...' for long text
    '    Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
    '    fmt.LineAlignment = StringAlignment.Center

    '    fmt.Trimming = StringTrimming.EllipsisCharacter
    '    Dim y As Int32 = 180 'Dim   y As Int32 = e.MarginBounds.Top
    '    Dim rc As Rectangle
    '    Dim x As Int32 = 50 ' Dim x As Int32
    '    Dim h As Int32 = 0
    '    Dim row As DataGridViewRow
    '    e.Graphics.DrawString("POSYANDU MEKAR", New Font("Century Gothic", 14, FontStyle.Bold), Brushes.Black, New Point(450, YAxis + 30))
    '    e.Graphics.DrawString("LAPORAN REKAPITULASI DATA ANAK", New Font("Times New Roman", 13, FontStyle.Bold), Brushes.Black, New Point(380, YAxis + 150))
    '    e.Graphics.DrawString("Jl. Kuningan 2. No.12A,Kec. Tambun, Kabupaten Bekasi, ", New Font("Century Gothic", 12, FontStyle.Regular), Brushes.Black, New Point(290, YAxis + 55))
    '    e.Graphics.DrawString("Telp : +62 812 8787  | Email: posyandumawar@gmail.com", New Font("Century Gothic", 12, FontStyle.Regular), Brushes.Black, New Point(265, YAxis + 75))
    '    e.Graphics.DrawImage(PictureBox1.Image, 100, 10)
    '    e.Graphics.DrawImage(PictureBox2.Image, 70, 100)
    '    e.Graphics.DrawImage(PictureBox3.Image, 70, 700)
    '    e.Graphics.DrawString("Tanggal Cetak Laporan: " & Date.Now.ToLongDateString, New Font("Century Gothic", 10, FontStyle.Bold), Brushes.Black, New Point(70, YAxis + 120))
    '    e.Graphics.DrawString("Mengetahui, Pemeriksa Posyandu", New Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, New Point(705, YAxis + 500))
    '    e.Graphics.DrawString("(Dea Anatasya)", New Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, New Point(750, YAxis + 600))
    '    e.Graphics.DrawString("Page1 OF 1", New Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, New Point(820, YAxis + 725))
    '    e.Graphics.DrawString("Bagian Admin", New Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, New Point(100, YAxis + 725))
    '    ''e.Graphics.DrawString("Client Name: " & FULLNAME.Text, New Font("Century Gothic", 10, FontStyle.Bold), Brushes.Black, New Point(100, YAxis + 50))
    '    ''e.Graphics.DrawString("Total Amount: " & LabelTotalPayment.Text, New Font("Century Gothic", 10, FontStyle.Bold), Brushes.Black, New Point(100, YAxis + 100))

    '    ' print the header text for a new page
    '    '   use a grey bg just like the control
    '    If newpage Then
    '        row = dganak.Rows(mRow)
    '        x = 100 'x = e.MarginBounds.Left
    '        For Each cell As DataGridViewCell In row.Cells
    '            ' since we are printing the control's view,
    '            ' skip invidible columns
    '            If cell.Visible Then
    '                rc = New Rectangle(x, y, cell.Size.Width, cell.Size.Height)

    '                e.Graphics.FillRectangle(Brushes.LightGray, rc)
    '                e.Graphics.DrawRectangle(Pens.Black, rc)

    '                ' reused in the data pront - should be a function
    '                Select Case dganak.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
    '                    Case DataGridViewContentAlignment.BottomRight,
    '                         DataGridViewContentAlignment.MiddleRight
    '                        fmt.Alignment = StringAlignment.Far
    '                        rc.Offset(-1, 0)
    '                    Case DataGridViewContentAlignment.BottomCenter,
    '                        DataGridViewContentAlignment.MiddleCenter
    '                        fmt.Alignment = StringAlignment.Center
    '                    Case Else

    '                        fmt.Alignment = StringAlignment.Center
    '                        rc.Offset(2, 0)
    '                End Select
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)
    '                e.Graphics.DrawString(dganak.Columns(cell.ColumnIndex).HeaderText, dganak.Font, Brushes.Black, rc, fmt)




    '                x += rc.Width
    '                h = Math.Max(h, rc.Height)
    '            End If
    '        Next
    '        y += h

    '    End If
    '    newpage = False

    '    ' now print the data for each row
    '    Dim thisNDX As Int32
    '    For thisNDX = mRow To dganak.RowCount - 1
    '        ' no need to try to print the new row
    '        If dganak.Rows(thisNDX).IsNewRow Then Exit For

    '        row = dganak.Rows(thisNDX)
    '        x = 100 'x = e.MarginBounds.Left
    '        h = 0

    '        ' reset X for data
    '        x = 100 'e.MarginBounds.Left

    '        ' print the data
    '        For Each cell As DataGridViewCell In row.Cells
    '            If cell.Visible Then
    '                rc = New Rectangle(x, y, cell.Size.Width, cell.Size.Height)

    '                ' SAMPLE CODE: How To 
    '                ' up a RowPrePaint rule
    '                'If Convert.ToDecimal(row.Cells(5).Value) < 9.99 Then
    '                '    Using br As New SolidBrush(Color.MistyRose)
    '                '        e.Graphics.FillRectangle(br, rc)
    '                '    End Using
    '                'End If

    '                e.Graphics.DrawRectangle(Pens.Black, rc)

    '                Select Case DG1.Columns(cell.ColumnIndex).DefaultCellStyle.Alignment
    '                    Case DataGridViewContentAlignment.BottomRight,
    '                         DataGridViewContentAlignment.MiddleRight
    '                        fmt.Alignment = StringAlignment.Far
    '                        rc.Offset(-1, 0)
    '                    Case DataGridViewContentAlignment.BottomCenter,
    '                        DataGridViewContentAlignment.MiddleCenter
    '                        fmt.Alignment = StringAlignment.Center
    '                    Case Else
    '                        fmt.Alignment = StringAlignment.Center
    '                        rc.Offset(2, 0)
    '                End Select

    '                e.Graphics.DrawString(cell.FormattedValue.ToString(),
    '                                      dganak.Font, Brushes.Black, rc, fmt)

    '                x += rc.Width
    '                h = Math.Max(h, rc.Height)
    '            End If

    '        Next
    '        y += h
    '        ' next row to print
    '        mRow = thisNDX + 1

    '        If y + h > e.MarginBounds.Bottom Then 'e.MarginBounds.Bottom
    '            e.HasMorePages = True
    '            '  mRow -= 1   \causes last row To rePrint On Next page
    '            newpage = True
    '            Return
    '        End If

    '    Next

    'End Sub


    Dim Cmd As OdbcCommand
    Dim Conn As OdbcConnection
    Dim Ds As DataSet
    Dim Da As OdbcDataAdapter
    Dim Dr As OdbcDataReader
    Dim Rd As OdbcDataReader
    Dim MyDB As String

    Sub koneksi()
        MyDB = "Driver={Mysql ODBC 3.51 Driver};Database=db_kasir;server=localhost;uid=root"
        Conn = New OdbcConnection(MyDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Sub munculgrid()
        Call koneksi()
        Da = New OdbcDataAdapter("Select * from tb_stok", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "tb_stok")
        DG1.DataSource = Ds.Tables("tb_stok")
        'responsive dg
        DG1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub
    Sub kondisiawal()
        ComboBox1.Text = ""
        txtjr.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        If ComboBox1.Text = "" Or txtjr.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Pastikan semua terisi!")
        Else
            Call koneksi()
            Dim InputData As String = "insert into tb_stok values ('" & ComboBox1.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "')"
            Cmd = New OdbcCommand(InputData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Disimpan")
            Call kondisiawal()
            Call munculgrid()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

        Call munculgrid()
    End Sub

    Private Sub DG1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG1.CellClick
        Dim i As Integer
        i = DG1.CurrentRow.Index

        ComboBox1.Text = DG1.Item(0, i).Value
        txtjr.Text = DG1.Item(1, i).Value
        TextBox1.Text = DG1.Item(2, i).Value
        TextBox2.Text = DG1.Item(3, i).Value
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call munculgrid()
        Call kondisiawal()

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
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call koneksi()
        Dim hapus As String = "delete from tb_stok where Kode_Rotan='" & ComboBox1.Text & "'"
        Cmd = New OdbcCommand(hapus, Conn)
        Cmd.ExecuteNonQuery()
        MsgBox("Data Berhasil di Hapus", MsgBoxStyle.Information, "INFORMATION")
        Call kondisiawal()
        Call munculgrid()
    End Sub

    Private Sub txtCari_TextChanged(sender As Object, e As EventArgs) Handles txtCari.TextChanged
        Call koneksi()
        Cmd = New OdbcCommand("select * from tb_stok where Jenis_Rotan like '%" & txtCari.Text & "%'", Conn)
        Dr = Cmd.ExecuteReader
        Dr.Read()
        If Dr.HasRows Then
            Call koneksi()
            Da = New OdbcDataAdapter("select * from tb_stok where Jenis_Rotan like '%" & txtCari.Text & "%'", Conn)
            Ds = New DataSet
            Da.Fill(Ds)
            DG1.DataSource = Ds.Tables(0)
        Else
            MsgBox("Data tidak ditemukan")
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Or txtjr.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Pastikan semua terisi!")
        Else
            Call koneksi()
            Dim edit As String = "update tb_stok set Jenis_Rotan='" & txtjr.Text & "',Harga_Rotan='" & TextBox1.Text & "',Jumlah_Rotan='" & TextBox2.Text & "' where Kode_Rotan='" & ComboBox1.Text & "'"
            Cmd = New OdbcCommand(edit, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Edit Data Berhasil", MsgBoxStyle.Information, "INFORMATION")
            Call kondisiawal()
            Call munculgrid()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call munculgrid()
        Call kondisiawal()
        Form2.Show()
        Me.Hide()
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
        ElseIf txtjr.Text = "Rotan Calamus Sp" Then
            TextBox1.Text = 20000
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
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

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        'Call Print()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If txtjr.Text = "" Or ComboBox1.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Pastikan semua terisi!")
        Else
            Call koneksi()
            Dim InputData As String = "insert into tb_stok values ('" & ComboBox1.Text & "','" & txtjr.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "')"
            Cmd = New OdbcCommand(InputData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Disimpan")
            Call kondisiawal()
            Call munculgrid()
        End If
    End Sub
End Class